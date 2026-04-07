"""Minimal Flask web GUI for the PDF to DOCX converter.

Wraps the conversion functions from `pdf_to_docx.py` in a single-page upload
form. Uploaded PDFs are converted in a temporary directory and the result is
streamed back to the browser as a download — nothing is persisted on disk
beyond the access log.

Two loggers:
* ``app.logger`` (Flask default) — application diagnostics, errors,
  exceptions. Writes to stderr (captured by ``docker compose logs``).
* ``access_log`` — dedicated request log. Writes ONLY to a rotating
  file (default ``./logs/access.log``, override via ``ACCESS_LOG_FILE``).
  Each conversion request is recorded with timestamp, client IP, file
  name + size, OCR settings, and outcome.

To get the real client IP when running behind a reverse proxy
(nginx, Caddy, Cloudflare tunnel, etc.), set ``TRUST_PROXY=1`` in the
environment so ``ProxyFix`` parses the ``X-Forwarded-For`` header.
Without ``TRUST_PROXY``, the logged IP will be the proxy's address.

Security posture (see also: web security audit notes in the README):

* Strict response security headers (CSP, X-Frame-Options,
  X-Content-Type-Options, Referrer-Policy, Permissions-Policy).
* Server banner stripped (no Werkzeug version disclosure).
* Per-IP rate limit on /convert (default 30 requests / 60 seconds).
* Filename length cap, magic-byte PDF sniff, OCR-language whitelist
  (multi-language `slk+eng` syntax allowed but each part validated).
* Cache-Control: no-store on conversion responses (so the DOCX isn't
  cached by browsers/intermediate proxies).
* Production WSGI server (waitress) when available, falls back to the
  Flask dev server with a warning otherwise.

Known limitations (out of scope for this in-process app):

* No CSRF token on /convert. Impact is low because the endpoint has no
  auth and produces a one-shot file response, but if you add auth in
  front you should add CSRF protection too.
* No conversion timeout (Python can't kill running threads cleanly).
  Use your reverse proxy's read timeout as the practical mitigation.
* No sandboxing of the PDF parser. PyMuPDF/pdf2docx have had CVE
  history. If exposing to untrusted internet, run the container with
  reduced capabilities (read-only root FS, dropped capabilities, etc.)
  or behind a process-isolation layer.

Run locally:
    python app.py

Run in Docker:
    docker compose up --build
"""
from __future__ import annotations

import logging
import logging.handlers
import os
import tempfile
import time
from collections import defaultdict, deque
from io import BytesIO
from pathlib import Path
from threading import Lock

from flask import Flask, abort, make_response, render_template, request, send_file
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

from _version import __author__, __codename__, __release_date__, __version__
from pdf_to_docx import (
    convert_pdf,
    fix_bullet_fonts,
    fix_first_table_header,
    list_tesseract_languages,
    ocr_to_docx,
)

# General app logging → stderr (captured by docker logs).
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)


# ----- access log: dedicated logger writing to a rotating file --------

ACCESS_LOG_FILE = Path(os.environ.get("ACCESS_LOG_FILE", "logs/access.log"))
ACCESS_LOG_MAX_BYTES = 5 * 1024 * 1024  # 5 MB per file
ACCESS_LOG_BACKUPS = 5                  # keep 5 rotated files

ACCESS_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
access_log = logging.getLogger("access")
access_log.setLevel(logging.INFO)
# Don't propagate to the root logger — we want this in the file ONLY,
# not also dumped to stderr alongside the general app logs.
access_log.propagate = False
_access_handler = logging.handlers.RotatingFileHandler(
    ACCESS_LOG_FILE,
    maxBytes=ACCESS_LOG_MAX_BYTES,
    backupCount=ACCESS_LOG_BACKUPS,
    encoding="utf-8",
)
_access_handler.setFormatter(
    logging.Formatter("%(asctime)s %(message)s")
)
access_log.addHandler(_access_handler)


class _RateLimiter:
    """Tiny in-memory sliding-window rate limiter, thread-safe.

    Used to put a per-IP cap on /convert requests so a single client
    can't trivially DoS the server. In-memory means the counter is
    process-local and resets on restart — fine for a single-instance
    converter; if you scale horizontally use Flask-Limiter with redis.
    """

    def __init__(self, max_requests: int, window_seconds: float) -> None:
        self._max = max_requests
        self._window = window_seconds
        self._hits: dict[str, deque[float]] = defaultdict(deque)
        self._lock = Lock()

    def check(self, key: str) -> bool:
        """Return True if the call is allowed, False if it should be
        rejected (HTTP 429)."""
        now = time.monotonic()
        cutoff = now - self._window
        with self._lock:
            q = self._hits[key]
            while q and q[0] < cutoff:
                q.popleft()
            if len(q) >= self._max:
                return False
            q.append(now)
            return True


# 30 requests per 60 seconds per client IP. Tunable via env vars in
# case you want to tighten or relax it without editing code.
_RATE_LIMIT = _RateLimiter(
    max_requests=int(os.environ.get("RATE_LIMIT_REQUESTS", "30")),
    window_seconds=float(os.environ.get("RATE_LIMIT_WINDOW", "60")),
)


app = Flask(__name__)
# Reject anything obviously too large before we hit pdf2docx.
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# Hard cap on filename length even after secure_filename. secure_filename
# strips dangerous characters but does not enforce a length limit, so a
# malicious client could submit a 100KB filename and waste log space /
# hit OS path limits. 200 chars leaves room for the .docx suffix and
# any temp-dir prefix while still being well within filesystem limits.
_MAX_FILENAME_LEN = 200


@app.after_request
def _apply_security_headers(response):
    """Apply baseline security headers to every response.

    * X-Content-Type-Options: nosniff — prevents MIME sniffing attacks.
    * X-Frame-Options: DENY — prevents clickjacking via iframe embed.
    * Referrer-Policy: no-referrer — don't leak the converter URL.
    * Content-Security-Policy: tight default-src 'self'. Inline script
      and style are allowed because the form template uses both;
      tightening that further is a future hardening opportunity.
    * Permissions-Policy: deny everything we don't need (camera, mic,
      geolocation, payment, etc.).
    * Strip the Server header so we don't disclose Werkzeug + Python
      versions to potential attackers.
    """
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline'; "
        "img-src 'self' data:; "
        "form-action 'self'; "
        "frame-ancestors 'none'; "
        "base-uri 'self'"
    )
    response.headers["Permissions-Policy"] = (
        "camera=(), microphone=(), geolocation=(), payment=(), "
        "usb=(), interest-cohort=()"
    )
    # Strip the version-disclosing Server header.
    response.headers["Server"] = "pdf2docx-web"
    # Expose our own version (not Werkzeug's). Useful for ops/observability
    # without revealing the underlying stack.
    response.headers["X-App-Version"] = __version__
    return response

# Behind a reverse proxy, request.remote_addr is the proxy's IP, not the
# real client. ProxyFix parses X-Forwarded-For (and friends) so the
# correct IP propagates. Only enable when actually behind a trusted proxy
# — otherwise an attacker could spoof the X-Forwarded-For header.
if os.environ.get("TRUST_PROXY", "").lower() in ("1", "true", "yes"):
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    app.logger.info("ProxyFix enabled (TRUST_PROXY=%s)",
                    os.environ.get("TRUST_PROXY"))

# Log app codename + version + author at startup so docker logs /
# journald show exactly what's running without anyone needing to grep
# the source.
app.logger.info(
    "%s v%s (%s) — by %s",
    __codename__, __version__, __release_date__, __author__,
)

# Detect tesseract once at startup. None = not installed → no OCR
# controls in the form. A list = available language codes for the
# OCR language picker.
TESSERACT_LANGS: list[str] | None = list_tesseract_languages()
if TESSERACT_LANGS:
    app.logger.info("tesseract detected, languages: %s", TESSERACT_LANGS)
else:
    app.logger.info("tesseract not detected, OCR option disabled")


def _client_ip() -> str:
    """Return the best-known client IP for the current request.

    With TRUST_PROXY enabled, this is the parsed-out X-Forwarded-For
    address. Without it, this is the immediate TCP peer (which is the
    reverse proxy when there is one).
    """
    return request.remote_addr or "unknown"


def _looks_like_pdf(file_storage) -> bool:
    """Sniff the first bytes of the upload — real PDFs start with '%PDF-'."""
    head = file_storage.stream.read(5)
    file_storage.stream.seek(0)
    return head == b"%PDF-"


@app.get("/")
def index():
    return render_template(
        "index.html",
        tesseract_langs=TESSERACT_LANGS or [],
        app_codename=__codename__,
        app_version=__version__,
        app_release_date=__release_date__,
        app_author=__author__,
    )


def _validate_ocr_language(raw: str) -> str | None:
    """Return the validated OCR language string, or None if invalid.

    Accepts ``slk+eng``-style multi-language strings; each component
    must be present in TESSERACT_LANGS. This both keeps the validator
    strict (no shell-injection-like values reach the tesseract argv)
    and gives the user the multi-language feature that single-string
    matching would deny.
    """
    if not raw:
        return None
    parts = raw.split("+")
    if not parts or len(parts) > 4:
        # 4 is a generous cap — tesseract works fine with 2-3 langs;
        # an unbounded list would be a sign of abuse or fuzzing.
        return None
    allowed = TESSERACT_LANGS or []
    for p in parts:
        if not p or p not in allowed:
            return None
    return "+".join(parts)


@app.post("/convert")
def convert():
    client = _client_ip()
    user_agent = (request.headers.get("User-Agent") or "?")[:80]

    # Per-IP rate limit. Reject before doing any expensive work.
    if not _RATE_LIMIT.check(client):
        access_log.info("ip=%s REJECT reason=rate-limit ua=%r", client, user_agent)
        abort(429, "Too many requests — slow down and try again in a minute.")

    upload = request.files.get("pdf")
    if upload is None or upload.filename == "":
        access_log.info("ip=%s REJECT reason=no-file ua=%r", client, user_agent)
        abort(400, "No file uploaded.")
    if not _looks_like_pdf(upload):
        access_log.info(
            "ip=%s REJECT reason=not-pdf file=%r ua=%r",
            client, upload.filename, user_agent,
        )
        abort(400, "Uploaded file is not a PDF.")

    no_stream = request.form.get("no_stream_tables") == "on"
    ocr_enabled = request.form.get("ocr") == "on"
    ocr_lang_raw = request.form.get("ocr_language", "eng")

    # Defensive: if the client somehow submits ocr=on but the server
    # has no tesseract installed, fall back to the regular path with a
    # warning rather than crashing.
    if ocr_enabled and not TESSERACT_LANGS:
        app.logger.warning(
            "OCR requested but tesseract is not installed; "
            "falling back to non-OCR conversion (ip=%s)", client,
        )
        ocr_enabled = False

    ocr_lang: str | None = None
    if ocr_enabled:
        ocr_lang = _validate_ocr_language(ocr_lang_raw)
        if ocr_lang is None:
            # Reject unknown language codes — protects against shell-
            # injection-like values being passed to tesseract.
            access_log.info(
                "ip=%s REJECT reason=bad-lang ocr_lang=%r",
                client, ocr_lang_raw,
            )
            abort(400, f"Unknown OCR language: {ocr_lang_raw!r}")

    safe_name = secure_filename(upload.filename) or "input.pdf"
    # Cap filename length to prevent OS path-limit / log-flood abuse.
    if len(safe_name) > _MAX_FILENAME_LEN:
        stem_part = Path(safe_name).stem[: _MAX_FILENAME_LEN - 4]
        safe_name = f"{stem_part}.pdf"
    stem = Path(safe_name).stem or "output"

    access_log.info(
        "ip=%s START file=%r ocr=%s ocr_lang=%s no_stream=%s ua=%r",
        client, safe_name, ocr_enabled, ocr_lang if ocr_enabled else "-",
        no_stream, user_agent,
    )

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        pdf_path = tmp_path / safe_name
        docx_path = tmp_path / f"{stem}.docx"
        upload.save(pdf_path)
        in_size = pdf_path.stat().st_size

        try:
            if ocr_enabled:
                # OCR path: bypass pdf2docx, build DOCX directly from
                # tesseract output (matches gui.py and CLI behavior).
                ocr_to_docx(pdf_path, docx_path, language=ocr_lang)
            else:
                convert_pdf(
                    pdf_path, docx_path, parse_stream_table=not no_stream,
                )
                fix_bullet_fonts(docx_path)
                fix_first_table_header(pdf_path, docx_path)
        except Exception as exc:
            app.logger.exception(
                "convert FAILED: ip=%s file=%r in_size=%d",
                client, safe_name, in_size,
            )
            access_log.info(
                "ip=%s FAIL file=%r in=%dB error=%s",
                client, safe_name, in_size, type(exc).__name__,
            )
            abort(500, "Conversion failed — see server logs for details.")

        # Read the file into memory before the temp dir is cleaned up so
        # send_file can stream it after the `with` block exits.
        out_size = docx_path.stat().st_size
        buf = BytesIO(docx_path.read_bytes())

    access_log.info(
        "ip=%s OK file=%r in=%dB out=%dB",
        client, safe_name, in_size, out_size,
    )

    buf.seek(0)
    response = make_response(send_file(
        buf,
        as_attachment=True,
        download_name=f"{stem}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ))
    # Conversion output may contain sensitive content from the source
    # PDF — prevent caching by browsers and intermediate proxies.
    response.headers["Cache-Control"] = "no-store, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


if __name__ == "__main__":
    # Default to 4567; port 5000 is squatted by macOS ControlCenter (AirPlay Receiver).
    port = int(os.environ.get("PORT", "4567"))

    # Use waitress (production WSGI) if available; fall back to the Flask
    # dev server with a clear warning otherwise. Flask's dev server is
    # single-threaded, prints "do not use in production" on every start,
    # and lacks graceful shutdown — fine for local prototyping, not for
    # any container that's reachable from the network.
    try:
        from waitress import serve
        app.logger.info("starting waitress on 0.0.0.0:%d", port)
        serve(app, host="0.0.0.0", port=port, ident="pdf2docx-web")
    except ImportError:
        app.logger.warning(
            "waitress not installed — falling back to Flask dev server. "
            "Run `pip install waitress` for production use."
        )
        app.run(host="0.0.0.0", port=port)
