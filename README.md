# Vellum — PDF → DOCX Converter

> *Vellum* (from Latin **_vitulinum_**, "of the calf", via Old French
> *velin*) was the smooth, durable writing material medieval scribes
> prepared from calf skin: cleaned, scraped flat, and stretched until
> it could carry a pen. Whole libraries lived on it, long before paper.
> This converter borrows the name as a small homage — it takes a PDF
> (itself a flattened, fixed-position simulation of the printed page)
> and lifts the text, tables, and structure back onto a fresh surface
> where they can be edited, rewritten, and bound into something new.

A Python tool that converts PDFs to editable Word documents, built around
[`pdf2docx`](https://github.com/dothinking/pdf2docx) for text-based PDFs and a
custom [Tesseract](https://github.com/tesseract-ocr/tesseract)-based pipeline
for scanned PDFs. Ships with three frontends sharing the same conversion
core:

- **Command line** (`src/pdf_to_docx.py`) — scriptable, batch-friendly.
- **Desktop GUI** (`src/gui.py`) — Tkinter, bilingual (Slovak default + English).
- **Web app** (`src/app.py`) — Flask, dockerized, with per-IP rate limiting and an access log.

## Project layout

```
.
├── README.md
├── LICENSE
├── requirements.txt
├── src/                          # Application code
│   ├── _version.py               # Single source of truth for version + author
│   ├── pdf_to_docx.py            # Conversion core (also a CLI)
│   ├── gui.py                    # Tkinter desktop GUI
│   ├── app.py                    # Flask web app
│   └── templates/
│       └── index.html
├── docker/                       # Web-app Docker setup
│   ├── Dockerfile
│   ├── docker-compose.yml
│   └── .dockerignore
├── build-windows/                # Cross-build a Windows .exe via Wine
│   ├── Dockerfile
│   ├── build.sh
│   └── requirements-windows.txt
└── tests/data/                   # Public synthetic sample PDFs
    ├── README.md
    ├── generate_samples.py       # Regenerator (needs Pillow)
    ├── text_with_tables.pdf      # Exercises lattice table path
    └── scanned_text.pdf          # Exercises OCR fallback path
```

## Quick start (local)

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Command line

```bash
# Text PDF → DOCX
python src/pdf_to_docx.py input.pdf output.docx

# Scanned PDF → DOCX (uses local tesseract)
python src/pdf_to_docx.py scan.pdf scan.docx --ocr --ocr-language eng
# Multi-language OCR
python src/pdf_to_docx.py scan.pdf scan.docx --ocr --ocr-language slk+eng
```

By default the converter uses **lattice-only** table detection (only tables
with visible borders). Stream-table mode is opt-in via `--stream-tables` —
it tends to invent fake tables out of label/value blocks for the kind of
school-curriculum PDFs this tool was originally built for.

### Desktop GUI

```bash
python src/gui.py
```

A Tkinter window opens. The OCR row only appears when `tesseract` is
detected on `PATH`. Conversion runs on a background thread, so the UI stays
responsive; expand the log panel to watch progress.

### Web app

```bash
python src/app.py
```

Then open <http://localhost:4567>. The web app uses
[waitress](https://github.com/Pylons/waitress) as the WSGI server when
available, falling back to Flask's dev server otherwise.

## Docker (web app)

```bash
docker compose -f docker/docker-compose.yml up --build
```

Then open <http://localhost:5001>.

The container ships with `tesseract-ocr-eng` and `tesseract-ocr-slk` so the
OCR option appears in the form. To add more languages, edit
`docker/Dockerfile` and add `tesseract-ocr-XXX` packages — see Debian's
[tesseract package list](https://packages.debian.org/search?keywords=tesseract-ocr).

### Persistent access log

The container's access log lives at `/app/logs/access.log` inside the
container, mounted from `./logs/` on the host. Tail it with:

```bash
tail -f logs/access.log
```

Sample line:

```
2026-04-07 20:30:15 ip=192.168.1.5 START file='input.pdf' ocr=False ocr_lang=- no_stream=True ua='Mozilla/5.0 ...'
2026-04-07 20:30:17 ip=192.168.1.5 OK    file='input.pdf' in=343138B out=94770B
```

The file rotates automatically (5 × 5 MB).

### Behind a reverse proxy

If you put nginx, Caddy, or a Cloudflare tunnel in front of the container,
set `TRUST_PROXY=1` in the compose `environment:` block. Without it, the
access log records the proxy's IP instead of the real client's. Only enable
when actually behind a trusted proxy — otherwise an attacker could spoof
the `X-Forwarded-For` header.

### Tunable environment variables

| Var | Default | What it does |
|---|---|---|
| `PORT` | `4567` (local), `5001` (compose) | Server listen port |
| `TRUST_PROXY` | unset | Set to `1` to honor `X-Forwarded-For` |
| `ACCESS_LOG_FILE` | `logs/access.log` | Path to the access log file |
| `RATE_LIMIT_REQUESTS` | `30` | Max requests per window per IP |
| `RATE_LIMIT_WINDOW` | `60` | Window length in seconds |

## Cross-compile Windows .exe

There are two ways to build the Windows distribution. Use whichever
matches your situation.

### Recommended: GitHub Actions (real Windows runner)

A workflow at `.github/workflows/build-windows.yml` builds the Windows
`.exe` on a real `windows-latest` runner. No Wine, no Rosetta, no
container — it just runs PyInstaller natively on Windows. Trigger it by
**cutting a release tag**:

```bash
make git-release
```

`make git-release` reads the version from `src/_version.py`, refuses if
the working tree is dirty or the tag already exists, creates an
annotated `vX.Y.Z` tag, and pushes it to `origin`. The push triggers the
workflow, which builds the `.exe`, zips `dist/PDF-to-DOCX/` as
`Vellum-PDF-to-DOCX-vX.Y.Z-windows.zip`, and attaches it to a fresh
GitHub Release.

You can also trigger the workflow manually from the **Actions** tab in
the GitHub UI (the `workflow_dispatch` button) — useful for smoke
testing the build without cutting a real release. In that case the zip
is uploaded as a workflow artifact (downloadable from the run page) but
no Release is created.

### Local: Wine + Docker (offline, slower)

If you don't have a GitHub repo or want to build offline:

```bash
./build-windows/build.sh
```

This builds a Wine + Windows-Python image, runs PyInstaller against
`src/gui.py`, and emits `dist/PDF-to-DOCX/` containing `PDF-to-DOCX.exe`
plus its runtime files. The whole `PDF-to-DOCX/` folder is the
distributable — zip it and send to Windows users.

The build uses **`--onedir --noupx`**: the resulting `.exe` is bigger
than a `--onefile --upx` build but doesn't trigger Windows Defender
false positives, doesn't suffer from `--onefile`'s temp-extract startup
penalty, and inspecting the contents in Explorer just works.

Notes:

- Tesseract is **not** bundled into the Windows distribution. Windows
  users who want OCR have to install Tesseract for Windows separately
  (the [UB Mannheim build](https://github.com/UB-Mannheim/tesseract/wiki)
  is the standard one). The .exe detects it at startup and hides the
  OCR option if it isn't on `PATH`.
- The base image (`tobix/pywine:3.12`) is only published for
  `linux/amd64`, so on Apple Silicon hosts the build runs through
  Rosetta emulation. The first build is slow; subsequent ones reuse
  the layer cache.

## Security posture (web app)

The Flask app has been audited for the OWASP-style basics:

- **Strict response security headers** (CSP, X-Frame-Options DENY,
  X-Content-Type-Options nosniff, Referrer-Policy no-referrer,
  Permissions-Policy lockdown).
- **No version disclosure** — the `Server` header is overridden so we
  don't leak Werkzeug + Python versions.
- **Per-IP rate limiting** on `/convert` (default 30 req / 60 s).
- **Input validation**: magic-byte PDF sniff, filename length cap,
  OCR-language whitelist with per-component validation for `slk+eng`
  multi-language strings.
- **Cache-Control: no-store** on conversion responses so the DOCX
  isn't cached by browsers or intermediate proxies.
- **Production WSGI** (`waitress`) instead of Flask's dev server.
- **No debug mode** (`debug=True` is never set — debug mode is RCE).

Known limitations (out of scope for an in-process minimal converter):

- **No CSRF token** on `/convert`. Impact is low because the endpoint
  has no auth and produces a one-shot file response, but if you put it
  behind authentication you should add CSRF protection.
- **No conversion timeout**. Python can't kill running threads cleanly;
  use your reverse proxy's `proxy_read_timeout` (or equivalent) as the
  practical mitigation.
- **No PDF parser sandboxing**. PyMuPDF and pdf2docx have had CVE
  history. If exposing to untrusted internet, run the container with
  reduced capabilities (read-only root FS, dropped capabilities) or
  behind a process-isolation layer.
- **No authentication**. Recommend a reverse proxy with HTTP basic auth
  if exposing publicly.

## License

[MIT](LICENSE) © 2026 Pavol Calfa.
