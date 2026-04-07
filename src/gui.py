"""Minimal Tkinter desktop GUI for the PDF to DOCX converter.

Bilingual (Slovak default + English) with a language selector in the
top-right corner. Wraps the conversion functions from pdf_to_docx.py in a
simple file-picker window. Designed to be packaged as a standalone Windows
executable via PyInstaller (see build-windows/).

Run locally:
    python gui.py
"""
from __future__ import annotations

import contextlib
import logging
import re
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from _tesseract_langs import TESSERACT_LANG_NAMES
from _version import __author__, __codename__, __release_date__, __version__
from pdf_to_docx import (
    convert_pdf,
    fix_bullet_fonts,
    fix_first_table_header,
    list_tesseract_languages,
    ocr_to_docx,
)


# ANSI color escape codes — pdf2docx prints things like
# `[INFO] [1;36m[1/4] Opening document...[0m`. Tk's Text widget can't
# render ANSI, so we strip the codes before writing to the log widget.
_ANSI_ESCAPE_RE = re.compile(r"\x1b\[[0-9;]*[A-Za-z]")


class _LogWriter:
    """File-like object that redirects writes to a Tk Text widget.

    Used to capture stdout/stderr from the conversion worker thread and
    stream it into the GUI's log widget. All actual widget writes are
    marshalled back to the main thread via ``Tk.after(0, …)`` because
    Tk widgets are not thread-safe.

    Buffers partial lines so we batch widget updates by full line —
    avoids hammering the event loop with one ``after()`` call per
    character of output.
    """

    def __init__(self, app: "ConverterApp") -> None:
        self._app = app
        self._buffer = ""

    def write(self, s: str) -> int:
        if not s:
            return 0
        self._buffer += _ANSI_ESCAPE_RE.sub("", s)
        if "\n" in self._buffer:
            *complete_lines, tail = self._buffer.split("\n")
            self._buffer = tail
            chunk = "\n".join(complete_lines) + "\n"
            self._app.after(0, self._app._log_append, chunk)
        return len(s)

    def flush(self) -> None:
        if self._buffer:
            self._app.after(0, self._app._log_append, self._buffer)
            self._buffer = ""


class _TkLogHandler(logging.Handler):
    """logging.Handler that emits records into a Tk Text widget.

    Required because pdf2docx calls ``logging.basicConfig(...)`` at
    module import, which installs a StreamHandler bound to the *original*
    sys.stderr. Once installed, that handler holds a direct reference to
    the original stream — ``contextlib.redirect_stderr`` later has no
    effect on it. So to capture pdf2docx's ``[INFO]`` progress lines we
    have to install our own logging handler, not (only) redirect the
    streams.

    Uses the same format as pdf2docx's basicConfig call so the lines
    look identical to running pdf2docx from a terminal.

    Thread-safety: ``emit`` is called from the worker thread (because
    pdf2docx logs from there); we marshal back to the main thread via
    ``Tk.after(0, …)``.
    """

    def __init__(self, app: "ConverterApp") -> None:
        super().__init__()
        self._app = app
        self.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
        except Exception:  # noqa: BLE001
            return
        # Strip ANSI color codes pdf2docx puts inside log messages.
        msg = _ANSI_ESCAPE_RE.sub("", msg)
        self._app.after(0, self._app._log_append, msg + "\n")


def _enable_high_dpi() -> None:
    """Enable per-monitor DPI awareness on Windows.

    Without this, Windows treats Tkinter apps as DPI-unaware and
    bitmap-stretches them to the user's scaling factor — fonts and widgets
    look blurry on high-DPI displays. After SetProcessDpiAwareness(2)
    Windows stops stretching and Tk renders at native resolution.

    Must be called BEFORE the first Tk root is created. No-op on
    non-Windows platforms.
    """
    if sys.platform != "win32":
        return
    try:
        from ctypes import windll  # type: ignore[attr-defined]
        # PROCESS_PER_MONITOR_DPI_AWARE = 2 (Windows 8.1+)
        windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError, ImportError):
        # Fallback for older Windows: system DPI aware (not per-monitor)
        try:
            windll.user32.SetProcessDPIAware()  # type: ignore[name-defined]
        except Exception:
            pass


# Translation tables. Keys are language-agnostic identifiers; values are
# user-visible strings. Add a new language by adding another entry here and
# extending LANGUAGE_NAMES below.
TRANSLATIONS: dict[str, dict[str, str]] = {
    "sk": {
        "window_title": "Konvertor PDF na DOCX",
        "heading": "Konvertor PDF na DOCX",
        "subtitle": "Skonvertuje PDF na upraviteľný Word dokument.",
        "language_label": "Jazyk:",
        "pdf_file_label": "PDF súbor:",
        "browse_button": "Prehľadávať…",
        "no_stream_checkbox": "Detekovať len ohraničené tabuľky (odporúčané)",
        "ocr_checkbox": "Použiť OCR (pre skenované PDF)",
        "ocr_lang_label": "Jazyk OCR:",
        "ocr_multi_hint": "Tip: pre viacero jazykov držte Cmd/Ctrl a klikajte; Shift+klik pre rozsah.",
        "show_log": "▸ Zobraziť záznam",
        "hide_log": "▾ Skryť záznam",
        "convert_button": "Konvertovať",
        "status_initial": "Vyberte PDF na konverziu.",
        "status_ocr_starting": "Spúšťam OCR…",
        "status_ocr": "OCR strany {current}/{total}…",
        "status_converting": "Konvertujem…",
        "status_saved": "Uložené: {name}",
        "status_failed": "Konverzia zlyhala.",
        "no_file_title": "Žiadny súbor",
        "no_file_body": "Najskôr vyberte PDF súbor.",
        "not_found_title": "Súbor nenájdený",
        "not_found_body": "{path} neexistuje.",
        "select_pdf_title": "Vyberte PDF",
        "save_docx_title": "Uložiť DOCX ako",
        "done_title": "Hotovo",
        "done_body": "Uložené do:\n{path}",
        "failed_title": "Konverzia zlyhala",
        "filetype_pdf": "PDF súbory",
        "filetype_docx": "Word dokumenty",
        "filetype_all": "Všetky súbory",
    },
    "en": {
        "window_title": "PDF to DOCX Converter",
        "heading": "PDF to DOCX Converter",
        "subtitle": "Convert a PDF to an editable Word document.",
        "language_label": "Language:",
        "pdf_file_label": "PDF file:",
        "browse_button": "Browse…",
        "no_stream_checkbox": "Detect only tables with visible borders (recommended)",
        "ocr_checkbox": "Apply OCR (for scanned PDFs)",
        "ocr_lang_label": "OCR language:",
        "ocr_multi_hint": "Tip: for multiple languages hold Cmd/Ctrl and click; Shift+Click for ranges.",
        "show_log": "▸ Show log",
        "hide_log": "▾ Hide log",
        "convert_button": "Convert",
        "status_initial": "Pick a PDF to convert.",
        "status_ocr_starting": "Starting OCR…",
        "status_ocr": "OCR page {current}/{total}…",
        "status_converting": "Converting…",
        "status_saved": "Saved: {name}",
        "status_failed": "Conversion failed.",
        "no_file_title": "No file",
        "no_file_body": "Please pick a PDF file first.",
        "not_found_title": "File not found",
        "not_found_body": "{path} does not exist.",
        "select_pdf_title": "Select PDF",
        "save_docx_title": "Save DOCX as",
        "done_title": "Done",
        "done_body": "Saved to:\n{path}",
        "failed_title": "Conversion failed",
        "filetype_pdf": "PDF files",
        "filetype_docx": "Word documents",
        "filetype_all": "All files",
    },
}

# Native language names shown in the selector → language code.
LANGUAGE_NAMES: dict[str, str] = {
    "Slovenčina": "sk",
    "English": "en",
}
DEFAULT_LANGUAGE = "sk"


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        # Must run before super().__init__() — Tk locks in DPI mode when
        # the root window is created.
        _enable_high_dpi()
        super().__init__()
        self._lang = DEFAULT_LANGUAGE

        # Detect tesseract once at startup. None = not installed → no OCR
        # row will be built at all. A list = available language codes.
        self._tesseract_langs: list[str] | None = list_tesseract_languages()

        # OCR language pairs in display-name order. Each entry is
        # (code, display_name). Used to populate the multi-select
        # listbox and look up codes from listbox indices when the
        # user starts a conversion. Unknown codes (e.g. custom
        # tesseract training data with no entry in TESSERACT_LANG_NAMES)
        # fall back to displaying the code itself.
        self._ocr_lang_pairs: list[tuple[str, str]] = []
        if self._tesseract_langs:
            self._ocr_lang_pairs = sorted(
                ((c, TESSERACT_LANG_NAMES.get(c, c)) for c in self._tesseract_langs),
                key=lambda p: p[1].lower(),
            )

        # The window has two heights: a compact one with the log hidden
        # (the default), and an expanded one when the user clicks the
        # "Show log" disclosure button. The OCR multi-select listbox
        # needs more vertical space than the old combobox so the
        # tesseract-present heights are bumped accordingly.
        self._window_width = 620
        self._compact_height = 470 if self._tesseract_langs else 360
        self._expanded_height = 640 if self._tesseract_langs else 540
        x = (self.winfo_screenwidth() - self._window_width) // 2
        y = (self.winfo_screenheight() - self._compact_height) // 2
        self.geometry(f"{self._window_width}x{self._compact_height}+{x}+{y}")
        # Resizable so the user can grow the log widget when needed.
        self.resizable(True, True)
        self.minsize(540, 320)
        self._log_visible = False

        self._input_path = tk.StringVar()
        self._no_stream = tk.BooleanVar(value=True)
        self._status = tk.StringVar()
        # Track status by translation key + format kwargs so it can be
        # re-rendered into the new language when the user switches.
        self._status_key = "status_initial"
        self._status_kwargs: dict[str, object] = {}

        # OCR state — only meaningful when tesseract was detected. The
        # actual selection lives in the listbox built in _build_ui;
        # this just tracks whether OCR mode is on.
        self._ocr_enabled = tk.BooleanVar(value=False)

        self._build_ui()
        self._apply_language()
        # Defer focus until after Tk maps the window. focus_force() and
        # lift() on an unmapped window are unreliable across platforms;
        # after(0, …) queues the call so it runs once mainloop() is up
        # and the window is realized.
        self.after(0, self._focus_window)

    def _default_ocr_index(self) -> int | None:
        """Return the listbox index of the default OCR language, or
        None if no tesseract languages are available. Prefers Slovak,
        then English, then the first language in the list (which is
        already sorted by display name)."""
        if not self._ocr_lang_pairs:
            return None
        for preferred in ("slk", "eng"):
            for i, (code, _display) in enumerate(self._ocr_lang_pairs):
                if code == preferred:
                    return i
        return 0

    def _focus_window(self) -> None:
        """Bring the window to the foreground on launch.

        Cross-platform Tk-only "kitchen sink" focus combo:

        1. update_idletasks() forces Tk to flush any pending geometry/map
           events so the window is fully realized before we touch it.
        2. deiconify() ensures the window isn't iconified (minimized) —
           some WMs leave it that way until first focus.
        3. lift() raises the window to the top of the stacking order.
        4. attributes("-topmost", True) briefly marks it as always-on-top
           so the WM grants it focus, then we revert with after_idle so
           the user can put other windows over it normally afterward.
        5. focus_force() pulls keyboard focus.

        Works well on Linux and Windows. On macOS the reliability depends
        on how Python was launched: bundled .app/pythonw → reliable; bare
        `python gui.py` from Terminal → may still need a manual click
        because Terminal-launched Python is treated as a background
        process by the macOS WindowServer (a fundamental Python-on-macOS
        limitation that no Tk API can override).
        """
        self.update_idletasks()
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(self.attributes, "-topmost", False)
        self.focus_force()

    # ----- translation helpers -------------------------------------------

    def _t(self, key: str, **kwargs: object) -> str:
        template = TRANSLATIONS[self._lang][key]
        return template.format(**kwargs) if kwargs else template

    def _set_status(self, key: str, **kwargs: object) -> None:
        self._status_key = key
        self._status_kwargs = kwargs
        self._status.set(self._t(key, **kwargs))

    # ----- UI construction -----------------------------------------------

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)

        # Top row: language selector aligned right
        top_row = ttk.Frame(outer)
        top_row.pack(fill="x")
        self._lang_selector = ttk.Combobox(
            top_row,
            values=list(LANGUAGE_NAMES.keys()),
            state="readonly",
            width=12,
        )
        # Pre-select the default language by its native name
        for name, code in LANGUAGE_NAMES.items():
            if code == self._lang:
                self._lang_selector.set(name)
                break
        self._lang_selector.pack(side="right")
        self._lang_selector.bind("<<ComboboxSelected>>", self._on_language_change)
        self._lang_label = ttk.Label(top_row, text="")
        self._lang_label.pack(side="right", padx=(0, 6))

        # Heading + subtitle
        self._heading = ttk.Label(
            outer, text="", font=("TkDefaultFont", 14, "bold")
        )
        self._heading.pack(anchor="w", pady=(12, 0))
        self._subtitle = ttk.Label(outer, text="", foreground="#666")
        self._subtitle.pack(anchor="w", pady=(0, 12))

        # File picker row
        self._file_label = ttk.Label(outer, text="")
        self._file_label.pack(anchor="w")
        row = ttk.Frame(outer)
        row.pack(fill="x", pady=(4, 0))
        ttk.Entry(row, textvariable=self._input_path).pack(
            side="left", fill="x", expand=True
        )
        self._browse_btn = ttk.Button(row, text="", command=self._pick_input)
        self._browse_btn.pack(side="left", padx=(8, 0))

        # OCR section — only built when tesseract is installed.
        # Layout is now vertical (checkbox row, then a multi-select
        # listbox, then a hint) because the listbox needs more
        # vertical space than the old combobox and the user has
        # ~160 languages to scroll through.
        self._ocr_chk: ttk.Checkbutton | None = None
        self._ocr_lang_label_widget: ttk.Label | None = None
        self._ocr_lang_listbox: tk.Listbox | None = None
        self._ocr_lang_hint_widget: ttk.Label | None = None
        if self._tesseract_langs:
            self._ocr_chk = ttk.Checkbutton(
                outer,
                text="",
                variable=self._ocr_enabled,
                command=self._on_ocr_toggle,
            )
            self._ocr_chk.pack(anchor="w", pady=(8, 4))

            self._ocr_lang_label_widget = ttk.Label(outer, text="", foreground="#666")
            self._ocr_lang_label_widget.pack(anchor="w", padx=(20, 0))

            # Listbox + vertical scrollbar in their own frame so the
            # scrollbar sits flush against the listbox.
            lb_frame = ttk.Frame(outer)
            lb_frame.pack(fill="x", padx=(20, 0), pady=(2, 0))
            scrollbar = ttk.Scrollbar(lb_frame, orient="vertical")
            self._ocr_lang_listbox = tk.Listbox(
                lb_frame,
                selectmode="extended",  # Cmd/Ctrl+Click multi, Shift+Click range
                height=6,
                exportselection=False,  # don't lose selection on focus loss
                yscrollcommand=scrollbar.set,
            )
            scrollbar.config(command=self._ocr_lang_listbox.yview)
            scrollbar.pack(side="right", fill="y")
            self._ocr_lang_listbox.pack(side="left", fill="both", expand=True)

            # Populate with display names — _run_conversion uses the
            # listbox indices + self._ocr_lang_pairs to recover the
            # tesseract codes.
            for _code, display in self._ocr_lang_pairs:
                self._ocr_lang_listbox.insert("end", display)

            # Pre-select the default language (slk → eng → first).
            default_idx = self._default_ocr_index()
            if default_idx is not None:
                self._ocr_lang_listbox.selection_set(default_idx)
                self._ocr_lang_listbox.see(default_idx)

            # Start disabled — the OCR checkbox toggles state.
            self._ocr_lang_listbox.config(state="disabled")

            self._ocr_lang_hint_widget = ttk.Label(
                outer, text="", foreground="#999",
            )
            self._ocr_lang_hint_widget.pack(anchor="w", padx=(20, 0), pady=(2, 0))

        # Options
        self._no_stream_chk = ttk.Checkbutton(
            outer, text="", variable=self._no_stream
        )
        self._no_stream_chk.pack(anchor="w", pady=(8, 12))

        # Convert button
        self._convert_btn = ttk.Button(outer, text="", command=self._on_convert)
        self._convert_btn.pack(fill="x")

        # Status line — single-line at-a-glance state. We don't wrap it
        # because the full path goes to the log widget below; the status
        # only shows short messages like "Converting…" or "Saved: <name>".
        # The version label is right-aligned on the same row so it
        # always remains visible regardless of status updates.
        status_row = ttk.Frame(outer)
        status_row.pack(fill="x", pady=(12, 4))
        ttk.Label(
            status_row, textvariable=self._status, foreground="#666"
        ).pack(side="left")
        ttk.Label(
            status_row,
            text=f"{__codename__} v{__version__} ({__release_date__})",
            foreground="#999",
        ).pack(side="right")

        # "Show log ▸" / "Hide log ▾" disclosure button — collapses the
        # log widget by default to keep the window compact, lets the
        # user expand it on demand to watch progress.
        self._log_toggle_btn = ttk.Button(
            outer, text="", command=self._toggle_log, style="Toolbutton"
        )
        self._log_toggle_btn.pack(anchor="w", pady=(0, 4))

        # Log widget — created but NOT packed initially (hidden by
        # default). _toggle_log packs/forgets it on demand. Captures
        # pdf2docx's [INFO] progress lines via a logging.Handler plus
        # any stdout/stderr from the worker.
        self._log = ScrolledText(
            outer,
            wrap="word",
            height=6,
            font=("Menlo", 12) if sys.platform == "darwin" else ("Consolas", 11),
            background="#1e1e1e",
            foreground="#dcdcdc",
            insertbackground="#dcdcdc",
            relief="flat",
            borderwidth=1,
            state="disabled",
        )
        # Intentionally NOT packed — _toggle_log handles visibility.

    def _apply_language(self) -> None:
        """Re-render every translatable widget after a language switch."""
        self.title(
            f"{__codename__} — {self._t('window_title')}  "
            f"·  v{__version__} ({__release_date__})"
        )
        self._lang_label.config(text=self._t("language_label"))
        self._heading.config(text=self._t("heading"))
        self._subtitle.config(text=self._t("subtitle"))
        self._file_label.config(text=self._t("pdf_file_label"))
        self._browse_btn.config(text=self._t("browse_button"))
        self._no_stream_chk.config(text=self._t("no_stream_checkbox"))
        if self._ocr_chk is not None:
            self._ocr_chk.config(text=self._t("ocr_checkbox"))
        if self._ocr_lang_label_widget is not None:
            self._ocr_lang_label_widget.config(text=self._t("ocr_lang_label"))
        if self._ocr_lang_hint_widget is not None:
            self._ocr_lang_hint_widget.config(text=self._t("ocr_multi_hint"))
        self._convert_btn.config(text=self._t("convert_button"))
        self._log_toggle_btn.config(
            text=self._t("hide_log" if self._log_visible else "show_log")
        )
        # Re-render the status line in the new language using the tracked
        # key and kwargs, so an in-progress "Converting…" or a "Saved: …"
        # message gets translated instead of being clobbered.
        self._status.set(self._t(self._status_key, **self._status_kwargs))

    def _on_ocr_toggle(self) -> None:
        """Enable/disable the OCR language listbox to match the checkbox."""
        if self._ocr_lang_listbox is None:
            return
        if self._ocr_enabled.get():
            self._ocr_lang_listbox.config(state="normal")
        else:
            self._ocr_lang_listbox.config(state="disabled")

    # ----- log helpers (main thread only) --------------------------------

    def _log_clear(self) -> None:
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _log_append(self, text: str) -> None:
        """Append text to the log and auto-scroll. MUST be called from
        the main thread (use ``self.after(0, self._log_append, text)``
        from worker threads)."""
        self._log.config(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.config(state="disabled")

    def _toggle_log(self) -> None:
        """Show or hide the log widget and resize the window accordingly."""
        if self._log_visible:
            self._log.pack_forget()
            self._log_visible = False
            self._log_toggle_btn.config(text=self._t("show_log"))
            target_h = self._compact_height
        else:
            self._log.pack(fill="both", expand=True, pady=(4, 0))
            self._log_visible = True
            self._log_toggle_btn.config(text=self._t("hide_log"))
            target_h = self._expanded_height
        # Preserve current width and the manually-resized state — only
        # change height. winfo_width() may return 1 before the first
        # geometry pass, so fall back to the configured width.
        cur_w = self.winfo_width()
        if cur_w <= 1:
            cur_w = self._window_width
        self.geometry(f"{cur_w}x{target_h}")

    def _on_language_change(self, _event: object = None) -> None:
        selected_name = self._lang_selector.get()
        new_code = LANGUAGE_NAMES.get(selected_name)
        if new_code and new_code != self._lang:
            self._lang = new_code
            self._apply_language()

    # ----- conversion flow -----------------------------------------------

    def _pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title=self._t("select_pdf_title"),
            filetypes=[
                (self._t("filetype_pdf"), "*.pdf"),
                (self._t("filetype_all"), "*.*"),
            ],
        )
        if path:
            self._input_path.set(path)

    def _on_convert(self) -> None:
        raw = self._input_path.get().strip()
        if not raw:
            messagebox.showwarning(
                self._t("no_file_title"), self._t("no_file_body")
            )
            return
        pdf_path = Path(raw)
        if not pdf_path.is_file():
            messagebox.showerror(
                self._t("not_found_title"),
                self._t("not_found_body", path=pdf_path),
            )
            return

        out_path_str = filedialog.asksaveasfilename(
            title=self._t("save_docx_title"),
            initialdir=str(pdf_path.parent),
            initialfile=pdf_path.stem + ".docx",
            defaultextension=".docx",
            filetypes=[(self._t("filetype_docx"), "*.docx")],
        )
        if not out_path_str:
            return
        out_path = Path(out_path_str)

        self._convert_btn.config(state="disabled")
        self._set_status("status_converting")

        # Reset the log for this run so the user sees just the current
        # conversion's output, then write a header line so multi-run
        # sessions are visually delimited.
        self._log_clear()
        self._log_append(f">>> {pdf_path}\n")

        threading.Thread(
            target=self._run_conversion,
            args=(pdf_path, out_path),
            daemon=True,
        ).start()

    def _run_conversion(self, pdf_path: Path, out_path: Path) -> None:
        # Capture conversion output into the log widget by two means:
        #
        # (1) A logging.Handler attached to the root logger captures
        #     pdf2docx's [INFO] progress lines. pdf2docx calls
        #     logging.basicConfig() at module import, which installs a
        #     StreamHandler bound to the *original* sys.stderr — by the
        #     time we redirect_stderr, that handler is already pointing
        #     at the wrong stream and ignores our redirect. The custom
        #     logging handler is the only way to intercept those lines.
        #
        # (2) contextlib.redirect_stdout/stderr around the conversion
        #     captures any plain print() output (e.g. PyMuPDF warnings)
        #     and routes it through the same _LogWriter.
        writer = _LogWriter(self)
        log_handler = _TkLogHandler(self)
        root_logger = logging.getLogger()
        root_logger.addHandler(log_handler)
        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                if self._ocr_enabled.get() and self._tesseract_langs:
                    # Scanned-PDF path: bypass pdf2docx entirely. pdf2docx
                    # cannot extract text from OCR'd PDFs (its layout analyzer
                    # requires real glyph metrics, which OCR text doesn't have)
                    # so we build the DOCX directly from tesseract's TSV output.
                    self.after(0, self._set_status, "status_ocr_starting")
                    # Read the multi-select listbox and join the chosen
                    # codes with "+" — the syntax tesseract accepts for
                    # multi-language passes (e.g. "slk+eng").
                    selected_codes: list[str] = []
                    if self._ocr_lang_listbox is not None:
                        for idx in self._ocr_lang_listbox.curselection():
                            if 0 <= idx < len(self._ocr_lang_pairs):
                                selected_codes.append(self._ocr_lang_pairs[idx][0])
                    if not selected_codes:
                        # Empty selection — fall back to the default so
                        # the conversion doesn't fail with no language.
                        default_idx = self._default_ocr_index()
                        if default_idx is not None:
                            selected_codes = [self._ocr_lang_pairs[default_idx][0]]
                        else:
                            selected_codes = ["eng"]
                    lang_code = "+".join(selected_codes)
                    ocr_to_docx(
                        pdf_path,
                        out_path,
                        language=lang_code,
                        progress_callback=self._on_ocr_progress,
                    )
                else:
                    # Regular text-PDF path: pdf2docx + post-processing fixes.
                    convert_pdf(
                        pdf_path,
                        out_path,
                        parse_stream_table=not self._no_stream.get(),
                    )
                    fix_bullet_fonts(out_path)
                    fix_first_table_header(pdf_path, out_path)
        except Exception as exc:  # noqa: BLE001 — surface anything to the user
            writer.flush()
            root_logger.removeHandler(log_handler)
            self.after(0, self._on_done, False, str(exc), out_path)
            return
        writer.flush()
        root_logger.removeHandler(log_handler)
        self.after(0, self._on_done, True, "", out_path)

    def _on_ocr_progress(self, current: int, total: int) -> None:
        """Called from the worker thread once per page during OCR."""
        # Marshal back onto the Tk main loop — Tk widgets are not thread-safe.
        self.after(
            0,
            lambda: self._set_status("status_ocr", current=current, total=total),
        )

    def _on_done(self, success: bool, error: str, out_path: Path) -> None:
        self._convert_btn.config(state="normal")
        if success:
            # Status shows just the filename (so it fits the window);
            # the full destination path goes into the log widget.
            self._set_status("status_saved", name=out_path.name)
            self._log_append(f"<<< saved to: {out_path}\n")
            messagebox.showinfo(
                self._t("done_title"), self._t("done_body", path=out_path)
            )
        else:
            self._set_status("status_failed")
            self._log_append(f"<<< FAILED: {error}\n")
            messagebox.showerror(self._t("failed_title"), error)


def main() -> None:
    ConverterApp().mainloop()


if __name__ == "__main__":
    main()
