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
docker-compose -f docker/docker-compose.yml up --build
# (or `docker compose ...` if you have the v2 plugin)
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

### Code signing — free options

Without a code signature, Windows SmartScreen and Defender flag every
download as an "unknown publisher" binary and prompt the user to
report it. Vellum doesn't currently sign its Windows release because
all paid options have ongoing fees and signing-cert prices have
risen sharply since the CA/Browser Forum's 2023 hardware-key rules.

Realistic free options:

| Option | Fixes warning? | Effort |
|---|---|---|
| **Don't sign + document the bypass** | No, but users can click "Run anyway" | Zero (status quo) |
| **Self-signed certificate** | No, still warns — but power users can manually trust your CA once | Medium per user |
| **Submit each release to Microsoft Defender for analysis** | **Yes, actually fixes it** | Per-release manual submission, ~1–7 days wait |
| **Microsoft Store (MSIX package)** | Yes (Store apps are MS-signed automatically) | $19 one-time, big distribution-model change |

#### Recommended free path: submit each release to Microsoft Defender

<https://www.microsoft.com/en-us/wdsi/filesubmission> is Microsoft's
official "this is not malware" submission portal. Many open source
maintainers use it as their de-facto code-signing substitute when
they don't want to pay for a cert.

**About PyInstaller bundles**: Vellum builds with `--onedir`, so the
release artifact is a *folder* (`PDF-to-DOCX/`) containing a small
launcher `PDF-to-DOCX.exe` (~1 MB) plus an `_internal/` directory
with ~50 MB of Python runtime DLLs (`python311.dll`, `pdf2docx`,
`fitz`, etc). **Defender's heuristic almost always flags the
launcher `.exe`** — the bootstrap stub that PyInstaller generates
matches `Trojan:Win32/Wacatac.B!ml` and similar false-positive
signatures. The DLLs in `_internal/` are legitimate Python runtime
binaries with their own signatures and rarely trigger warnings.

So you submit two things to Microsoft per release: the **launcher
`.exe`** (Defender's actual target) and the **release `.zip`**
(SmartScreen's download-reputation target).

**Per-release flow:**

1. Cut the release as usual (`make git-release`) → unsigned bundle
   lands in the GitHub Release as
   `Vellum-PDF-to-DOCX-vX.Y.Z-windows.zip`.
2. On a Windows machine, download and extract the zip:
   ```powershell
   Expand-Archive Vellum-PDF-to-DOCX-v1.1.0-windows.zip -DestinationPath .
   ```
3. Go to <https://www.microsoft.com/en-us/wdsi/filesubmission>.
4. **First submission** — the launcher:
   - **Role**: Software developer
   - **File**: upload `PDF-to-DOCX\PDF-to-DOCX.exe` (the ~1 MB
     launcher inside the extracted folder, NOT the zip)
   - **Detection name**: whatever Defender flagged it as (commonly
     `Trojan:Win32/Wacatac.B!ml` for PyInstaller binaries)
   - **Reason**: Incorrect detection (false positive)
   - **Description**:
     > Open source PDF to DOCX converter, MIT licensed, source at
     > <https://github.com/om0wt/vellum>, built by GitHub Actions
     > from a public workflow. The `.exe` is a PyInstaller `--onedir`
     > launcher; the false-positive heuristic matches against the
     > PyInstaller bootloader, not actual malicious code. The build
     > workflow that produced this binary is at
     > <https://github.com/om0wt/vellum/actions>.
   - Link to the specific GitHub Release
5. **Second submission** — the zip:
   - Same form, upload `Vellum-PDF-to-DOCX-v1.1.0-windows.zip` this
     time. This addresses SmartScreen's *download-reputation* layer
     (the URL + zip hash) in addition to the executable layer.
6. You'll get an email confirmation + tracking ID for each
   submission.
7. Wait 1–7 days. Microsoft analysts review. For an MIT-licensed
   Python+PyInstaller GUI with a public CI build, they almost
   always conclude it's a false positive and add both hashes to
   Defender's cloud safe list.
8. Future downloads of **those specific files** (same SHA-256
   hashes) no longer trigger SmartScreen / Defender warnings.
   **You have to redo this per release** because each PyInstaller
   build has different hashes (the launcher and the zip both
   change).

This is slow and manual, but it's the only way to actually clear
the Defender warning for free.

**Tip**: don't bother re-submitting between manual workflow
dispatches and a real release — only submit the binary that's
actually attached to a tagged GitHub Release, since that's the one
end users will download.

#### Tell users how to bypass the warning when it does appear

For releases that haven't been submitted yet (or whose Defender
review is still pending), include this snippet in your release notes
so users know what's happening:

> **Windows SmartScreen / Defender warning?** Vellum is open source
> and unsigned (signing certificates cost €25–400 per year and the
> project doesn't yet have a budget for one). The binary is built by
> a public GitHub Actions workflow from the source you can see in
> this repo. To run it the first time:
>
> 1. Click **More info** on the SmartScreen dialog
> 2. Click **Run anyway**
>
> Or, if Defender quarantines the file, restore it via Virus &
> threat protection → Protection history → Allow.

#### Project policy

**Vellum is and will remain unsigned.** Code signing certificates
cost €25–400/year and have no free option that the project considers
acceptable. The Microsoft Defender submission flow above + the
user-facing "Run anyway" instructions in release notes are the
official solution.

A dormant `signtool` workflow step in
`.github/workflows/build-windows.yml` is kept for **forks** that may
have access to a signing certificate (commercial, EV, or enterprise).
It gracefully skips when the `WINDOWS_CERT_PFX` secret isn't set, so
it costs nothing to leave in place — the upstream Vellum builds just
never trigger it.

#### Important: post-2023 hardware-key requirement

Since June 2023 the CA/Browser Forum requires that the private key
for any new Standard Code Signing certificate be stored in **hardware**
(FIPS 140-2 Level 2+) — i.e. a physical USB token or a cloud HSM.
**A pure-PFX-on-disk workflow is no longer available** for new certs.
This affects Certum, Sectigo, DigiCert, and every other CA equally.

For Certum specifically, the OSS cert is delivered through **Certum
SimplySign**, their cloud HSM service. You sign by authenticating
to SimplySign over the internet — your private key never leaves
their hardware. The signing tool uses a virtual smart card driver
that signtool.exe can talk to.

**Practical consequence**: the GitHub Actions runner cannot directly
hold the Certum private key. You have two options:

1. **Sign locally, upload manually** — let the workflow build the
   unsigned `.exe`, download it from the workflow artifact page,
   sign it on your own machine using SimplySign + signtool, then
   upload the signed `.zip` to the GitHub Release manually. Slower
   but doesn't require any special infrastructure.
2. **Self-hosted runner** with the SimplySign virtual smart card
   pre-installed and authenticated. The workflow's `Sign Windows
   .exe` step then runs against the cloud HSM transparently. Faster
   for frequent releases, but requires a Windows machine you keep
   running and maintaining.

For Vellum's release cadence (occasional, not daily), **option 1** is
the right call. Instructions below.

#### Step 1: Apply for the Certum OSS certificate

1. Go to <https://shop.certum.eu/> and search for **"Open Source Code
   Signing"** (the exact product page URL changes occasionally).
2. Start the order. The OSS cert is free, but you still go through
   the checkout flow.
3. Fill out the application form. You'll need:
   - Your full legal name and address
   - A government-issued ID (passport or national ID card) — they
     verify identity manually
   - The project's GitHub URL (`https://github.com/om0wt/vellum`)
   - The MIT LICENSE file in the repo (Certum verifies the project
     is actually open source)
4. Submit identity verification documents through their secure
   portal. Some applicants are also asked to do a short video call.
5. **Wait 1–2 weeks** for review. They'll email you when the cert is
   ready.

#### Step 2: Activate SimplySign + install signtool

Once Certum approves the cert, they email you SimplySign credentials.

1. Install **SimplySign Desktop** from
   <https://support.certum.eu/en/cert-offer-simplysigndesktop/>. It
   provides the virtual smart card driver and the SimplySign login
   client.
2. Install the **Windows SDK** (for `signtool.exe`) — easiest via
   Visual Studio Build Tools, or directly from
   <https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/>.
3. Log into SimplySign Desktop with the credentials Certum sent.
   Your cert appears as a virtual smart card available to Windows
   crypto APIs (and therefore to signtool).

#### Step 3: Sign Vellum locally before releasing

Cut the release as normal:

```bash
make git-release
```

This pushes the tag, the GitHub Actions workflow runs, and produces
an **unsigned** `Vellum-PDF-to-DOCX-vX.Y.Z-windows.zip` as a workflow
artifact (because `WINDOWS_CERT_PFX` isn't set in this setup).

On your Windows machine:

```powershell
# Download the unsigned zip from the workflow run page
# Extract it
Expand-Archive Vellum-PDF-to-DOCX-v1.1.0-windows.zip -DestinationPath .

# Make sure SimplySign Desktop is logged in (system tray icon should
# show "Connected"). Then sign the .exe — signtool will pick up the
# Certum cert from the virtual smart card automatically.
signtool sign `
  /sha1 <YOUR_CERT_THUMBPRINT> `
  /tr http://timestamp.sectigo.com `
  /td sha256 `
  /fd sha256 `
  /d "Vellum — PDF to DOCX Converter" `
  /du "https://github.com/om0wt/vellum" `
  PDF-to-DOCX\PDF-to-DOCX.exe

# Verify
signtool verify /pa /v PDF-to-DOCX\PDF-to-DOCX.exe

# Re-zip the now-signed folder
Compress-Archive -Path PDF-to-DOCX -DestinationPath Vellum-PDF-to-DOCX-v1.1.0-windows.zip -Force
```

You can find your cert thumbprint with:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -match "Pavol Calfa" }
```

Copy the `Thumbprint` field (40 hex characters, no spaces).

#### Step 4: Upload the signed zip to the GitHub Release

The GitHub Actions workflow already created a draft Release at
`github.com/om0wt/vellum/releases/tag/v1.1.0`. Edit that release:

1. **Delete the unsigned zip** that the workflow attached.
2. **Upload the new signed zip** from your Windows machine.
3. Save.

Anyone downloading the release now gets the signed binary.

#### Step 5: Verify the signature on a fresh Windows install

```powershell
# After downloading and extracting on a Windows machine you didn't
# sign on:
Get-AuthenticodeSignature .\PDF-to-DOCX\PDF-to-DOCX.exe
```

Expected output:
```
SignerCertificate:  [...your name...] (Certum Code Signing CA SHA2)
Status:             Valid
StatusMessage:      Signature verified.
```

If you see `Valid`, the cert chain validates against Windows's trust
store, the timestamp is present (so the signature stays valid past
cert expiry), and SmartScreen will start building reputation against
your publisher name. The very first downloads will still warn, but
"Show more → Run anyway" is now there as an option, and the warning
disappears entirely once Microsoft's reputation system trusts you
(typically a few hundred to a few thousand downloads over weeks).

#### Future: automating with a self-hosted runner (option 2)

If release cadence picks up and the manual local-sign step gets
annoying, you can wire SimplySign into a self-hosted Windows runner.
The workflow's existing `Sign Windows .exe` step (in
`.github/workflows/build-windows.yml`) is already structured to use
`signtool` — replace the `/f $pfxPath /p $env:PFX_PASSWORD` arguments
with `/sha1 <THUMBPRINT>` and the rest works as-is, *if* the runner
has SimplySign Desktop installed and authenticated. That's left as a
future exercise; the local-sign-and-upload flow above is enough for
the project's current release cadence.

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
