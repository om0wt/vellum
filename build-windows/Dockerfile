# Cross-compile the desktop GUI to a standalone Windows .exe via PyInstaller
# running under Wine inside Linux. The base image bundles a working Wine +
# Windows Python install so we don't have to fight with Wine ourselves.
#
# Build context MUST be the project root (so the Dockerfile can COPY
# build-windows/requirements-windows.txt). Build with build-windows/build.sh
# from the project root, or:
#
#   docker build -t pdf2docx-windows-builder -f build-windows/Dockerfile .
#   docker run --rm -v "$PWD:/src" pdf2docx-windows-builder
#
# After the run, the executable is at ./dist/PDF-to-DOCX.exe.

# tobix/pywine is only published for linux/amd64. The --platform pin is
# passed via `docker build --platform linux/amd64` (see build-windows/build.sh)
# rather than hardcoded in FROM, which avoids the FromPlatformFlagConstDisallowed
# BuildKit warning while still working on Apple Silicon hosts via Rosetta.
FROM tobix/pywine:3.12

WORKDIR /src

# Install the Windows Python deps once at image-build time so subsequent
# `docker run` invocations only need to re-run PyInstaller.
COPY build-windows/requirements-windows.txt /tmp/requirements-windows.txt
RUN xvfb-run wine pip install --no-cache-dir \
        -r /tmp/requirements-windows.txt \
        pyinstaller

# At runtime, mount the project root at /src and PyInstaller picks up
# gui.py + pdf_to_docx.py from there. Output goes to /src/dist on the host
# as a folder /src/dist/PDF-to-DOCX/ containing PDF-to-DOCX.exe alongside
# its runtime files. Distribute the whole folder (zipped).
#
# Build flag rationale:
# --onedir vs --onefile: onedir avoids the temp-extraction startup hit and
#   prevents the AV false-positives that --onefile self-extractors trigger.
# --noupx: UPX compression is the single biggest cause of Windows Defender
#   flagging Python-built .exes as malware. Disable it for distribution.
# --collect-all (fitz/pdf2docx/docx): defensive — these packages ship data
#   files that PyInstaller's static analysis sometimes misses. fitz and
#   pymupdf are the same package; only --collect-all fitz is needed.
# --exclude-module: trim known-unused modules out of the bundle. The
#   excludes here are conservative — nothing the runtime path actually
#   imports. Things like unittest/distutils/email are intentionally NOT
#   excluded because some libraries import them at module load.
CMD xvfb-run wine pyinstaller \
        --noconfirm \
        --onedir \
        --noupx \
        --windowed \
        --name PDF-to-DOCX \
        --paths /src \
        --collect-all pdf2docx \
        --collect-all fitz \
        --collect-all docx \
        --exclude-module IPython \
        --exclude-module jupyter \
        --exclude-module notebook \
        --exclude-module matplotlib \
        --exclude-module scipy \
        --exclude-module pandas \
        --exclude-module pytest \
        --exclude-module _pytest \
        --exclude-module sphinx \
        --exclude-module pylint \
        --exclude-module setuptools \
        --exclude-module pip \
        --exclude-module wheel \
        --exclude-module babel \
        --exclude-module tkinter.test \
        --exclude-module numpy.distutils \
        --exclude-module numpy.f2py \
        --exclude-module numpy.testing \
        --exclude-module numpy.tests \
        --distpath /src/dist \
        --workpath /tmp/build \
        --specpath /tmp \
        /src/gui.py
