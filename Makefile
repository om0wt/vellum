# Convenience targets for the Vellum project. Most day-to-day commands
# are still plain `python src/...` and `docker compose ...` — this
# Makefile only wraps the bits that benefit from a one-liner shortcut
# (mainly: cutting a release).
#
# Run `make help` for the full list.

# Read the version straight from src/_version.py so we never have to
# duplicate it in two places.
VERSION := $(shell python3 -c "import sys; sys.path.insert(0, 'src'); from _version import __version__; print(__version__)")

.PHONY: help version git-release run-cli run-gui run-web docker-up docker-down

help:
	@echo "Vellum — make targets"
	@echo
	@echo "  make version       Print the current version from src/_version.py"
	@echo "  make git-release   Tag the current commit as v\$$VERSION and push the"
	@echo "                     tag to origin. The push triggers .github/workflows/"
	@echo "                     build-windows.yml which builds the Windows .exe and"
	@echo "                     attaches it to a GitHub Release."
	@echo
	@echo "  make run-cli       (no args) — show CLI help"
	@echo "  make run-gui       Launch the Tkinter desktop GUI"
	@echo "  make run-web       Launch the Flask web app"
	@echo "  make docker-up     Build and start the web app container"
	@echo "  make docker-down   Stop the web app container"

version:
	@echo "$(VERSION)"

# Cut a release: tag the current HEAD and push the tag.
# Refuses if the working tree is dirty so a release always corresponds
# to a clean, committed state. The push to origin is what triggers the
# Windows-build workflow.
git-release:
	@if [ -z "$(VERSION)" ]; then \
		echo "ERROR: could not read version from src/_version.py"; \
		exit 1; \
	fi
	@if ! git diff --quiet || ! git diff --cached --quiet; then \
		echo "ERROR: working tree has uncommitted changes — commit or stash first."; \
		git status --short; \
		exit 1; \
	fi
	@if git rev-parse "v$(VERSION)" >/dev/null 2>&1; then \
		echo "ERROR: tag v$(VERSION) already exists. Bump src/_version.py first."; \
		exit 1; \
	fi
	@if ! git remote get-url origin >/dev/null 2>&1; then \
		echo "ERROR: no 'origin' remote configured. Run:"; \
		echo "       git remote add origin git@github.com:USER/REPO.git"; \
		exit 1; \
	fi
	@echo ">> Tagging current commit as v$(VERSION)…"
	git tag -a "v$(VERSION)" -m "Release v$(VERSION)"
	@echo ">> Pushing tag to origin…"
	git push origin "v$(VERSION)"
	@echo
	@echo "✓ v$(VERSION) released."
	@echo "  GitHub Actions will now build the Windows .exe and attach it"
	@echo "  to https://github.com/\$$(git remote get-url origin | sed 's|.*github.com[:/]||;s|\.git$$||')/releases/tag/v$(VERSION)"

run-cli:
	python src/pdf_to_docx.py --help

run-gui:
	python src/gui.py

run-web:
	python src/app.py

docker-up:
	docker compose -f docker/docker-compose.yml up --build

docker-down:
	docker compose -f docker/docker-compose.yml down
