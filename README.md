# PDF-to-PPTX Reference Placement Web App

Local desktop web app for cloning an uploaded PPTX, rendering the uploaded PDF into page images, and parking each PDF page beside its corresponding slide as a movable reference.

## Requirements

- macOS or Windows
- LibreOffice installed
- Microsoft PowerPoint only if you want optional PowerPoint fallback export
- Python 3.9+

## Install

macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows PowerShell:

```powershell
python -m venv .venv-windows
.\.venv-windows\Scripts\Activate.ps1
pip install -r requirements.txt
```

Note:

- The dependency set uses a NumPy pin compatible with both Python 3.9 and newer Windows x64 builds, so Python 3.13 x64 is fine for the Windows packaging flow.

## Run

macOS:

```bash
source .venv/bin/activate
uvicorn app.main:app --reload
```

or:

```bash
source .venv/bin/activate
python launch_app.py
```

Windows PowerShell:

```powershell
.\.venv-windows\Scripts\Activate.ps1
uvicorn app.main:app --reload
```

or:

```powershell
.\.venv-windows\Scripts\Activate.ps1
python .\launch_app.py
```

Open [http://127.0.0.1:8000](http://127.0.0.1:8000).

## What It Does

- Accepts one `.pdf` and one `.pptx`
- Exports the PPTX to PDF with LibreOffice by default so PowerPoint does not open during normal runs
- Supports macOS and Windows with platform-aware export defaults
- Can optionally fall back to Microsoft PowerPoint if `PDF_PPTX_ENABLE_POWERPOINT_FALLBACK=1`
- Rasterizes both PDFs into page images
- Detects fonts used in the uploaded PDF and records whether they appear embedded
- Chooses the best PDF page rotation for each slide from `0`, `90`, and `270`
- Creates a new PPTX with the original slide content intact
- Parks each matching PDF page as an opaque movable picture named `PDF_ORIGINAL` just off the top-right of the slide so reviewers can drag it in when needed
- Inserts a full-slide PDF screenshot immediately after each matched PPTX slide for slide-by-slide review, with the slide name set to `PDF_ORIGINAL`
- Bakes a red border into each parked reference image so the screenshot and its outline move as one piece
- Appends unmatched PDF pages as full-slide reference slides at the end of the deck
- Returns `<original-name>_with_pdf_pages.pptx`
- Writes a sidecar font report to `runs/<job_id>/output/pdf_fonts.json`

## Notes

- Jobs are stored under `runs/<job_id>/` and are cleaned up after 24 hours.
- The worker is intentionally single-threaded so PowerPoint automation never runs concurrently.
- Page matching is page-index based in v1.
- PDF pages are fit into the slide canvas without stretching. If a PDF page has a different aspect ratio, you will see padding rather than geometric distortion.
- Runtime job files are stored in the user application-data area by default instead of the project folder, which avoids `uvicorn --reload` restarts during uploads.

## Optional Environment Overrides

- `PDF_PPTX_LIBREOFFICE_BIN` points to a custom `soffice` or `soffice.exe`
- `PDF_PPTX_ENABLE_POWERPOINT_FALLBACK=1` enables PowerPoint export fallback
- `PDF_PPTX_POWERPOINT_APP_PATH` overrides the macOS PowerPoint app location
- `PDF_PPTX_POWERSHELL_BIN` overrides the Windows PowerShell executable used for PowerPoint fallback
- `PDF_PPTX_RUNS_DIR` overrides where temporary job folders are stored

## Packaging For Windows Users

If you want to send this to a non-technical Windows user, the easiest path is to build a one-folder app bundle:

```powershell
.\scripts\build_windows_bundle.ps1
```

That produces `dist-windows\PDFtoPPTXReference\`. Zip that whole folder and send it. The recipient can extract it and double-click `PDFtoPPTXReference.exe`.

Note:

- The Windows build script intentionally uses `.venv-windows` so it does not conflict with a macOS `.venv` when the project is opened through a Parallels shared folder.

Important:

- The packaged app still needs a renderer on the recipient machine
- LibreOffice is used first when available
- On Windows, PowerPoint fallback is enabled by default in packaged or normal runs, so a user with Microsoft PowerPoint installed can still export even without LibreOffice

## Packaging For Mac Users

If you want a Mac app bundle:

```bash
./scripts/build_macos_bundle.sh
```

That produces `dist/PDFtoPPTXReference.app` and you can zip it for sharing.

If you want a more normal Mac handoff as a disk image:

```bash
./scripts/build_macos_dmg.sh
```

That produces `deliverables/PDFtoPPTXReference-macOS.dmg` containing:

- `PDFtoPPTXReference.app`
- `RECIPIENT_MAC_README.txt`
- an `Applications` shortcut

Important:

- The packaged app still needs a renderer on the recipient machine
- LibreOffice is used first when available
- Because the app is not notarized, macOS may show a Gatekeeper warning the first time it is opened
