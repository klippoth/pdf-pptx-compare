# PDF-to-PPTX Reference Placement Web App

Local desktop web app for cloning an uploaded PPTX, rendering the uploaded PDF into page images, and parking each PDF page beside its corresponding slide as a movable reference.

## Requirements

- macOS or Windows
- LibreOffice or Microsoft PowerPoint installed
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
- Automatically uses Microsoft PowerPoint when LibreOffice is not installed but PowerPoint is available on the machine
- Can still force PowerPoint fallback behavior with `PDF_PPTX_ENABLE_POWERPOINT_FALLBACK=1`
- Rasterizes both PDFs into page images
- Detects fonts used in the uploaded PDF and records whether they appear embedded
- Runs rendered-slide QC against the exported PPTX PDF and the reference PDF
- Uses parallelized GPT slide-by-slide image comparison when `OPENAI_API_KEY` is configured
- Sends the rendered PDF page and the rendered PPT slide as the two image inputs for each comparison
- Lets you append optional run-specific AI QC instructions from the web app without replacing the built-in review prompt
- On Windows, PowerPoint can export the candidate slide images directly through native PowerPoint automation; no AppleScript or VBA add-in is required there
- On macOS, direct candidate slide-image export uses the `ExportSlidesToFolder` PowerPoint VBA macro when you want native PowerPoint slide PNGs
- Falls back to native/OCR/visual QC only when GPT QC is not configured
- Focuses QC on obvious mistakes: missing elements, extra elements, wrong text, wrong colors, and obvious line breaks
- Annotates the original PPT slide with QC rectangles and a concise GPT-generated summary comment, while keeping the inserted `PDF_ORIGINAL` reference slide after it
- Preserves the PDF page orientation instead of auto-rotating it to match the slide
- Creates a new PPTX with the original slide content intact
- Inserts a full-slide PDF screenshot immediately after each matched PPTX slide for slide-by-slide review, with the slide and screenshot shape name set to `PDF_ORIGINAL`
- Appends unmatched PDF pages as full-slide reference slides at the end of the deck
- Returns `<original-name>_with_pdf_pages.pptx`
- Writes a sidecar font report to `runs/<job_id>/output/pdf_fonts.json`
- Writes a sidecar QC report to `runs/<job_id>/output/qc_report.json`

## Notes

- Jobs are stored under `runs/<job_id>/` and are cleaned up after 24 hours.
- The worker is intentionally single-threaded so PowerPoint automation never runs concurrently.
- Page matching is page-index based in v1.
- PDF pages are fit into the slide canvas without stretching. If a PDF page has a different aspect ratio, you will see padding rather than geometric distortion.
- Runtime job files are stored in the user application-data area by default instead of the project folder, which avoids `uvicorn --reload` restarts during uploads.
- If `OPENAI_API_KEY` is set, slide QC is driven by a GPT vision model comparing the PDF and PPT renders directly.
- Without `OPENAI_API_KEY`, the app falls back to rule-based native/OCR/visual QC.

## Windows Notes

- Windows does not use the macOS AppleScript / VBA add-in path.
- If Microsoft PowerPoint is installed on Windows, the app can:
  - export PPTX to PDF through native PowerPoint automation
  - export slide PNGs directly through native PowerPoint automation
- That makes the Windows AI-comparison path simpler than the macOS setup.
- For the strongest Windows QC path, install Microsoft PowerPoint. LibreOffice is still supported as a PDF export fallback.

## Optional Environment Overrides

- `PDF_PPTX_LIBREOFFICE_BIN` points to a custom `soffice` or `soffice.exe`
- `PDF_PPTX_ENABLE_POWERPOINT_FALLBACK=1` enables PowerPoint export fallback
- `PDF_PPTX_POWERPOINT_APP_PATH` overrides the macOS PowerPoint app location
- `PDF_PPTX_POWERSHELL_BIN` overrides the Windows PowerShell executable used for PowerPoint fallback
- `PDF_PPTX_RUNS_DIR` overrides where temporary job folders are stored
- `PDF_PPTX_GOOGLE_DOC_AI_PROJECT_ID` enables Google Document AI OCR fallback
- `PDF_PPTX_GOOGLE_DOC_AI_LOCATION` sets the Google Document AI processor location
- `PDF_PPTX_GOOGLE_DOC_AI_PROCESSOR_ID` sets the Google Document AI processor ID
- `PDF_PPTX_GOOGLE_DOC_AI_PROCESSOR_VERSION` optionally pins a processor version
- `PDF_PPTX_OPENAI_API_KEY` overrides the OpenAI key for this project only
- `OPENAI_API_KEY` enables GPT-based slide QC
- `PDF_PPTX_OPENAI_QC_MODEL` overrides the GPT model used for slide QC
- `PDF_PPTX_OPENAI_QC_PARALLELISM` controls how many slide comparisons run in parallel
- `PDF_PPTX_OPENAI_QC_TIMEOUT_SECONDS` sets the per-slide GPT request timeout
- `PDF_PPTX_OPENAI_QC_MAX_IMAGE_DIMENSION` caps the long edge of images sent to OpenAI; set `0` to send the original render quality
- each slide QC debug folder also includes `00-upload-metadata.json`, which records the exact uploaded PNG dimensions, byte sizes, and hashes for the candidate and reference images

Project-local key option:

- You can create [`.env.local.example`](/Users/kimlippoth/Desktop/Ethos/PDF%20to%20PPTX/Comp/.env.local.example) as `.env.local` in the project root.
- Put `PDF_PPTX_OPENAI_API_KEY=...` there to override any global `OPENAI_API_KEY` only for this app.
- `.env.local` is gitignored.
- In packaged builds, the app also looks for `.env` and `.env.local` beside the executable. On macOS it additionally checks the folder containing the `.app` bundle, so a shared build can carry its own project-local key without touching a colleague's global shell environment.

## Packaging For Windows Users

If you want to send this to a non-technical Windows user, the easiest path is to build a one-folder app bundle:

```powershell
.\scripts\build_windows_bundle.ps1
```

That produces `dist-windows\PDFtoPPTXReference\`. Zip that whole folder and send it. The recipient can extract it and double-click `PDFtoPPTXReference.exe`.

If `.env.local` exists in the project root when you build, the script copies it into the packaged folder automatically. That is the easiest way to distribute a project-specific OpenAI key with the bundle.

Note:

- The Windows build script intentionally uses `.venv-windows` so it does not conflict with a macOS `.venv` when the project is opened through a Parallels shared folder.

Important:

- The packaged app still needs a renderer on the recipient machine
- LibreOffice is used first when available
- On Windows, PowerPoint fallback is enabled by default in packaged or normal runs, so a user with Microsoft PowerPoint installed can still export even without LibreOffice
- On macOS, the app will also use Microsoft PowerPoint automatically if LibreOffice is missing and PowerPoint is installed
- Anyone who receives a build with `.env.local` inside it can extract and reuse that API key, so only do this for trusted internal colleagues.

## Packaging For Mac Users

If you want a Mac app bundle:

```bash
./scripts/build_macos_bundle.sh
```

That produces `dist/PDFtoPPTXReference.app` and you can zip it for sharing.

If `.env.local` exists in the project root when you build, the script copies it into `PDFtoPPTXReference.app/Contents/MacOS/.env.local` automatically.

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
- If LibreOffice is missing but Microsoft PowerPoint is installed, the app will use PowerPoint automatically on macOS
- Because the app is not notarized, macOS may show a Gatekeeper warning the first time it is opened
- Anyone who receives a build with `.env.local` inside it can extract and reuse that API key, so only do this for trusted internal colleagues.
