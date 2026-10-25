# Private developer setup

Use a personal computer and personal accounts. Do not use company equipment, network resources or confidential documents. Windows is the intended maintainer setup; this initial execution/QA occurred on Linux, so Windows commands are guidance rather than tested Windows results.

Install Node >=22.18.0, Python, LibreOffice, and a reviewed Poppler build for `pdftoppm` on the personal machine. Install the public Noto Sans CJK TC font for reproducible Chinese fixture rendering. Use a project-local Python virtual environment.

```powershell
npm ci
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
$env:SLIDERECAST_TEST_FONT = 'C:\path\to\NotoSansCJKtc-Regular.otf'
$env:SLIDERECAST_SOFFICE = 'C:\Program Files\LibreOffice\program\soffice.exe'
$env:SLIDERECAST_PDFTOPPM = 'C:\path\to\pdftoppm.exe'
npm run build
npm run fixtures
npm run benchmark
node scripts/edit-proof.mjs
npm test
npm run package
python scripts/clean-smoke.py
```

`npm ci` installs exact lockfile dependencies; upgrading them is a deliberate change requiring new output and rendering checks. The two JavaScript libraries have separate bundling entry points. Preview/Sharp remain developer dependencies. The clean smoke test removes access to project node_modules from the candidate execution path and imports only extracted bundle files.

The developer fixture generator currently imports PyMuPDF to create original PDFs. Skill helper scripts do not. Do not redistribute external binaries/fonts or change project licensing without reviewing the relevant terms. Full schemas and Microsoft UI tests are separate from ZIP/XML checks.

For rendering-only verification, call `python skills/image-to-editable-ppt/scripts/render.py input.pptx output-folder`. `SLIDERECAST_SOFFICE` and `SLIDERECAST_PDFTOPPM` make tool paths portable. The original PPTX is hashed and kept intact. Changing rasterizers changes RGB benchmark results, so compare measurements only with the same rendering method.
