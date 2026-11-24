# Ship Report Builder

Cross-platform (macOS + Windows) desktop app that reads an **OK-TO-SHIP** report
plus a folder of raw test reports, then produces a shipping report that matches a
**sample template** exactly — preserving all formatting and formulas.

## Inputs (app window)
1. **Sample report** — template whose format + formulas are cloned.
2. **PO Number** — optional.
3. **Build Code** — optional.
4. **Raw-data folder** — directory of unique-ID-coded test reports.
5. **OK-TO-SHIP report** — Excel workbook with tabs:
   - `Log file verification` → IMEIs
   - `Count of Station` → full list of unique devices
   - `Balance status` → units NOT shipping
6. **Save report as** — output path.

### Device lookup
Enter a device ID (IMEI/DevEUI) in **Look up device** and click **Look Up** to see
that device's full test history (every attempt: operator, station, time, result,
fail reason) pulled straight from the raw files. PO Number / Build Code narrow the
search if given. Instant, because each raw file is named after its device id.

## How it works
The app copies the sample file (keeping every formula and style), then injects the
device data — marking each unit as shipping / not-shipping based on `Balance status`.

The UI runs in your **web browser** (a tiny local server; no data leaves your
machine). This replaced the old Tk window, which renders black on the system
Python's Tk 8.5. `src/app.py` (Tk) is kept only as a fallback.

## Run (development)
```bash
# macOS  — opens the UI in your browser
.venv/bin/python src/webapp.py
# Windows
.venv\Scripts\python src\webapp.py
```

## Set up on another machine
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt     # Windows: .venv\Scripts\pip
```

## Inspect a workbook's structure (dev)
```bash
.venv/bin/python src/inspect_xlsx.py path/to/file.xlsx
```

## Standalone app

### macOS (already built)
Double-click **`dist/ShipReportBuilder.app`**. No terminal or Python needed.
To rebuild after code changes:
```bash
.venv/bin/pyinstaller --noconfirm --clean ShipReportBuilder.spec
```

### Windows (build on a Windows PC)
```bat
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller --windowed --noconfirm --clean ^
  --name "ShipReportBuilder" --paths src --collect-submodules openpyxl --collect-data tzdata src\app.py
```
Produces `dist\ShipReportBuilder\ShipReportBuilder.exe`.

## Hosted web version (reach it from any device)
`src/web_app.py` is a Flask UI over the same `engine`/`analysis`/`filler` core,
for deploying somewhere reachable from any PC — unlike `webapp.py`, it takes
file **uploads** (the raw-data folder as a single `.zip`) instead of native OS
file dialogs, and the generated report comes back as a **download**.

Gated by one shared password, set via the `SHIPREPORT_PASSWORD` environment
variable on the host — **never commit a real password to the repo**. If that
variable isn't set, the app refuses to serve anything (so it can't
accidentally run open to the internet).

### Deploy to Render.com
1. Push this repo to GitHub (see below).
2. In Render: **New → Blueprint**, pick this repo. Render reads `render.yaml`
   and creates the service automatically.
3. Render will prompt for the `SHIPREPORT_PASSWORD` secret during setup (not
   stored in the repo) — set it to whatever password you want to share with
   your team. `FLASK_SECRET_KEY` is auto-generated.
4. Once deployed, the service URL is reachable from any device, anywhere.

Free-tier services spin down after ~15 min idle — the next request after that
takes ~30s to cold-start. Upgrade the plan on Render if you need it always-on.

### Run it locally first (optional, to test before deploying)
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements-web.txt
SHIPREPORT_PASSWORD=mypassword .venv/bin/python src/web_app.py
# open http://127.0.0.1:5000
```

### Push this repo to GitHub
```bash
git add -A
git commit -m "Initial commit"
gh repo create ShipReportApp --private --source=. --remote=origin
git push -u origin main
```
(No `gh` CLI? Create an empty repo at github.com/new, then
`git remote add origin <url>` and `git push -u origin main`.)

## Modules
- `src/raw_parser.py` — parse per-device files (multiple JSON attempts each)
- `src/analysis.py`   — filter by PO/Build Code, aggregate all metrics
- `src/ok_to_ship.py` — read OK-TO-SHIP Balance status (held-back units)
- `src/filler.py`     — write computed data into a copy of the sample template
- `src/lookup.py`     — look up one device by ID (reads only its named file)
- `src/engine.py`     — orchestrates the above
- `src/app.py`        — Tkinter GUI
- `src/webapp.py`     — local browser UI with native file dialogs (desktop app)
- `src/web_app.py`    — hosted browser UI with file uploads (deploy anywhere)

## Status
- [x] Raw JSON attempt parser (validated on real data)
- [x] PO/Build-Code filtering + per-device aggregation (matches sample)
- [x] OK-TO-SHIP Balance-status parsing
- [x] Fill all 9 tabs preserving formatting, formulas, charts
- [x] GUI wired + end-to-end tested
- [x] macOS standalone app packaged
