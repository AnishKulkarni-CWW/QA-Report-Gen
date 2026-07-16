# QA Work Hours Analytics Dashboard

A single-page Streamlit dashboard that turns your QA hours tracker workbook into a
management-ready view: KPI summary, Donut, Waffle, Range and Bar/Column charts,
per-QA cards, and Excel/PDF export — filterable by Daily / Monthly / Yearly.

## What it expects in the Excel file

The app scans **every sheet** in the uploaded workbook (skipping `Master Data`,
`Config`, `Instructions`) and auto-detects the header row on each one. It looks for
these columns (any reasonable naming variant is matched automatically):

- QA Name (or it uses the sheet name if the column is blank)
- Date
- Day (auto-derived from Date if missing)
- Month (auto-derived from Date if missing)
- Billable Hours
- Non-Billable Hours
- Hours Not Worked
- Total Hours (auto-recalculated if blank/zero but the three components exist)

It automatically:
- Skips blank separator rows and monthly "Total" summary rows
- Skips Saturday/Sunday placeholder rows
- Normalizes inconsistent name spellings (e.g. `Saujanya.Gouda` vs `Saujanya Gouda` → one person)
- Merges all QA sheets into one clean dataset

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`),
and upload your `.xlsx` file.

## Step 2 — Packaging into a Windows .exe / Mac .dmg

Streamlit apps aren't compiled directly into an exe — you package the whole
Python app + Streamlit runtime using **PyInstaller** or, more reliably for
Streamlit specifically, a small wrapper approach. Two proven paths:

### Option A — `stlite` (Streamlit compiled to WebAssembly, no Python needed)
Best if you eventually want a fully standalone desktop app with no installer
dependency on Python at all. Works well for dashboards like this one without
heavy backend logic. You'd package the `stlite` desktop build with
[`stlite desktop`](https://github.com/whitphx/stlite) which wraps it in an
Electron shell producing native `.exe` / `.dmg` installers.

### Option B — PyInstaller + a launcher script (keeps full Python/Pandas power)
Recommended here since this app uses pandas/openpyxl/reportlab which are easiest
to keep as regular Python:

1. Create `run_app.py`:
   ```python
   import sys, os
   from streamlit.web import cli as stcli

   def main():
       sys.argv = ["streamlit", "run", os.path.join(os.path.dirname(__file__), "app.py"),
                   "--server.headless=false", "--global.developmentMode=false"]
       sys.exit(stcli.main())

   if __name__ == "__main__":
       main()
   ```
2. Build:
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --add-data "app.py:." run_app.py
   ```
3. On Windows this produces a `.exe` in `dist/`; on macOS it produces a Unix
   executable you then wrap with `create-dmg` or Platypus to get a `.dmg`.

Either way, keep developing and testing the dashboard here in Streamlit first —
the app code itself (`app.py`) does not need to change for either packaging
route, only the launcher/build step differs.

## Files

- `app.py` — the full dashboard application
- `requirements.txt` — Python dependencies
