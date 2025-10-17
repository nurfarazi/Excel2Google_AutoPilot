# Excel -> Google Sheets Automation

Python utility that clears a Google Sheet tab and repopulates it with data pulled from an Excel workbook. Configuration lives in a separate `.env` file so credentials never get committed.

## Setup

1. Ensure Python 3.11+ is available.
2. Create and activate a virtual environment:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
3. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```
4. Copy the sample environment file and fill in real values:
   ```powershell
   Copy-Item .env.example .env
   ```
   - `EXCEL_FILE_PATH` — path to the workbook that will be uploaded.
   - `GOOGLE_SERVICE_ACCOUNT_FILE` — service account JSON file (keep it outside version control).
   - `GOOGLE_SPREADSHEET_ID` — ID segment from the Google Sheets URL.
   - `GOOGLE_WORKSHEET_NAME` — tab name to target.
   - `GOOGLE_COLUMNS` (optional) — comma separated column names to export in order.

> Share the target Google Sheet with the service account email so the script can edit it.

## Usage

Run a dry run to validate configuration without touching the sheet:
```powershell
python src/excel_to_gsheet.py --dry-run --verbose
```

Run the full transfer:
```powershell
python src/excel_to_gsheet.py
```

### GUI Mode

Launch the desktop interface (built with Tkinter) for manual runs:
```powershell
python src/gui_app.py
```
- Fields are pre-filled from `.env` when present.
- Click **Dry Run** or **Run Transfer** to execute; logs stream in the window.
- Use **Save Config** to persist the current inputs back to `.env`.

## Scheduling (Windows Task Scheduler)

1. Create a basic task that runs `powershell.exe`.
2. Program/script: `powershell.exe`
3. Add arguments:
   ```text
   -ExecutionPolicy Bypass -File "D:\git\test\Excel2Google_AutoPilot\run.ps1"
   ```
4. Create `run.ps1` alongside the project (not tracked by git) with:
   ```powershell
   $env:Path = "$PSScriptRoot\.venv\Scripts;$env:Path"
   python "$PSScriptRoot\src\excel_to_gsheet.py"
   ```

Adjust trigger for the desired cadence, enable “Run whether user is logged on or not,” and configure failure notifications as needed.

## Next Steps

- Add alerting (email/Teams) for non-zero exit codes.
- Extend transformation logic if the Excel structure requires preprocessing.
- Add tests around custom data cleansing rules when implemented.
