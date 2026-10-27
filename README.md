# HydroPipe — MS Pipe Design Website (prototype)

This is a locally runnable web app that uses the linked Google spreadsheet as its primary workbook source and the included `Design of MS-Pipe.xlsx` file as a fallback. It writes form values to the `Details` sheet, asks LibreOffice Calc to recalculate the linked formulas, reads outputs from `Final Report`, and generates a PDF plus a recalculated Excel file.

## Requirements

- Python 3.10+
- LibreOffice installed and available as `libreoffice` on your PATH
- Internet is not required at runtime for calculations; the Google Fonts import is only for optional visual styling.

## Run locally

### Windows / macOS / Linux

1. Install Python and LibreOffice.
2. Open a terminal in this folder.
3. Create a virtual environment (recommended) and install packages:

   ```bash
   python -m venv .venv
   # Windows: .venv\Scripts\activate
   # macOS/Linux: source .venv/bin/activate
   pip install -r requirements.txt
   ```

4. Start the app:

   ```bash
   python app.py
   ```

5. Open `http://127.0.0.1:5000` in your browser.

Generated reports and calculated workbooks are saved under `generated/`.

## Google Sheets access

The app downloads an XLSX export of the configured Google spreadsheet when a report is generated. Set the spreadsheet sharing permission to **Anyone with the link — Viewer** so the server can export it. If the online spreadsheet cannot be downloaded, the app falls back to the included workbook file. Google Sheets and Excel can interpret some formulas differently, so compare the results against Excel before relying on them.

## How the workbook is used

- `Details` is the input sheet; the website writes to its existing cells.
- The other sheets and formulas remain in the workbook copy.
- LibreOffice recalculates formulas before the app reads the `Final Report` sheet.
- The original workbook in this folder is not overwritten.

## Important before publishing online

This is a working local prototype, not a deployed public website. For public use, host it on a server that supports LibreOffice, use HTTPS, add authentication/rate limiting, configure data retention and backups, and have the engineering calculations reviewed and validated. The uploaded workbook includes formula features that may behave differently across Excel/LibreOffice versions. Test representative projects against desktop Excel and verify every design equation, unit, lookup table, and pass/fail check before using reports for construction, procurement, or approval. The generated PDF is a web-formatted report; the downloaded Excel file preserves the workbook's sheets and formulas.

## Troubleshooting

- If calculation fails, ensure `libreoffice --headless --version` works.
- If you change the workbook's input cell locations or sheet names, update `FIELDS` and the `Final Report` output mappings in `app.py`.


## Windows quick start

1. Install Python 3.10+ and LibreOffice.
2. Double-click `Start Website Windows.bat`.
3. Open `http://127.0.0.1:5000`.
4. If report generation says LibreOffice is missing, verify that `C:\Program Files\LibreOffice\program\soffice.exe` exists. The app checks the common 64-bit and 32-bit install locations and PATH.


## Deploy online

See [`DEPLOY_ONLINE.md`](DEPLOY_ONLINE.md). The project includes a Dockerfile that installs LibreOffice on the server and runs Flask through Gunicorn.
