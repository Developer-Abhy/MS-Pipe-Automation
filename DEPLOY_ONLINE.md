# Deploy HydroPipe online

This project now includes a Dockerfile with Python, Flask, Gunicorn and LibreOffice Calc. Users only need a browser; LibreOffice runs on the hosting server.

## Easiest path: Render (Docker)

1. Extract this ZIP.
2. Create a private GitHub repository and upload the project files to it. Include `Dockerfile`, `requirements.txt`, `app.py`, `templates/`, `static/`, and `Design of MS-Pipe.xlsx`.
3. In Google Sheets, set **Share → General access → Anyone with the link → Viewer** if you want the app to fetch the online sheet. Do not grant Editor access to a public app.
4. Sign in to Render and create a **New Web Service** connected to your GitHub repository.
5. Choose **Docker** as the runtime. Render should detect the included `Dockerfile`.
6. Deploy. Open the public URL Render provides and test report generation with a known design case.

`render.yaml` is included as a reference. You can also create the service manually in Render's dashboard.

## Important hosting notes

- This package is configured for Docker-based hosting and installs LibreOffice inside the container.
- The included setup is a starting deployment configuration, not a deployment performed on your behalf. You still need to create the hosting account and connect the repository.
- Generated files are stored in the service's local `generated/` directory. On many hosts, local disk is temporary and may be cleared on redeploy/restart. Add persistent storage and file cleanup/authentication before production use.
- The app currently uses the included workbook as a fallback if it cannot fetch the Google Sheets XLSX export. Ensure sheet permissions are set correctly.
- A publicly accessible calculator should use authentication or rate limiting if needed, appropriate retention controls, and secure handling of project/customer details.
- Compare workbook outputs against Microsoft Excel for representative cases before using the website for engineering decisions. Formula compatibility can differ between spreadsheet engines.
- The default Render plan may have costs and limitations that change; review current plan details before deploying.
