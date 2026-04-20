# Fleet CAN Compatibility Web App

This package turns your current workbook-based matcher into a browser app.

## What users do
1. Open the web page.
2. Paste one vehicle per line in the format: `Brand Model Year`.
3. Click **Run Compatibility Check**.
4. View summary, supported vehicles, unsupported vehicles.
5. Download the final Excel report.

## Files
- `app.py` - the Streamlit web application
- `requirements.txt` - Python dependencies
- `Dockerfile` - container deployment option
- `assets/Vehicle_Compatibility_Matcher_Focused_Report_Final.py` - matcher logic
- `assets/Vehicle_Compatibility_Workbench_Focused_Report_Final.xlsx` - workbook template

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy online with minimal effort

### Option 1: Streamlit Community Cloud
1. Unzip this package.
2. Upload the folder to a GitHub repository.
3. In Streamlit Community Cloud, create a new app from that repository.
4. Set the main file to `app.py`.
5. Deploy.

### Option 2: Docker / Render / Railway / Azure / AWS
Build and run the container:
```bash
docker build -t fleet-can-app .
docker run -p 8501:8501 fleet-can-app
```
Then deploy the same container to your hosting platform.

## Important note
This package is ready to deploy, but a public URL still requires a hosting account. I cannot publish it to the internet from here without access to your hosting credentials.
