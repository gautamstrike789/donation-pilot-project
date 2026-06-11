# Donation Pilot Project

Streamlit app for owner/BA donation entry backed by Google Sheets.

## Local run

```powershell
python -m pip install -r requirements.txt
python setup.py
python -m streamlit run app.py
```

## Deploy on Streamlit Community Cloud

This app supports Streamlit Cloud by reading Google Sheets settings and Google service account credentials from Streamlit secrets.

1. Push this repo to GitHub.
2. Create a Streamlit Community Cloud app from the repository and choose `app.py` as the entry point.
3. In Streamlit Cloud, add these secrets:

```toml
[google_sheets]
admin_sheet_id = "your-admin-sheet-id"
donations_sheet_id = "your-donations-sheet-id"

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "your-private-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "your-service-account@your-project.iam.gserviceaccount.com"
client_id = "your-client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/your-service-account%40your-project.iam.gserviceaccount.com"
```

4. Share the Google Sheets used by the app with the service account email.
5. Redeploy.

## Notes

- Local OAuth files (`client_secret.json`, `token.pickle`, `sheets_config.json`) still work for local development.
- On Streamlit Cloud, the app uses `st.secrets` instead of the local OAuth token file.
