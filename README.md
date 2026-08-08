# Donation Pilot Project

Streamlit app for owner/BA donation entry backed by Google Sheets.

## Local run

```powershell
python -m pip install -r requirements.txt
python setup.py
python setup_events.py     # one-time: adds the Events worksheet (Event dropdown source)
python -m streamlit run app.py
```

When SOD is "Events" the form offers a second dropdown of Event Names mapped to
the selected owner (from the Admin sheet's "Events" worksheet, uploaded by
`setup_events.py` from `EventDataForForm.xlsx`). When SOD is "Airport" it offers
a fixed list of airport cities. Either selection is saved into its own column
(`Event Name`, `Airport Name`) in the Donations sheet, right after `SOD`. The
preview step before Submit is an editable grid — fix a wrong entry or delete a
row before writing to the sheet.

## One-Off form

`app_oneoff.py` is a second, fully independent form with its own Admin sheet
("TMO Admin One-Off", uploaded from `TMO Admin One-Off.xlsx`) and its own
Donations sheet ("TMO Donations One-Off") with columns
`SigninDT, OWNCODE, BAName, BACode, Forms, Supports, ADS`. ADS (Supports ÷ Forms)
is calculated automatically, not entered by the user. It does not share any
sheet with the main donation form.

```powershell
python -m pip install -r requirements.txt
python setup.py            # if you haven't already (signs in with your Google account)
python setup_oneoff.py     # one-time: creates the Admin + Donations One-Off sheets
python -m streamlit run app_oneoff.py
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

### Deploying the One-Off form

Deploy `app_oneoff.py` as its own separate Streamlit Community Cloud app (same
GitHub repo, main file `app_oneoff.py`). It has its own independent secrets store
— add:

```toml
[google_sheets]
oneoff_admin_sheet_id = "your-oneoff-admin-sheet-id"
oneoff_donations_sheet_id = "your-oneoff-donations-sheet-id"

[gcp_service_account]
# same service account block as above
```

Share the "TMO Admin One-Off" sheet and the "TMO Donations One-Off" sheet with
that service account's email.

## Notes

- Local OAuth files (`client_secret.json`, `token.pickle`, `sheets_config.json`) still work for local development.
- On Streamlit Cloud, the app uses `st.secrets` instead of the local OAuth token file.
