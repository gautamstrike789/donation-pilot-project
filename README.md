# Donation Pilot Project

Streamlit app for owner/BA donation entry backed by Google Sheets.

## Local run

```powershell
python -m pip install -r requirements.txt
python setup.py
python setup_events.py            # one-time: adds the Events worksheet (Event dropdown source)
python setup_airports.py          # one-time: adds the Airports worksheet (Airport dropdown source)
python setup_owner_passcodes.py   # one-time: adds the Passcode column to Owners
python setup_ba_joinee_date.py    # one-time: adds the New Joinee Date column to BAs
python -m streamlit run app.py
```

When SOD is "Events" the form offers a second dropdown of Event Names mapped to
the selected owner (from the Admin sheet's "Events" worksheet, uploaded by
`setup_events.py` from `EventDataForForm.xlsx`). When SOD is "Airport" it offers
the list from the Admin sheet's "Airports" worksheet — edit that sheet directly
(add/remove/rename a row) to change the dropdown, no code change needed. Either
selection is saved into its own column (`Event Name`, `Airport Name`) in the
Donations sheet, right after `SOD`. `OwnerName` is also saved next to `OWNCODE`,
looked up automatically from the Owners worksheet. The preview step before
Submit is an editable grid — fix a wrong entry or delete a row before writing
to the sheet.

### Owner passcode gate

The Owners worksheet has a `Passcode` column (added by `setup_owner_passcodes.py`).
It's self-service: the first time anyone picks an owner whose `Passcode` cell is
blank, the form asks them to set one (typed twice to confirm) and saves it to
that cell automatically — no admin step needed. From then on, picking that owner
shows a passcode box that blocks SignIn Date, BA, Donations, and the Owners Data
History view until the correct passcode is entered (once per browser session).
Clear an owner's `Passcode` cell yourself if you want the "set a passcode" prompt
to run again for them. New owners you add to the sheet get the same column
automatically.

### New BA joinee date

The BAs worksheet has a `New Joinee Date` column (added by `setup_ba_joinee_date.py`).
When someone adds a brand-new BA through the form's "Add a new BA" section, the
SignIn Date used for that submission is written into this column automatically.
It's for the Admin sheet's own record-keeping only — the form's BA dropdown still
only ever shows "BA Name · BA Code", unchanged.

## One-Off form

`app_oneoff.py` is a second, fully independent form with its own Admin sheet
("TMO Admin One-Off", uploaded from `TMO Admin One-Off.xlsx`) and its own
Donations sheet ("TMO Donations One-Off") with columns
`SigninDT, OWNCODE, BAName, BACode, Clients, Forms, Supports, ADS, No Forms`.
ADS (Supports ÷ Forms) is calculated automatically, not entered by the user.
It does not share any sheet with the main donation form. It has the same
SignIn Date gate, owner passcode self-service gate, BA joinee-date tracking,
"No Forms" (equivalent to the main form's "NO Production") checkbox, and
Owners Data History view (with per-date Excel download) as the main RG form
— see those sections above; the mechanics are identical, just Forms/Supports/
Clients in place of Amount/Age/SOD. A Passcode protects an OWNCODE as a whole
even though the same owner can appear multiple times in the Owners sheet with
a different Client (e.g. HAI, STC) — setting it once covers every Client row.

```powershell
python -m pip install -r requirements.txt
python setup.py                          # if you haven't already (signs in with your Google account)
python setup_oneoff.py                   # one-time: creates the Admin + Donations One-Off sheets
python setup_oneoff_owner_passcodes.py   # one-time: adds the Passcode column to Owners
python setup_oneoff_ba_joinee_date.py    # one-time: adds the New Joinee Date column to BAs
python setup_oneoff_no_forms_column.py   # one-time: adds the No Forms column to Donations One-Off
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
