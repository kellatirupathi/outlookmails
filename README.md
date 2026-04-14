# Outlook Desktop Mailer

Windows desktop app for sending or drafting Outlook emails from a locally signed-in Outlook desktop account.

## What it does

- Uses the Outlook desktop client on the same PC
- Lets you pick the Outlook sender account
- Accepts recipient data as CSV text or imported CSV files
- Supports placeholders like `{name}`, `{email}`, `{company}` in the subject and body
- Saves reusable templates in `%LOCALAPPDATA%\OutlookDesktopMailer\templates.json`
- Can create drafts first or send directly
- Supports the same attachments for every generated mail

## Prerequisites

- Windows
- Python 3
- Microsoft Outlook desktop installed on the same machine
- Outlook already signed in with the account you want to send from

If Outlook desktop is not installed, account loading and sending will fail.

## Run

```powershell
python .\outlook_desktop_mailer.py
```

## Build EXE

To create a portable Windows executable:

```powershell
.\build_release.ps1
```

After the build:

- EXE: `release\OutlookDesktopMailer.exe`
- ZIP for GitHub Releases: `release\OutlookDesktopMailer-portable.zip`

The packaged app still requires:

- Windows
- Classic Outlook desktop installed
- Outlook signed in with the sender account

## Recipient CSV format

Paste or import CSV with a header row. Example:

```csv
name,email,company
John Doe,john@example.com,Acme
Jane Smith,jane@example.com,Globex
```

Then use placeholders in the template:

- Subject: `Welcome {name}`
- Body: `Hi {name}, welcome to {company}.`

## Notes

- `Create Drafts` saves the messages in Outlook drafts without sending.
- `Send Emails` sends immediately through the selected Outlook account.
- In the packaged EXE, templates are stored in `%LOCALAPPDATA%\OutlookDesktopMailer\templates.json` so user changes persist.
- If you need per-recipient attachments or richer automation rules, extend the CSV and app logic from here.
