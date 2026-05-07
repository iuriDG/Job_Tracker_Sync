Job Tracker Sync

Automatically updates your job application Excel tracker by scanning your Outlook inbox for rejection and interview emails.
What it does

Reads company names from your Excel tracker
Scans your Outlook inbox for emails from those companies
Detects rejection emails and sets status to Declined
Detects interview invitations and sets status to Interview
Works with emails from ATS platforms (Greenhouse, Lever, iCIMS, etc.) by searching the email body for the company name

Requirements

Windows with Microsoft Outlook installed and logged in
Python 3.7+

pip install openpyxl pywin32
python -m pywin32_postinstall -install
Setup

Open job_tracker_sync.py
Set TARGET_EMAIL to your Outlook email address
Set DEFAULT_EXCEL to the path of your tracker, or place the script in the same folder as the Excel file
Adjust COL_COMPANY, COL_STATUS, SHEET_NAME, and DATA_START_ROW if your spreadsheet layout is different

Usage
python job_tracker_sync.py
python job_tracker_sync.py --dry-run   # preview changes without saving
python job_tracker_sync.py --debug     # show all emails and matched phrases
Excel format expected
ColumnContentACompany nameDStatus (Applied / Interview / Declined / Offer)
Notes

The script only reads from Outlook — no credentials or API keys needed
Company names are read dynamically from your Excel file on every run
Rejections always override interviews if both are detected for the same company
Use --debug if no matches are found — it prints every email and flags which phrases triggered
