# Job Tracker Sync

Automatically updates your job application Excel tracker by scanning your Outlook inbox for rejection, interview, and new application emails.

## What it does

- Reads company names from your Excel tracker
- Scans your Outlook inbox for emails from those companies
- Detects **rejection emails** and sets status to `Declined`
- Detects **interview invitations** and sets status to `Interview`
- Detects **new application confirmations** and adds missing rows automatically
- Works with emails from ATS platforms (Greenhouse, Lever, Workday, iCIMS, etc.) by searching the email body for the company name

## Requirements

- Windows with Microsoft Outlook installed and logged in
- Python 3.7+

```
pip install openpyxl pywin32
python -m pywin32_postinstall -install
```

## Setup

1. Open `job_tracker_sync.py`
2. Set `TARGET_EMAIL` to your Outlook email address
3. Place the script in the same folder as `job_tracker.xlsx`, or set `DEFAULT_EXCEL` to the full path of your tracker
4. Adjust `SHEET_NAME`, `DATA_START_ROW`, `COL_COMPANY`, and `COL_STATUS` if your spreadsheet layout differs from the defaults

## Usage

```
python job_tracker_sync.py                  # run normally
python job_tracker_sync.py --dry-run        # preview changes without saving
python job_tracker_sync.py --debug          # show all emails and matched phrases
python job_tracker_sync.py --excel path/to/tracker.xlsx  # use a custom file path
```

## Excel format expected

| Column | Content |
|--------|---------|
| A | Company name |
| B | Role / position |
| C | Date applied |
| D | Status (`Applied` / `Interview` / `Declined`) |
| E | Notes |

Rows 1–4 are reserved for title, summary, and header. Data starts at row 5 by default.

## Notes

- The script only reads from Outlook — no credentials or API keys needed
- Company names are read dynamically from your Excel file on every run
- Rejections always override interviews if both are detected for the same company
- New applications not yet in the tracker are added automatically at the bottom and sorted on save
- After each save, rows are sorted: Interviews → Applied → Other → Declined, each group alphabetically
- Use `--debug` if no matches are found — it prints every email and flags which phrases triggered
