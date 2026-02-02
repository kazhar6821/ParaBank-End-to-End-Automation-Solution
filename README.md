# ParaBank Automation

Simple end-to-end automation for the ParaBank demo site.

## What it does
- Registers users from a file
- Opens bank accounts
- Requests a $10,000 loan with 20% down payment
- Converts results to EUR
- Generates a formatted Excel report
- Logs out and closes the browser safely

## Input
`ParaBank_users.csv` – list of customer profiles.

## Output
- `Parabank_Automation_Report.xlsx` (business-ready report)
- `screenshots/` (only if errors occur)

## How to run
```bash
pip install selenium pandas openpyxl python-dateutil webdriver-manager
python parabank_automation.py
