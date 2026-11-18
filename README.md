# SAR / STR Narrative Generator

Turns Airtel Money statement exports (`.xlsx`) into transaction narratives and an AML case report for STR/SAR risk reporting.

For each statement (e.g. `Subs1.xlsx`) the script creates `Output\Subs1\` with:

| File | Contents |
|---|---|
| `Subs1.xlsx` | The statement with a narrative for every transaction in column **AA** |
| `AML-2026-.docx` | Case report: customer profile and transaction summary (review period, total credits and debits). Add the case number after `2026-`. |

## Setup

1. Install Python from python.org and tick **"Add Python to PATH"**.
2. Install the libraries:
   ```bash
   pip install -r requirements.txt
   ```
3. That's it. The Word template is included at `template\CUSTOMER PROFILE sample.docx` (placeholders only, no customer data). The `Output` and `Processed` folders are created automatically.

To change the report's look, edit that template in Word. You can also put your own `CUSTOMER PROFILE sample.docx` next to the script; it takes priority over the one in `template\`. The script copies the template's formatting and replaces only the values.

## Daily use

1. Copy the day's statements (`.xlsx`) into the script's folder.
2. Run:
   ```bash
   py str_narratives.py
   ```
3. Collect the results from `Output\`.

- Processed statements are moved to `Processed\`, so the next run only picks up new ones.
- Existing output is never overwritten: a repeat of `Subs1` goes to `Output\Subs1 (2)\`.
- If there are no new statements, the script says so and stops.
- A statement that fails (e.g. open in Excel) stays in place so you can rerun.

## Settings

At the top of `str_narratives.py`: `SHOW_CENTS`, `REPORT_NAME`, `KYC_STATUS`, `OCCUPATION_PLACEHOLDER`, output/processed folder names, and the `COL` column map. To add or change a transaction type's wording, edit `build_narrative()`.

## Data protection

`.gitignore` blocks `.xlsx`, `.docx`, `Output\` and `Processed\` so customer statements and reports are never committed. The only exception is the blank template in `template\`. Never put real customer details into that file.
