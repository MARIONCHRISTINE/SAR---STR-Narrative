"""
STR / SAR narrative generator
-----------------------------
Reads every Airtel Money statement workbook (.xlsx) in a folder and, for each
statement (e.g. Subs1.xlsx), creates Output\\Subs1\\ containing:
  1. Subs1.xlsx       - the statement with a narrative for every row in column AA
  2. AML-2026-.docx   - the case report (customer profile, transaction summary)
                        built from the Word template "CUSTOMER PROFILE sample.docx"

Needs:  pip install openpyxl python-docx

Usage (from VS Code terminal):
    python str_narratives.py                  -> processes the folder this script is in
    python str_narratives.py "C:\\path\\to\\folder"
    python str_narratives.py Subs1.xlsx Subs2.xlsx   -> specific files

Daily routine: drop new statements into the folder and run the script.
  - Each statement that is processed successfully is MOVED to the "Processed"
    folder, so the next run only picks up new statements (nothing is redone).
  - Existing output folders are never overwritten. If Output\\Subs1 already
    exists, the new one is Output\\Subs1 (2), then Subs1 (3), and so on.
  - If a statement fails, it stays where it is so you can fix it and run again.
  - If there are no new statements, the script just says so and stops.
The contents of the original statements are never changed.

To add or change a sentence, edit the build_narrative() function below.
Each transaction type is its own small "if" block.
"""

import copy
import json
import re
import shutil
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

try:
    from docx import Document
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
    from openpyxl.worksheet.table import TableColumn
except ModuleNotFoundError as missing:
    package = {"docx": "python-docx"}.get(missing.name, missing.name)
    sys.exit(
        f"\nMissing library: {package}\n"
        f"Install the libraries this script needs by running this command, then try again:\n\n"
        f'    "{sys.executable}" -m pip install openpyxl python-docx\n'
    )

# ---------------------------------------------------------------- settings
OUTPUT_COLUMN = "AA"          # where the narrative goes
OUTPUT_HEADER = "NARRATIVE"
OUTPUT_FOLDER = "Output"      # results go to Output\<statement name>\
PROCESSED_FOLDER = "Processed"  # finished statements are moved here ("" = leave them in place)
TEMPLATE_FILE = "CUSTOMER PROFILE sample.docx"   # Word template: next to this script, else template\ folder
REPORT_NAME = "AML-2026-.docx"                   # add the case number manually after 2026-
KYC_STATUS = "Complete"
OCCUPATION_PLACEHOLDER = "OCCUPATION"            # filled in manually
SHOW_CENTS = True           # False = "Kes 4,034" (like the Excel formula); True = "Kes 4,033.95"

# Column letters, same as the Excel formula. If a column moves, change it here only.
COL = {
    "txn_id": "A",        # TRANSFER_ID
    "time": "C",          # TRANSFER_ON
    "service": "F",       # SERVICE_TYPE
    "amount": "G",        # AMOUNT
    "party_msisdn": "H",  # SECOND_PARTY_ACCOUNT_ID
    "party_cat": "I",     # SECOND_PARTY_CATEGORY_CODE
    "post_bal": "M",      # POST_BAL
    "partner": "N",       # PARTNER_DATA (JSON)
    "direction": "O",     # TRANSACTION_TYPE (MR = received, MP = sent)
    "msisdn": "D",        # ACCOUNT_ID (customer MSISDN)
    "first_name": "P",    # FIRST_NAME (customer)
    "last_name": "Q",     # LAST_NAME
    "reg": "R",           # PRIMARY_REGISTRATION
    "dob": "S",           # PRIMARY_DOB
    "id": "T",            # PRIMARY_ID
    "p_first": "U",       # FIRST_NAME2 (other party)
    "p_last": "V",        # LAST_NAME2
    "p_reg": "W",         # SECONDARY_REGISTRATION
    "p_dob": "X",         # SECONDARY_DOB
    "p_id": "Y",          # SECONDARY_ID
}

BULLET = "\u2022"


# ---------------------------------------------------------------- helpers
def kes(value):
    """Format an amount like Excel TEXT(x,"#,##0")."""
    if value in (None, ""):
        return "N/A"
    try:
        d = Decimal(str(value))
    except Exception:
        return str(value)
    if SHOW_CENTS and d != d.to_integral_value():
        return f"{d.quantize(Decimal('0.01'), ROUND_HALF_UP):,.2f}"
    return f"{d.quantize(Decimal('1'), ROUND_HALF_UP):,.0f}"


def as_date(value):
    if isinstance(value, datetime):
        return value
    if value in (None, ""):
        return None
    for fmt in ("%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%d-%b-%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(value).strip(), fmt)
        except ValueError:
            pass
    return None


def fmt_date(value):
    d = as_date(value)
    return d.strftime("%d-%b-%y") if d else "N/A"


def fmt_txn_time(value):
    d = as_date(value)
    if not d:
        return str(value or "N/A")
    return d.strftime("%d-%b-%y") + " at " + d.strftime("%I:%M%p")   # 08-Nov-26 at 05:53PM


def text(value):
    """Cell value as clean text (whole numbers without .0)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split())       # trims and collapses double spaces


def partner_field(raw, key):
    """Read a field from the PARTNER_DATA JSON (falls back to a text search)."""
    if not raw:
        return ""
    try:
        data = json.loads(raw)
        if isinstance(data, dict) and data.get(key) not in (None, ""):
            return str(data[key]).strip()
    except (ValueError, TypeError):
        pass
    m = re.search(r'"%s"\s*:\s*"([^"]*)"' % re.escape(key), str(raw))
    return m.group(1).strip() if m else ""


# ---------------------------------------------------------------- narrative
def build_narrative(r):
    """r = dict of the row values keyed by COL names."""
    service = text(r["service"]).upper()
    party_cat = text(r["party_cat"]).upper()
    direction = text(r["direction"]).upper()
    partner = r["partner"]

    when = fmt_txn_time(r["time"])
    customer = (text(r["first_name"]).split(" ") or [""])[0]
    amt = kes(r["amount"])
    bal = kes(r["post_bal"])
    tid = text(r["txn_id"])
    party_name = f"{text(r['p_first'])} {text(r['p_last'])}".strip()
    head = f"On {when}, {customer}"
    line = f"{BULLET} Kes {amt} - (Transaction ID - {tid}) - "

    party_details = (
        f"{line}{party_name} of ID NO: {text(r['p_id'])} and MSISDN {text(r['party_msisdn'])}"
        f", was born on {fmt_date(r['p_dob'])}"
        f", and got registered on Airtel Money on {fmt_date(r['p_reg'])}"
    )

    if service == "IMTIN":
        sender = f"{partner_field(partner, 'senderFirstName')} {partner_field(partner, 'senderLastName')}".strip()
        return (f"{head} received an amount of Kes {amt} via IMT Thunes, bringing the balance to Kes {bal}"
                f" with details as shown below:\n"
                f"{line}IMT incoming transaction from {sender or 'an unnamed sender'} via THUNES.")

    if service == "SUBSCASHOUT":
        return (f"{head} cashed out an amount of Kes {amt} through an Airtel Money Agent leaving the balance at"
                f" Kes {bal}, with details shown below:\n{party_details}")

    if service == "SUBSCASHIN":
        return (f"{head} received a cash deposit of Kes {amt} from an Airtel Money Agent bringing the balance to"
                f" Kes {bal}, with details shown below:\n{party_details}")

    if service == "TXN_ROLLBACK":
        return (f"{head} received an automated system reversal of Kes {amt} bringing the balance to Kes {bal},"
                f" with details shown below:\n{line}Automated System Reversal.")

    if service == "BONUSDISB":
        return (f"{head} received an amount of Kes {amt} from Airtel Money Bonus on Transaction Disbursement"
                f" Wallets bringing the balance to Kes {bal}, with details shown below:\n"
                f"{line}Airtel Money Bonus on Transaction Disbursement Wallets.")

    if service == "SEW2W":
        return (f"{head} moved an amount of Kes {amt} from one Airtel Money wallet to another Airtel Money wallet,"
                f" with details shown below:\n{party_details}")

    if service == "BUNSLF":
        return (f"{head} bought bundles worth Kes {amt} leaving the balance at Kes {bal},"
                f" with details shown below:\n{line}{party_name}")

    if service == "P2PIOPIN":
        return (f"{head} received an amount of Kes {amt} from a Safaricom User bringing the balance to Kes {bal},"
                f" with details shown below:\n{line}{party_name}")

    # --- NEW (not in the Excel formula): bank-to-wallet, e.g. Equity Disbursement
    if service == "SUBSB2WNTB":
        ref = partner_field(partner, "referenceNumber")
        return (f"{head} received an amount of Kes {amt} from {party_name} (bank to Airtel Money wallet)"
                f" bringing the balance to Kes {bal}, with details shown below:\n"
                f"{line}{party_name}" + (f" - Bank Reference {ref}." if ref else "."))

    # --- NEW (not in the Excel formula): Paybill / bank bill payments
    if service == "BILLPAYSLF":
        ref = partner_field(partner, "referenceNumber")
        return (f"{head} paid an amount of Kes {amt} to {party_name}"
                + (f" account {ref}" if ref else "")
                + f" leaving the balance at Kes {bal}, with details shown below:\n"
                f"{line}{party_name}" + (f" - Paybill/Account No. {ref}." if ref else "."))

    if party_cat == "SUBS":
        return (f"{head} sent an amount of Kes {amt} to an Airtel Money User leaving the balance at {bal},"
                f" with the details shown below:\n{party_details}")

    msisdn = partner_field(partner, "subscriberMsisdn")
    if msisdn:
        verb = "received" if direction == "MR" else "sent"
        return (f"{head} {verb} an amount of Kes {amt} to a Safaricom user with details shown below:\n"
                f"{line}Sent to Safaricom user of MSISDN {msisdn} via MPESA.")

    if "lipa na mpesa via airtel money" in text(r["p_first"]).lower():
        return (f"{head} sent an amount of Kes {amt} to a Safaricom Till, with the details shown below:\n"
                f"{line}{text(r['p_first'])}")

    verb = "received" if direction == "MR" else "sent"
    return f"{head} {verb} an amount of Kes {amt} with the Airtel Money Transaction ID {tid}\n{party_details}"


# ---------------------------------------------------------------- case report (Word)
def short_date(d):
    """1 Apr ’26 (review-period style used in the template)."""
    return f"{d.day} {d:%b} ’{d:%y}"


def summarise(rows):
    """Customer details + credit/debit totals from all statement rows."""
    first = rows[0]
    credits = [Decimal(str(r["amount"] or 0)) for r in rows if text(r["direction"]).upper() == "MR"]
    debits = [Decimal(str(r["amount"] or 0)) for r in rows if text(r["direction"]).upper() == "MP"]
    dates = [d for d in (as_date(r["time"]) for r in rows) if d]
    dob, reg = as_date(first["dob"]), as_date(first["reg"])
    id_no = text(first["id"])

    full_upper = f"{text(first['first_name'])} {text(first['last_name'])}".strip().upper()
    period = f"{short_date(min(dates))} – {short_date(max(dates))}" if dates else "N/A"
    return {
        "Full Name": ": " + full_upper.title(),
        "MSISDN": ": " + text(first["msisdn"]),
        "ID Type / Number": ": " + ("National ID – " if id_no.isdigit() else "ID – ") + id_no,
        "DOB": ": " + (f"{dob.month}/{dob.day}/{dob.year}" if dob else "N/A"),
        "Registration Date": ": " + (reg.strftime("%d-%b-%y").upper() if reg else "N/A"),
        "Declaration Occupation": ": " + OCCUPATION_PLACEHOLDER,
        "KYC Status": ": " + KYC_STATUS,
        "TRANSACTION SUMMARY": f"TRANSACTION SUMMARY (Review Period: {period})",
        "Total Credits": f" - KES {kes(sum(credits))} ({len(credits)} transactions)",
        "Total Debits": f" - KES {kes(sum(debits))} ({len(debits)} transactions)",
    }


def _set_value(paragraph, value):
    """Keep the bold label run(s); put value in the first normal run; drop the rest."""
    runs = paragraph.runs
    idx = next((i for i, r in enumerate(runs) if not r.bold), None)
    if idx is None:                       # no value run yet: copy the label's run and un-bold it
        new = copy.deepcopy(runs[-1]._r)
        runs[-1]._r.addnext(new)
        runs = paragraph.runs
        idx = len(runs) - 1
        runs[idx].bold = False
        runs[idx].underline = False
    runs[idx].text = value
    for r in runs[idx + 1:]:
        r._r.getparent().remove(r._r)


def write_case_report(template: Path, dest: Path, info):
    doc = Document(template)
    for p in doc.paragraphs:
        label = "".join(r.text for r in p.runs if r.bold).strip()
        if label.startswith("TRANSACTION SUMMARY"):
            p.runs[0].text = info["TRANSACTION SUMMARY"]
            for r in p.runs[1:]:
                r._r.getparent().remove(r._r)
        elif label in info:
            _set_value(p, info[label])
    dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dest)


# ---------------------------------------------------------------- workbook handling
def include_narrative_in_table(ws):
    """Statements come as an Excel Table (A:Z). Stretch it (and any filter) to cover the
    narrative column, so sorting/filtering moves the narrative together with its row."""
    out_idx = column_index_from_string(OUTPUT_COLUMN)
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if max_col >= out_idx or max_col != out_idx - 1:
            continue                      # already includes AA, or is not next to it
        table.ref = f"{get_column_letter(min_col)}{min_row}:{OUTPUT_COLUMN}{max_row}"
        if table.autoFilter is not None:
            table.autoFilter.ref = table.ref
        next_id = max(c.id for c in table.tableColumns) + 1
        table.tableColumns.append(TableColumn(id=next_id, name=OUTPUT_HEADER))
    if ws.auto_filter.ref:
        min_col, min_row, max_col, max_row = range_boundaries(ws.auto_filter.ref)
        if max_col == out_idx - 1:
            ws.auto_filter.ref = f"{get_column_letter(min_col)}{min_row}:{OUTPUT_COLUMN}{max_row}"


def process_file(src: Path, out_dir: Path, template: Path):
    wb = load_workbook(src)
    ws = wb.worksheets[0]
    ws[f"{OUTPUT_COLUMN}1"] = OUTPUT_HEADER
    ws.column_dimensions[OUTPUT_COLUMN].width = 110
    wrap = Alignment(wrap_text=True, vertical="top")

    done = errors = 0
    rows = []
    for row in range(2, ws.max_row + 1):
        if ws[f"A{row}"].value in (None, ""):
            continue
        values = {name: ws[f"{letter}{row}"].value for name, letter in COL.items()}
        rows.append(values)
        try:
            narrative = build_narrative(values)
        except Exception as exc:                      # never stop the whole file for one bad row
            narrative = f"ERROR building narrative: {exc}"
            errors += 1
        cell = ws[f"{OUTPUT_COLUMN}{row}"]
        cell.value = narrative
        cell.alignment = wrap
        done += 1

    include_narrative_in_table(ws)
    out_dir.mkdir(parents=True, exist_ok=True)
    wb.save(out_dir / src.name)

    report = None
    if rows and template.exists():
        report = out_dir / REPORT_NAME
        write_case_report(template, report, summarise(rows))
    return done, errors, report


def collect_files(args):
    if not args:
        args = [str(Path(__file__).resolve().parent)]
    files = []
    for a in args:
        p = Path(a)
        if p.is_dir():
            files += sorted(f for f in p.glob("*.xlsx") if not f.name.startswith("~$"))
        elif p.suffix.lower() == ".xlsx":
            files.append(p)
        else:
            print(f"Skipping {a} (not a folder or .xlsx file)")
    return files


def unused_path(path: Path):
    """Subs1 -> Subs1 (2) -> Subs1 (3) ... so nothing already there is overwritten."""
    if not path.exists():
        return path
    n = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def main():
    files = collect_files(sys.argv[1:])
    if not files:
        where = sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parent
        print(f"No new statements to process in {where}.\n"
              f"Put the .xlsx statements in that folder and run the script again.")
        return
    here = Path(__file__).resolve().parent
    template = here / TEMPLATE_FILE                     # your own template, if you keep one here
    if not template.exists():
        template = here / "template" / TEMPLATE_FILE    # blank template shipped with the repo
    if not template.exists():
        print(f"WARNING: Word template '{TEMPLATE_FILE}' not found - "
              f"only the Excel files will be produced.")
    ok = failed = 0
    for src in files:
        out_dir = unused_path(src.parent / OUTPUT_FOLDER / src.stem)
        try:
            done, errors, report = process_file(src, out_dir, template)
        except PermissionError:
            print(f"FAIL {src.name}: a file is open in Excel/Word - close it and run again.")
            failed += 1
            continue
        except Exception as exc:
            print(f"FAIL {src.name}: {exc}")
            failed += 1
            continue
        ok += 1
        note = f", {errors} row(s) with errors" if errors else ""
        extra = f" + {report.name}" if report else ""
        print(f"OK   {src.name}: {done} narratives{note}{extra} -> {out_dir}")

        if PROCESSED_FOLDER:
            dest = unused_path(src.parent / PROCESSED_FOLDER / src.name)
            dest.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.move(str(src), str(dest))
            except OSError:
                print(f"     (could not move {src.name} to {PROCESSED_FOLDER} - is it open in Excel? "
                      f"Move it yourself so it is not processed again.)")
    print(f"\nFinished: {ok} statement(s) done, {failed} failed.")


if __name__ == "__main__":
    main()
