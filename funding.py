"""
Daily funding sheet builder (sample project, synthetic data).

Reads the daily summary and detail files plus the client configuration file,
and produces:
  - the funding sheet (Excel), in the standard layout
  - one upload file per ACH bank template (ACH Collections, ACH Gross,
    ACH Interchange Withheld, ACH Net) for the bank portal
  - a list of wires, which are entered manually in the bank portal
  - an exceptions report
  - a reconciliation report proving all of the above tie out

Funding rules (from the real process):
  - NET clients:   credit = payments - chargebacks/returns. No interchange withheld.
  - GROSS clients: payments are credited in full (by ACH or wire, per config).
                   Chargebacks/returns are collected from the client separately (ACH).
                   A few high-volume clients (withhold_interchange = Y) also have
                   their daily interchange collected from their account as a
                   separate ACH collection, shown in its own section.

Layout rules:
  - Funding sheet sections, always in this order: ACH Gross Collection,
    ACH Gross Credit, ACH Gross Interchange Withheld, ACH Net, Wire.
  - Every client in the configuration file appears in every section that
    applies to it, every day, even with $0.00. Clients within a section follow
    the configuration file order.
  - Each bank upload file lists exactly the same lines, in the same order, as
    its funding sheet section. No header row; two columns: funding ID, amount.

Safety checks (nothing is dropped silently):
  1. Every funding ID in the data must exist in the configuration file.
     Unknown or blank funding IDs go to the exceptions report, not the funding sheet.
  2. The summary file must tie out to the detail file (counts and amounts).
  3. A net client whose chargebacks/returns exceed payments is HELD: $0.00 on
     the funding sheet and bank file, and listed on the exceptions report.
  4. Reconciliation: after writing, every output file is re-read from disk and
     proven against the funding sheet, the source data and the configuration.

Usage:
    python funding.py 20261120            # build all output files, then reconcile
    python funding.py 20261120 --verify   # re-check existing output files only

File names use the report date as YYYYMMDD, for example funding_sheet_20261120.xlsx.
"""

import csv
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook

BASE = Path(__file__).parent
ZERO = Decimal("0.00")
CENT = Decimal("0.01")

# Funding sheet sections, in the fixed order they always appear.
COLLECTION = "ACH Gross Collection"
GROSS_CREDIT = "ACH Gross Credit"
INTERCHANGE = "ACH Gross Interchange Withheld"
NET = "ACH Net"
WIRE = "Wire"
SECTION_ORDER = [COLLECTION, GROSS_CREDIT, INTERCHANGE, NET, WIRE]

# Funding sheet section -> bank portal template it is uploaded under.
# Collections and interchange are always ACH. Wires are not uploaded; they are
# entered manually in the bank portal.
BANK_TEMPLATES = {
    COLLECTION: "ACH_Collections",
    GROSS_CREDIT: "ACH_Gross",
    INTERCHANGE: "ACH_Interchange_Withheld",
    NET: "ACH_Net",
}


def parse_report_date(text: str) -> str:
    """Accept 20261120 or 2026-11-20. Returns the date as YYYY-MM-DD."""
    for fmt in ("%Y%m%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    raise SystemExit(f"Report date '{text}' must be YYYYMMDD, for example 20261120")


def stamp(day: str) -> str:
    """Date stamp used in every file name: YYYYMMDD."""
    return day.replace("-", "")


def read_csv(path: Path) -> list[dict]:
    with path.open(newline="") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, rows: list[dict], fields: list[str]) -> None:
    with path.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def to_money(value):
    """Convert a cell or CSV value to Decimal cents. Returns None if not a valid amount."""
    if value is None or value == "" or isinstance(value, bool):
        return None
    try:
        d = Decimal(repr(value)) if isinstance(value, float) else Decimal(str(value))
    except InvalidOperation:
        return None
    return d if d == d.quantize(CENT) else None


def load_config(path: Path) -> list[dict]:
    """Configuration rows in file order. The order drives the funding sheet layout."""
    rows = read_csv(path)
    seen = set()
    for r in rows:
        if r["funding_id"] in seen:
            raise SystemExit(f"Configuration error: funding ID {r['funding_id']} listed twice")
        seen.add(r["funding_id"])
        client_sections(r)  # validates the row
    return rows


def client_sections(client: dict) -> list[str]:
    """The funding sheet sections a configured client always appears in."""
    method = client["credit_method"]
    if method not in ("ACH", "Wire"):
        raise SystemExit(f"Configuration error: credit method '{method}' for {client['funding_id']}")
    if client["settlement_type"] == "Gross":
        sections = [GROSS_CREDIT if method == "ACH" else WIRE, COLLECTION]
        if client["withhold_interchange"] == "Y":
            sections.append(INTERCHANGE)
        return sections
    if client["settlement_type"] == "Net":
        return [NET if method == "ACH" else WIRE]
    raise SystemExit(f"Configuration error: settlement type '{client['settlement_type']}' "
                     f"for {client['funding_id']}")


def expected_layout(config: list[dict]) -> list[tuple[str, str]]:
    """(section, funding_id) in the exact order the funding sheet must show them."""
    return [(section, c["funding_id"]) for section in SECTION_ORDER
            for c in config if section in client_sections(c)]


def rule_amount(client: dict, section: str, payments: Decimal, cb: Decimal, ic: Decimal) -> Decimal:
    """The funding rule for one client in one section."""
    if section == COLLECTION:
        return cb
    if section == INTERCHANGE:
        return ic
    if client["settlement_type"] == "Net":
        return payments - cb
    return payments  # Gross credit, ACH or wire


def tie_out(summary: list[dict], detail: list[dict]) -> list[dict]:
    """Check that summary totals match the detail file for every client/funding ID."""
    from_detail = defaultdict(lambda: {"payment_count": 0, "payment_amount": ZERO,
                                       "cb_return_count": 0, "cb_return_amount": ZERO,
                                       "interchange_amount": ZERO})
    for r in detail:
        d = from_detail[(r["client_id"], r["funding_id"])]
        amount = Decimal(r["amount"])
        if r["transaction_type"] == "PAYMENT":
            d["payment_count"] += 1
            d["payment_amount"] += amount
            d["interchange_amount"] += Decimal(r["interchange_amount"])
        else:
            d["cb_return_count"] += 1
            d["cb_return_amount"] += amount

    problems = []
    seen = set()
    for s in summary:
        key = (s["client_id"], s["funding_id"])
        seen.add(key)
        d = from_detail.get(key)
        if d is None:
            problems.append({"client_id": key[0], "funding_id": key[1] or "(blank)",
                             "issue": "In summary file but not in detail file"})
            continue
        for field in ("payment_count", "cb_return_count"):
            if int(s[field]) != d[field]:
                problems.append({"client_id": key[0], "funding_id": key[1],
                                 "issue": f"{field}: summary {s[field]} vs detail {d[field]}"})
        for field in ("payment_amount", "cb_return_amount", "interchange_amount"):
            if Decimal(s[field]) != d[field]:
                problems.append({"client_id": key[0], "funding_id": key[1],
                                 "issue": f"{field}: summary {s[field]} vs detail {d[field]}"})
    for key in from_detail:
        if key not in seen:
            problems.append({"client_id": key[0], "funding_id": key[1] or "(blank)",
                             "issue": "In detail file but not in summary file"})
    return problems


def build_funding(summary: list[dict], config: list[dict]):
    src = defaultdict(list)
    for s in summary:
        src[(s["funding_id"] or "").strip()].append(s)
    configured = {c["funding_id"] for c in config}
    lines = {}       # (section, funding_id) -> line
    exceptions = []

    # Source data with no configured funding ID: never funded, always reported.
    for fid, rows in src.items():
        if fid not in configured:
            for s in rows:
                exceptions.append({
                    "client_id": s["client_id"], "funding_id": fid or "(blank)",
                    "amount": Decimal(s["payment_amount"]) - Decimal(s["cb_return_amount"]),
                    "issue": "Funding ID not in configuration. Not onboarded / no approved "
                             "bank account mapped. Do not fund until resolved.",
                })

    for client in config:
        fid = client["funding_id"]
        rows = src.get(fid, [])
        if len(rows) > 1:
            raise SystemExit(f"Data error: funding ID {fid} appears {len(rows)} times in summary file")
        s = rows[0] if rows else None
        held = False
        if s is not None and s["client_id"] != client["client_id"]:
            exceptions.append({"client_id": s["client_id"], "funding_id": fid,
                               "amount": Decimal(s["payment_amount"]) - Decimal(s["cb_return_amount"]),
                               "issue": f"Funding ID belongs to {client['client_id']} in configuration. "
                                        "Do not fund until resolved."})
            s, held = None, True

        payments = Decimal(s["payment_amount"]) if s else ZERO
        cb = Decimal(s["cb_return_amount"]) if s else ZERO
        ic = Decimal(s["interchange_amount"]) if s else ZERO

        if client["settlement_type"] == "Net" and payments - cb < ZERO:
            held = True
            exceptions.append({"client_id": client["client_id"], "funding_id": fid,
                               "amount": payments - cb,
                               "issue": "Net client: chargebacks/returns exceed payments. "
                                        "HELD for review before funding ($0.00 on funding sheet)."})

        for section in client_sections(client):
            amount = ZERO if held else rule_amount(client, section, payments, cb, ic)
            status = "HELD" if held else ("FUND" if amount != ZERO else "NO ACTIVITY")
            lines[(section, fid)] = {"section": section, "client_id": client["client_id"],
                                     "client_name": client["client_name"], "funding_id": fid,
                                     "amount": amount, "status": status}

    sheet = [lines[key] for key in expected_layout(config)]
    return sheet, exceptions


# ---------------------------------------------------------------- funding sheet (Excel)

def write_funding_sheet(path: Path, sheet: list[dict], day: str) -> None:
    """
    Layout (matches the standard funding sheet):
      A1 "Report Date:"  B1 report date
      For each section: section name in column A (first section row also has
      "Amount" in column D), one line per client (B name, C funding ID, D amount),
      then a Total row (C "Total", D =SUM of the section), then a blank row.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Funding Sheet"
    ws["A1"] = "Report Date:"
    ws["B1"] = date.fromisoformat(day)
    ws["B1"].number_format = "mm/dd/yyyy"
    totals = {}  # total row -> section total, stored as the formula's saved result
    row = 3
    for i, section in enumerate(SECTION_ORDER):
        ws.cell(row=row, column=1, value=section)
        if i == 0:
            ws.cell(row=row, column=4, value="Amount")
        row += 1
        first = row
        for line in (r for r in sheet if r["section"] == section):
            ws.cell(row=row, column=2, value=line["client_name"])
            ws.cell(row=row, column=3, value=line["funding_id"])
            ws.cell(row=row, column=4, value=line["amount"]).number_format = "#,##0.00"
            row += 1
        ws.cell(row=row, column=3, value="Total")
        total = f"=SUM(D{first}:D{row - 1})" if row > first else 0
        ws.cell(row=row, column=4, value=total).number_format = "#,##0.00"
        totals[row] = sum((r["amount"] for r in sheet if r["section"] == section), ZERO)
        row += 2
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 22.57
    ws.column_dimensions["C"].width = 6.71
    ws.column_dimensions["D"].width = 14
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates when the file is opened
    wb.save(path)
    store_formula_results(path, totals)


def store_formula_results(path: Path, totals: dict) -> None:
    """
    openpyxl saves formulas without a calculated result, so viewers that don't
    recalculate (Excel Protected View, phone and browser previews) show blank
    totals. This writes each Total formula's result into the file next to the
    formula, so the totals display everywhere. The formula itself is kept.
    """
    sheet_xml = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    xml = parts[sheet_xml].decode("utf-8")
    for row, value in totals.items():
        pattern = rf'(<c r="D{row}"[^>]*><f>[^<]*</f>)<v\s*/?>(?:</v>)?'
        xml, n = re.subn(pattern, rf"\g<1><v>{value}</v>", xml)
        if n != 1:
            raise SystemExit(f"Could not store the total for row {row} in the funding sheet")
    parts[sheet_xml] = xml.encode("utf-8")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=path.parent) as tmp:
        tmp_path = Path(tmp.name)
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    shutil.move(str(tmp_path), path)


def read_funding_sheet(path: Path) -> dict:
    """Parse the funding sheet back from disk. Records layout problems instead of failing."""
    ws = load_workbook(path).active
    saved = load_workbook(path, data_only=True).active  # formula results as stored in the file
    parsed = {"report_date": ws["B1"].value, "a1": ws["A1"].value, "amount_header": ws["D3"].value,
              "sections": [], "problems": []}
    current = None
    for r in range(3, ws.max_row + 1):
        a, b, c, d = (ws.cell(row=r, column=k).value for k in range(1, 5))
        if a is not None:
            current = {"name": a, "lines": [], "total_formula": None, "total_saved": None, "rows": []}
            parsed["sections"].append(current)
            if b is not None or c is not None:
                parsed["problems"].append(f"Row {r}: section row has extra values")
        elif c == "Total":
            if current is None:
                parsed["problems"].append(f"Row {r}: Total row outside a section")
                continue
            current["total_formula"] = d
            current["total_saved"] = to_money(saved.cell(row=r, column=4).value)
            current = None
        elif b is not None or c is not None or d is not None:
            if current is None:
                parsed["problems"].append(f"Row {r}: client line outside a section")
                continue
            current["lines"].append({"row": r, "client_name": b, "funding_id": c,
                                     "amount_raw": d,
                                     # Must be a real number: Excel's SUM skips text cells
                                     "amount": to_money(d) if isinstance(d, (int, float))
                                     and not isinstance(d, bool) else None})
    return parsed


# ---------------------------------------------------------------- bank upload files

def write_bank_file(path: Path, lines: list[tuple]) -> None:
    """Bank upload format: no header row, two columns: funding ID, amount."""
    with path.open("w", newline="") as f:
        csv.writer(f).writerows(lines)


def read_bank_file(path: Path) -> list[dict]:
    with path.open(newline="") as f:
        rows = list(csv.reader(f))
    return [{"funding_id": r[0] if r else "", "amount": r[1] if len(r) > 1 else "",
             "columns": len(r)} for r in rows]


def write_bank_files(sheet: list[dict], out_dir: Path, day: str):
    """One upload file per ACH template, same lines and order as its sheet section."""
    totals = {}
    for section, template in BANK_TEMPLATES.items():
        lines = [(r["funding_id"], r["amount"]) for r in sheet if r["section"] == section]
        write_bank_file(out_dir / f"bank_upload_{template}_{stamp(day)}.csv", lines)
        totals[template] = (len(lines), sum((a for _, a in lines), ZERO))
    wires = [r for r in sheet if r["section"] == WIRE]
    return totals, wires


# ---------------------------------------------------------------- reconciliation

def reconcile(day: str, out_dir: Path, summary: list[dict], detail: list[dict],
              config: list[dict]) -> list[dict]:
    """
    Prove the output files before anything is uploaded. Everything is re-read
    from disk, so this checks the files exactly as they will be used.

      L. Layout: the funding sheet has the report date, the five sections in
         order, every configured client in every section that applies to it
         (configuration file order), valid amounts, and Total rows that sum
         exactly their own section.
      A. Every client/funding ID in the source data is accounted for: on the
         funding sheet, or on the exceptions report.
      B. Every funding sheet amount equals the funding rule applied to the
         source data ($0.00 if HELD).
      C. Detail file totals = summary file totals.
      D. Each bank upload file matches its funding sheet section line by line
         and in order (funding ID, amount), plus format, line count and total.
      E. Funding sheet total = bank upload files + wires.
    """
    results = []

    def check(name, scope, expected, actual):
        results.append({"check": name, "scope": scope, "expected": expected,
                        "actual": actual, "result": "PASS" if expected == actual else "FAIL"})

    fs = read_funding_sheet(out_dir / f"funding_sheet_{stamp(day)}.xlsx")
    exceptions = read_csv(out_dir / f"exceptions_{stamp(day)}.csv")
    bank = {t: read_bank_file(out_dir / f"bank_upload_{t}_{stamp(day)}.csv") for t in BANK_TEMPLATES.values()}
    cfg = {c["funding_id"]: c for c in config}
    src = {(s["client_id"], (s["funding_id"] or "").strip() or "(blank)"): s for s in summary}
    excepted = {(e["client_id"], e["funding_id"]) for e in exceptions if e["amount"] != ""}

    # L. Layout
    rd = fs["report_date"]
    check("L. Report date", "funding sheet", day,
          rd.date().isoformat() if isinstance(rd, datetime) else str(rd))
    check("L. Header labels", "funding sheet", ("Report Date:", "Amount"),
          (fs["a1"], fs["amount_header"]))
    check("L. No stray rows", "funding sheet", [], fs["problems"])
    check("L. Sections in fixed order", "funding sheet", SECTION_ORDER,
          [s["name"] for s in fs["sections"]])
    sheet_layout = [(s["name"], line["funding_id"]) for s in fs["sections"] for line in s["lines"]]
    check("L. All configured clients, fixed order", "funding sheet", expected_layout(config), sheet_layout)
    for s in fs["sections"]:
        for line in s["lines"]:
            c = cfg.get(line["funding_id"])
            scope = f"{line['funding_id']} / {s['name']}"
            check("L. Client name matches configuration", scope,
                  c["client_name"] if c else "configured client", line["client_name"])
            check("L. Amount is a valid number (2 decimals)", scope, True, line["amount"] is not None)
        rows = [line["row"] for line in s["lines"]]
        expected_total = f"=SUM(D{rows[0]}:D{rows[-1]})" if rows else 0
        check("L. Total row sums exactly its section", s["name"], expected_total, s["total_formula"])
        check("L. Total shown in file = sum of section lines", s["name"],
              sum((line["amount"] or ZERO for line in s["lines"]), ZERO), s["total_saved"])

    sheet = [{"section": s["name"], "funding_id": line["funding_id"], "amount": line["amount"]}
             for s in fs["sections"] for line in s["lines"]]

    # A. Every source key accounted for
    on_sheet = {(cfg[r["funding_id"]]["client_id"], r["funding_id"]) for r in sheet if r["funding_id"] in cfg}
    for key in src:
        check("A. Source accounted for", f"{key[0]} / {key[1]}", True, key in on_sheet or key in excepted)

    # B. Each amount = rule applied to source
    for r in sheet:
        c = cfg.get(r["funding_id"])
        if c is None:
            continue
        key = (c["client_id"], r["funding_id"])
        s = src.get(key)
        payments = Decimal(s["payment_amount"]) if s else ZERO
        cb = Decimal(s["cb_return_amount"]) if s else ZERO
        ic = Decimal(s["interchange_amount"]) if s else ZERO
        expected = ZERO if key in excepted else rule_amount(c, r["section"], payments, cb, ic)
        check("B. Amount = funding rule applied to source", f"{r['funding_id']} / {r['section']}",
              expected, r["amount"])

    # C. Detail file -> summary file control totals
    det = {"payments": ZERO, "cb_returns": ZERO, "interchange": ZERO}
    for r in detail:
        if r["transaction_type"] == "PAYMENT":
            det["payments"] += Decimal(r["amount"])
            det["interchange"] += Decimal(r["interchange_amount"])
        else:
            det["cb_returns"] += Decimal(r["amount"])
    sums = {"payments": sum((Decimal(s["payment_amount"]) for s in summary), ZERO),
            "cb_returns": sum((Decimal(s["cb_return_amount"]) for s in summary), ZERO),
            "interchange": sum((Decimal(s["interchange_amount"]) for s in summary), ZERO)}
    for k in det:
        check("C. Detail file total = summary file total", k, det[k], sums[k])

    # D. Bank files vs funding sheet: same lines, same order
    for section, template in BANK_TEMPLATES.items():
        sheet_lines = [(r["funding_id"], r["amount"]) for r in sheet if r["section"] == section]
        file_lines = [(r["funding_id"], to_money(r["amount"])) for r in bank[template]]
        check("D. Bank file format: 2 columns, no header", template, True,
              all(r["columns"] == 2 and to_money(r["amount"]) is not None for r in bank[template]))
        check("D. Bank file lines and order = funding sheet section", template, sheet_lines, file_lines)
        check("D. Bank file line count = funding sheet section", template, len(sheet_lines), len(file_lines))
        check("D. Bank file total = funding sheet section total", template,
              sum((a or ZERO for _, a in sheet_lines), ZERO), sum((a or ZERO for _, a in file_lines), ZERO))

    # E. Funding sheet total = bank files + wires
    sheet_total = sum((r["amount"] or ZERO for r in sheet), ZERO)
    bank_total = sum((to_money(r["amount"]) or ZERO for rows in bank.values() for r in rows), ZERO)
    wire_total = sum((r["amount"] or ZERO for r in sheet if r["section"] == WIRE), ZERO)
    check("E. Funding sheet total = bank files + wires", "all", sheet_total, bank_total + wire_total)

    write_csv(out_dir / f"reconciliation_{stamp(day)}.csv",
              [{**r, "expected": str(r["expected"]), "actual": str(r["actual"])} for r in results],
              ["check", "scope", "expected", "actual", "result"])
    return results


def print_reconciliation(results: list[dict]) -> bool:
    failed = [r for r in results if r["result"] == "FAIL"]
    print(f"Reconciliation checks: {len(results) - len(failed)} of {len(results)} passed")
    for r in failed:
        print(f"  FAIL {r['check']} [{r['scope']}]: expected {r['expected']}, got {r['actual']}")
    if failed:
        print("*** RECONCILIATION FAILED - DO NOT UPLOAD BANK FILES ***")
    else:
        print("Reconciliation PASSED - bank files match the funding sheet and source data")
    return not failed


def main() -> None:
    args = sys.argv[1:]
    verify_only = "--verify" in args
    args = [a for a in args if a != "--verify"]
    if len(args) != 1:
        print("Usage: python funding.py YYYYMMDD [--verify]")
        sys.exit(1)
    day = parse_report_date(args[0])

    summary = read_csv(BASE / "data" / f"summary_{stamp(day)}.csv")
    detail = read_csv(BASE / "data" / f"detail_{stamp(day)}.csv")
    config = load_config(BASE / "config" / "clients.csv")
    out_dir = BASE / "output"

    if verify_only:
        ok = print_reconciliation(reconcile(day, out_dir, summary, detail, config))
        sys.exit(0 if ok else 2)

    tie_problems = tie_out(summary, detail)
    sheet, exceptions = build_funding(summary, config)
    for p in tie_problems:
        exceptions.append({**p, "amount": ""})

    out_dir.mkdir(exist_ok=True)
    write_funding_sheet(out_dir / f"funding_sheet_{stamp(day)}.xlsx", sheet, day)
    write_csv(out_dir / f"exceptions_{stamp(day)}.csv", exceptions, ["client_id", "funding_id", "amount", "issue"])

    print(f"Funding sheet for {day}")
    print("-" * 60)
    for section in SECTION_ORDER:
        rows = [r for r in sheet if r["section"] == section]
        total = sum((r["amount"] for r in rows), ZERO)
        print(f"{section:<40}{len(rows):>3} lines {total:>12,.2f}")
    print("-" * 60)

    bank_totals, wires = write_bank_files(sheet, out_dir, day)
    print("Bank upload files (ACH):")
    for template, (count, amount) in bank_totals.items():
        print(f"  {template:<38}{count:>3} lines {amount:>12,.2f}")
    print("Wires to enter manually in the bank portal:")
    for w in wires:
        print(f"  {w['funding_id']:<8}{w['client_name']:<28}{w['status']:<12}{w['amount']:>12,.2f}")
    print("-" * 60)
    print(f"Summary/detail tie-out: {'PASS' if not tie_problems else 'FAIL'}")
    print(f"Exceptions needing review: {len(exceptions)}")
    for e in exceptions:
        print(f"  - {e['client_id']} / {e['funding_id']}: {e['issue']}")
    print("-" * 60)

    ok = print_reconciliation(reconcile(day, out_dir, summary, detail, config))
    sys.exit(0 if ok and not tie_problems else 2)


if __name__ == "__main__":
    main()
