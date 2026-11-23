"""
Generates synthetic daily data files for the sample project.

Creates two files for one report day, the way the development team's files worked:
  - a DETAIL file: one row per transaction, with a date/time stamp
  - a SUMMARY file: totals by client and funding ID, with a count and amount
    for payments, chargebacks/returns and interchange

Everything here is made up. The random seed is fixed so the files come out the
same every time.

Usage:
    python make_sample_data.py
"""

import csv
import random
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

REPORT_DATE = "2026-11-20"
DATA_DIR = Path(__file__).parent / "data"

random.seed(42)

# (client_id, funding_id, number_of_payments, avg_payment, chargebacks/returns, interchange_rate)
# F-1099 is the "case study" funding ID: the client added new bank accounts, the
# dev team turned them on, but payments operations was never told, so F-1099 is
# NOT in config/clients.csv.
PLAN = [
    ("C100", "F-1001", 40, 180, 2, Decimal("0.018")),
    ("C101", "F-1002", 25, 95, 1, Decimal("0.015")),
    ("C102", "F-1003", 60, 2400, 3, Decimal("0.022")),
    ("C103", "F-1004", 30, 310, 0, Decimal("0.019")),
    ("C200", "F-2001", 20, 220, 1, Decimal("0.017")),
    ("C201", "F-2002", 15, 140, 2, Decimal("0.018")),
    ("C202", "F-2003", 1, 120, 3, Decimal("0.018")),  # chargebacks exceed payments
    ("C101", "F-1099", 12, 210, 0, Decimal("0.015")),  # unmapped funding ID
]


def money(value) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    day_start = datetime.fromisoformat(REPORT_DATE)
    detail_rows = []
    txn_number = 1

    for client_id, funding_id, n_pay, avg, n_cb, ic_rate in PLAN:
        for _ in range(n_pay):
            amount = money(random.uniform(avg * 0.5, avg * 1.5))
            detail_rows.append({
                "transaction_id": f"T{txn_number:06d}",
                "client_id": client_id,
                "funding_id": funding_id,
                "transaction_type": "PAYMENT",
                "amount": amount,
                "interchange_amount": money(amount * ic_rate),
                "payment_date": REPORT_DATE,
                "timestamp": (day_start + timedelta(seconds=random.randint(0, 86399))).isoformat(),
            })
            txn_number += 1
        for _ in range(n_cb):
            amount = money(random.uniform(avg * 0.5, avg * 1.5))
            detail_rows.append({
                "transaction_id": f"T{txn_number:06d}",
                "client_id": client_id,
                "funding_id": funding_id,
                "transaction_type": random.choice(["CHARGEBACK", "RETURN"]),
                "amount": amount,
                "interchange_amount": money(0),
                "payment_date": REPORT_DATE,
                "timestamp": (day_start + timedelta(seconds=random.randint(0, 86399))).isoformat(),
            })
            txn_number += 1

    detail_rows.sort(key=lambda r: r["timestamp"])

    detail_path = DATA_DIR / f"detail_{REPORT_DATE.replace('-', '')}.csv"
    with detail_path.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(detail_rows[0].keys()))
        writer.writeheader()
        writer.writerows(detail_rows)

    # Build the summary from the detail so the two files tie out.
    summary = {}
    for r in detail_rows:
        key = (r["client_id"], r["funding_id"])
        s = summary.setdefault(key, {
            "client_id": r["client_id"], "funding_id": r["funding_id"], "payment_date": REPORT_DATE,
            "payment_count": 0, "payment_amount": money(0),
            "cb_return_count": 0, "cb_return_amount": money(0),
            "interchange_count": 0, "interchange_amount": money(0),
        })
        if r["transaction_type"] == "PAYMENT":
            s["payment_count"] += 1
            s["payment_amount"] += r["amount"]
            s["interchange_count"] += 1
            s["interchange_amount"] += r["interchange_amount"]
        else:
            s["cb_return_count"] += 1
            s["cb_return_amount"] += r["amount"]

    summary_path = DATA_DIR / f"summary_{REPORT_DATE.replace('-', '')}.csv"
    rows = sorted(summary.values(), key=lambda s: (s["client_id"], s["funding_id"]))
    with summary_path.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    print(f"Wrote {len(detail_rows)} detail rows to {detail_path}")
    print(f"Wrote {len(rows)} summary rows to {summary_path}")


if __name__ == "__main__":
    main()
