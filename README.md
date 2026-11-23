# Daily Funding & Balancing Automation (Sample Project)

A simplified rebuild of a daily merchant funding process I redesigned in payments operations. **All client names, IDs and amounts in this repository are made up.** No real company or customer data is used.

## The problem

For about 30 years, daily funding at the company was done by hand. Staff read spreadsheets and ledgers, calculated amounts and typed the numbers in. It took about 1.5 hours a day, and every manual step was a chance for human error.

## How the process evolved

| Stage | Tool | Daily time | What changed |
|---|---|---|---|
| Original | Ledgers and spreadsheets, by hand | ~1.5 hours | Manual calculation and keying |
| 1 | Excel | ~30 minutes | Raw data in, one standardized, repeatable funding sheet out |
| 2 | Alteryx | ~30 minutes | Same logic rebuilt in Alteryx. Worked, but needed a paid subscription and was harder to change |
| 3 | Python (written with Claude) | ~10 minutes | Same logic as a script. Free to run, easy to update, can be scheduled to run each morning |

Before switching to the Python script, I ran it side by side with the Alteryx workflow for two months and compared results every day. During that time the script caught errors that Alteryx had missed, including the one in the case study below.

## Funding rules

| Client type | How the day's funding is calculated |
|---|---|
| **Net** | Credit = payments − chargebacks/returns. No interchange withheld. |
| **Gross** | Payments are credited in full (ACH or wire). Chargebacks/returns are collected from the client separately. |
| **Gross, interchange withheld** | A few high-volume clients ran $50,000+ in interchange per day. Rather than float that for a month until it was recouped, it was collected daily. Payments are still credited in full, and the interchange is a separate collection from the client's account, shown in its own section of the funding sheet. |

Each client's type, credit method and interchange setting live in the configuration file, not in the code.

## Inputs

I had the development team build two daily data files, and I built a configuration file:

1. **Summary file.** For each report day, totals by client and funding ID: payments, chargebacks/returns and interchange withheld, each with a count and an amount. Includes the payment date.
2. **Detail file.** The same data at the individual payment level, with a date/time stamp on each payment.
3. **Configuration file.** The client list, each client's funding ID and other relevant client details. It also controls how the output is laid out, for example:
   - **Gross clients**: which ones chargebacks/returns are collected from, and which are credited by ACH or by wire
   - **Net clients**: a separate section

**Need-to-know by design:** bank account numbers never appear in the payment files. A funding ID links to the actual bank accounts, which were kept separately. Only payments operations needed to see account numbers.

## Output

1. **Funding sheet (Excel).** One standardized, easy-to-read sheet. It's meant to be the first document anyone checks during funding research: *What did we say we funded today, and does it match what the bank says we funded?*
2. **Bank upload files.** One CSV for each funding template set up with the bank, ready to upload in the bank portal to create the day's ACH payments and collections:
   - **ACH Collections**: chargebacks/returns collected from Gross clients
   - **ACH Gross**: full payment credits to Gross clients
   - **ACH Interchange Withheld**: daily interchange collected from high-volume Gross clients
   - **ACH Net**: net credits to Net clients
3. **Wire list.** Wires couldn't be uploaded, so they were entered manually in the bank portal. The script lists each one with its amount.
4. **Exceptions report.** Anything held back from funding for review (see the checks below).
5. **Reconciliation report.** The result of every check run before upload.

### Funding sheet layout

```
Report Date:   11/20/2026

ACH Gross Collection                                       Amount
               Lakeside Utilities Co       F-1001          336.82
               Northfield Water District   F-1002          127.46
               Summit Property Group       F-1003        5,111.35
               Pine Valley Medical         F-1004            0.00
                                           Total         5,575.63

ACH Gross Credit
               Lakeside Utilities Co       F-1001        6,728.58
               Northfield Water District   F-1002        2,548.08
                                           Total         9,276.66

ACH Gross Interchange Withheld
               Summit Property Group       F-1003        3,087.74
                                           Total         3,087.74

ACH Net
               Riverbend HOA Services      F-2001        4,395.62
               Cedar Creek Dental          F-2002        1,462.74
               Harbor Point Rentals        F-2003            0.00
               Maple Grove Storage         F-2004            0.00
                                           Total         5,858.36

Wire
               Summit Property Group       F-1003      140,352.28
               Pine Valley Medical         F-1004        9,062.23
                                           Total       149,414.51
```

### Layout rules

- **Same sections, same order, every day:** ACH Gross Collection, ACH Gross Credit, ACH Gross Interchange Withheld, ACH Net, Wire.
- **Every configured client, every day.** Every client in the configuration file appears in every section that applies to it, even with $0.00. Within a section, clients follow the configuration file order.
- **Collections and interchange are always ACH.** Only credits can go by wire.
- **Totals are live Excel formulas** (`=SUM(...)`) covering exactly their own section. The calculated result is also saved in the file, so totals display even in viewers that don't recalculate, like Excel's Protected View or a phone preview.
- **Bank files mirror the funding sheet.** Each bank upload file has exactly the same lines, in the same order, as its funding sheet section, including $0.00 lines.
- **Bank files are kept simple.** No header row, just two columns: funding ID and amount. For example, the ACH Net file:

  ```
  F-2001,4395.62
  F-2002,1462.74
  F-2003,0.00
  F-2004,0.00
  ```

The bank account details stay with each template in the bank portal, not in these files.

**File names** use the report date as YYYYMMDD, for example `funding_sheet_20261120.xlsx` and `bank_upload_ACH_Net_20261120.csv`.

## Funding research: three levels

1. **Funding sheet vs. bank.** Does the funded amount match the bank?
2. **Raw data.** If there's still a gap, check the summary and detail files against both the bank and the funding sheet.
3. **Backend.** If everything matches and the client still disputes the amount, work with the development team to review the backend tables and scripts.

Problems that reach Level 3 are usually an unusual payment scenario. Either it had never happened before, or it was a known possibility the company chose not to build for. Once it happens, it gets one of two fixes:
- a permanent fix in the backend, or
- a new check that catches it before that day's funding goes out.

## Case study: the payments with no funding ID

A client expanded its business and asked for new bank accounts to hold the new business's funds. The project manager told the development team but not payments operations.

When development finished, the Python script started producing funding amounts with **no funding ID attached**. The payments were real, but no approved bank account was mapped to them yet. When I investigated, I found the new accounts still needed the standard underwriting checks before they could be fully onboarded.

**The Python script surfaced these payments. The Alteryx workflow ignored them.**

Lessons:
- **Never drop data silently.** Every payment should either map to a funding ID or show up as an exception for someone to review.
- **Operations belongs in change management.** A change to how money moves needs payments operations involved before it goes live, not after.

## How AI was used

I designed the process, the business rules and the output. Claude wrote the Python code from that design. I tested it, ran it alongside the existing process and confirmed the results before relying on it.

AI has made it possible for almost anyone willing to do the due diligence to work like a junior developer. That opens the door for operations people to learn more about the technical side.

## Checks built into the sample script

Nothing is dropped silently. Anything that fails a check goes to the exceptions report:

1. **Unmapped funding ID.** A funding ID in the data that isn't in the configuration file (the case study above). It is never funded.
2. **Summary vs. detail tie-out.** Counts and amounts in the summary file must match the detail file.
3. **Negative net.** A net client whose chargebacks/returns exceed its payments is HELD: it stays on the funding sheet and bank file at $0.00 and is listed on the exceptions report for review.

## Reconciliation before upload

The bank upload files must match the funding sheet and all the data used to produce them. After writing the files, the script **re-reads them from disk**, including the Excel funding sheet, and proves:

| Check | What it proves |
|---|---|
| **L. Layout** | The funding sheet has the right report date, the five sections in order, every configured client in every section that applies to it (in configuration order), a real number for every amount, Total formulas that sum exactly their own section, and a saved total that equals the section's lines. |
| **A. Everything accounted for** | Every client/funding ID in the source data is on the funding sheet or on the exceptions report. |
| **B. Amounts match the rules** | Every funding sheet amount equals the funding rule applied to the source data ($0.00 if held). |
| **C. Control totals** | Detail file totals = summary file totals, for payments, chargebacks/returns and interchange. |
| **D. Bank files match the sheet** | Each bank upload file matches its funding sheet section line by line and in the same order (funding ID and amount), plus file format (two columns, no header), line count and total. |
| **E. Everything routed** | Funding sheet total = bank upload files + wires. |

Results are saved to `output/reconciliation_YYYYMMDD.csv`. If any check fails, the script prints **DO NOT UPLOAD BANK FILES** and exits with an error code, so a scheduled run can't quietly pass a bad file.

You can also re-run only the checks against existing files, for example right before uploading, to catch any edit made after the files were created:

```
python funding.py 20261120 --verify
```

These scenarios were tested by editing the output files on purpose. Each one was caught:

- **Funding sheet:**
  - an amount edited
  - a Total formula changed to skip a row
  - a $0.00 client row deleted
  - two clients swapped
  - the wrong report date
  - an amount entered as text, which Excel's SUM would silently skip
- **Bank files:**
  - an amount changed by $0.27
  - a line deleted
  - an unapproved funding ID added
  - a $0.00 line removed
  - two lines swapped
  - a header row added
  - an extra column added
- **Exceptions report:** the unmapped funding ID removed, leaving payments unaccounted for

## Run it yourself

Requires Python 3.9+ and the `openpyxl` package (for the Excel funding sheet):

```
pip install -r requirements.txt
python make_sample_data.py              # creates the synthetic data files in data/
python funding.py 20261120              # builds all output files, then reconciles them
python funding.py 20261120 --verify     # re-checks existing output files only
```

Sample console output:

```
Funding sheet for 2026-11-20
------------------------------------------------------------
ACH Gross Collection                      4 lines     5,575.63
ACH Gross Credit                          2 lines     9,276.66
ACH Gross Interchange Withheld            1 lines     3,087.74
ACH Net                                   4 lines     5,858.36
Wire                                      2 lines   149,414.51
------------------------------------------------------------
Bank upload files (ACH):
  ACH_Collections                         4 lines     5,575.63
  ACH_Gross                               2 lines     9,276.66
  ACH_Interchange_Withheld                1 lines     3,087.74
  ACH_Net                                 4 lines     5,858.36
Wires to enter manually in the bank portal:
  F-1003  Summit Property Group       FUND          140,352.28
  F-1004  Pine Valley Medical         FUND            9,062.23
------------------------------------------------------------
Summary/detail tie-out: PASS
Exceptions needing review: 2
  - C101 / F-1099: Funding ID not in configuration. Not onboarded / no approved bank account mapped. Do not fund until resolved.
  - C202 / F-2003: Net client: chargebacks/returns exceed payments. HELD for review before funding ($0.00 on funding sheet).
------------------------------------------------------------
Reconciliation checks: 82 of 82 passed
Reconciliation PASSED - bank files match the funding sheet and source data
```

## Repository contents

- `make_sample_data.py`: generates the synthetic summary and detail files
- `funding.py`: builds the funding sheet, bank upload files, wire list and exceptions report, then reconciles them
- `requirements.txt`: Python package needed (openpyxl)
- `config/clients.csv`: sample client configuration (no bank account numbers)
- `data/`: synthetic summary and detail files for one report day
- `output/`: the resulting funding sheet (Excel), four bank upload files, exceptions report and reconciliation report

## Disclaimer

This is a personal portfolio project. It is not affiliated with or endorsed by any former employer, and it contains no real company or customer data.
