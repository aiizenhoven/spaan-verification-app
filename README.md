# Spaan Verification Cost & Reconciliation App

A local app that ingests candidate verification "result report" documents, prices them
against the same business rules already established in the master workbook, and updates
the Verification Cost Analysis & Candidate Reconciliation spreadsheet.

## Setup (once)

```bash
python -m pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

This opens the app in your browser (usually http://localhost:8501).

## Using it

1. **Master workbook** — in the sidebar, upload your current
   `Verification Cost Analysis & Candidate Reconciliation.xlsx`.
2. **Ingest new reports** tab — drop in one or more result report files:
   - MIE `Captured Credentials.xlsx` exports
   - SAQA/NLRD PDF reports (you'll be asked whether each is a **bulk** or **individual**
     submission — bulk is R58 flat, individual is R82 flat)
   - UNISA R4C upload `.xlsx` files
   - MUT Candidates `.xlsx` files
3. Click **Parse & preview** — the app shows how many rows are genuinely new, how many
   are already in the workbook (skipped, so nothing gets double-counted), and the total
   new cost.
4. Click **Apply to workbook** to write the update, then **Download updated workbook**.
   The new file keeps every existing row untouched; Candidate Summary, Candidate
   Calculations, Afika's View and Total Spend Details are recalculated from the full
   ledger, using the same formula conventions as the original file.
5. **Browse workbook** tab — search/filter Candidate Summary and the raw Consolidated
   Results ledger, and see live totals (spend, profit/loss vs. DEL charge, duplicates).

## How pricing works

All rates, VAT treatment, institution surcharges (UJ, UP, Wits, NMMU, DOE Colleges, QCTO,
Exam Council Lesotho, SAESI) and the assumed Citizenship/PERSAL fallback rule are encoded
in `pricing_rules.py`, mirroring the "Pricing" sheet and its notes in the master workbook.
If a rate changes, update it there.

## Files

| File | Purpose |
|---|---|
| `pricing_rules.py` | Rates, VAT, institution surcharges |
| `parsers.py` | Reads each known report format into normalized records |
| `engine.py` | Dedup (already-ingested vs. new) + duplicate-verification flagging |
| `aggregate.py` | Builds Candidate Summary / Calculations / Afika's View / Total Spend Details |
| `workbook_writer.py` | Writes everything back into a copy of the master workbook |
| `app.py` | The Streamlit UI |

## Formats not yet covered

Only the four report formats above have dedicated parsers, built from real sample files.
A generic R4C upload file that routes qualifications via a "Source Acronym" field (rather
than being pre-split into a UNISA-only or MUT-only file) isn't handled yet — if you get one
of those, send a sample and it can be added the same way.
