"""
Writes the updated ledger back into a copy of the master workbook:
  - Consolidated Results: existing rows untouched, new rows appended at the end.
  - Candidate Summary, Candidate Calculations, Afika's View: fully rebuilt from the
    merged ledger (existing + new), so they always stay internally consistent.
  - Total Spend Details: left as formulas; their fixed-range bounds are patched to
    cover the new row counts.
  - Pricing: untouched, except one traceability note appended.
"""
import re
import shutil
import copy
import datetime
import openpyxl
from openpyxl.utils import get_column_letter

import engine
import aggregate


def _copy_row_style(ws, from_row, to_row, n_cols):
    for c in range(1, n_cols + 1):
        src = ws.cell(row=from_row, column=c)
        dst = ws.cell(row=to_row, column=c)
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.number_format = src.number_format
        dst.alignment = copy.copy(src.alignment)


def _write_rows(ws, headers, rows, start_row, style_template_row=None):
    n_cols = len(headers)
    for i, row_dict in enumerate(rows):
        r = start_row + i
        if style_template_row:
            _copy_row_style(ws, style_template_row, r, n_cols)
        for c, h in enumerate(headers, start=1):
            ws.cell(row=r, column=c, value=row_dict.get(h))
    return start_row + len(rows) - 1 if rows else start_row - 1


def _clear_data_rows(ws, first_data_row, last_row, n_cols):
    for r in range(first_data_row, last_row + 1):
        for c in range(1, n_cols + 1):
            ws.cell(row=r, column=c, value=None)


def write_updated_workbook(master_path, output_path, new_consolidated_rows,
                            batch_source_names=None):
    """
    new_consolidated_rows: list of dicts (engine.CONSOLIDATED_HEADERS keys) to append,
        already deduped/flagged by engine.merge_batch (+ any fallback rows).
    Returns a summary dict describing what was written.
    """
    shutil.copy(master_path, output_path)
    wb = openpyxl.load_workbook(output_path)

    # ---- Consolidated Results: append new rows -------------------------------------
    ws_cr = wb["Consolidated Results"]
    cr_headers = [c.value for c in ws_cr[1]]
    last_row = ws_cr.max_row
    while last_row > 1 and ws_cr.cell(row=last_row, column=1).value is None:
        last_row -= 1
    new_start = last_row + 1
    _write_rows(ws_cr, cr_headers, new_consolidated_rows, new_start, style_template_row=2)
    new_last_row = new_start + len(new_consolidated_rows) - 1 if new_consolidated_rows else last_row

    # ---- Build the full merged ledger for the derived sheets -----------------------
    full_ledger = engine.load_ledger(output_path)  # re-read what we just wrote, incl. old rows
    # load_ledger reads from disk; our in-memory wb hasn't been saved yet, so build directly instead:
    full_ledger = []
    for row in ws_cr.iter_rows(min_row=2, max_row=new_last_row, values_only=False):
        vals = [c.value for c in row]
        if vals[0] is None:
            continue
        full_ledger.append(dict(zip(cr_headers, vals)))

    candidate_summary_rows = aggregate.build_candidate_summary(full_ledger)
    candidate_calc_rows = aggregate.build_candidate_calculations(full_ledger)
    afika_rows = aggregate.build_afika_view(full_ledger)
    total_spend = aggregate.build_total_spend_details(full_ledger, candidate_summary_rows)

    # ---- Candidate Summary: rebuild ------------------------------------------------
    ws_cs = wb["Candidate Summary"]
    cs_headers = [c.value for c in ws_cs[1]]
    n_cols_cs = len(cs_headers)
    old_max = ws_cs.max_row
    _clear_data_rows(ws_cs, 2, max(old_max, len(candidate_summary_rows) + 1), n_cols_cs)
    _write_rows(ws_cs, cs_headers, candidate_summary_rows, 2, style_template_row=2 if old_max >= 2 else None)
    # formulas for Total Verification Cost (AB) and Over/Under (AD)
    col = {h: i + 1 for i, h in enumerate(cs_headers)}
    AB = get_column_letter(col["Total Verification Cost (Incl. VAT)"])
    AC = get_column_letter(col["DEL Charge"])
    AD = get_column_letter(col["Over / Under"])
    cost_cols = [get_column_letter(col[h]) for h in [
        "Citizenship Cost (Incl. VAT)", "PERSAL Cost (Incl. VAT)", "Criminal Cost (Incl. VAT)",
        "Qualification Cost (Incl. VAT)", "Matric Cost (Incl. VAT)",
        "SAQA (Matric/NSC) Cost (Incl. VAT)", "SAQA (Tertiary) Cost (Incl. VAT)",
        "UNISA Cost (Incl. VAT)", "MUT Cost (Incl. VAT)",
    ]]
    for i in range(len(candidate_summary_rows)):
        r = 2 + i
        ws_cs.cell(row=r, column=col["Total Verification Cost (Incl. VAT)"],
                   value="=" + "+".join(f"{cc}{r}" for cc in cost_cols))
        ws_cs.cell(row=r, column=col["Over / Under"], value=f"={AC}{r}-{AB}{r}")
    new_cs_last = 1 + len(candidate_summary_rows)

    # ---- Candidate Calculations: rebuild ------------------------------------------
    ws_cc = wb["Candidate Calculations"]
    cc_headers = [c.value for c in ws_cc[1]]
    n_cols_cc = len(cc_headers)
    old_max_cc = ws_cc.max_row
    _clear_data_rows(ws_cc, 2, max(old_max_cc, len(candidate_calc_rows) + 1), n_cols_cc)
    col_cc = {h: i + 1 for i, h in enumerate(cc_headers)}
    H = get_column_letter(col_cc["Cost (Incl. VAT)"])
    r = 2
    idx = 0
    while idx < len(candidate_calc_rows):
        block_size = candidate_calc_rows[idx]["_block_size"]
        block_start, block_end = r, r + block_size - 1
        for j in range(block_size):
            row_dict = candidate_calc_rows[idx + j]
            rr = r + j
            if old_max_cc >= 2:
                _copy_row_style(ws_cc, 2, rr, n_cols_cc)
            for h in cc_headers:
                if h == "Candidate Total (Incl. VAT)":
                    ws_cc.cell(row=rr, column=col_cc[h], value=f"=SUM(${H}${block_start}:${H}${block_end})")
                else:
                    ws_cc.cell(row=rr, column=col_cc[h], value=row_dict.get(h))
        idx += block_size
        r += block_size
    new_cc_last = r - 1

    # ---- Afika's View: rebuild (static values) ------------------------------------
    ws_af = wb["Afika's View"]
    # header occupies rows 1-2 (grouped headers + column labels); data starts row 3
    af_col_labels = [c.value for c in ws_af[2]]
    n_cols_af = len(af_col_labels)
    old_max_af = ws_af.max_row
    _clear_data_rows(ws_af, 3, max(old_max_af, len(afika_rows) + 2), n_cols_af)
    for i, row_dict in enumerate(afika_rows):
        rr = 3 + i
        if old_max_af >= 3:
            _copy_row_style(ws_af, 3, rr, n_cols_af)
        for c, key in enumerate(aggregate.AFIKA_VIEW_HEADERS, start=1):
            ws_af.cell(row=rr, column=c, value=row_dict.get(key))
    new_af_last = 2 + len(afika_rows)

    # ---- Total Spend Details: patch fixed-range formula bounds --------------------
    ws_tsd = wb["Total Spend Details"]
    old_cs_bound = str(old_max)  # e.g. 3011, before this update
    old_cr_bound = str(new_start - 1) if new_consolidated_rows else None
    for row in ws_tsd.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                v = cell.value
                v = re.sub(r"(Candidate Summary'!\$?[A-Z]+2:\$?[A-Z]+)\d+", rf"\g<1>{new_cs_last}", v)
                v = re.sub(r"(Consolidated Results'!\$?[A-Z]+2:\$?[A-Z]+)\d+", rf"\g<1>{new_last_row}", v)
                if v != cell.value:
                    cell.value = v

    # ---- Pricing: traceability note -----------------------------------------------
    if batch_source_names:
        ws_pr = wb["Pricing"]
        note_row = ws_pr.max_row + 2
        ws_pr.cell(row=note_row, column=1,
                    value=f"App update {datetime.date.today().isoformat()}: merged {', '.join(batch_source_names)} "
                          f"({len(new_consolidated_rows)} new verification rows added).")

    wb.save(output_path)

    return {
        "new_rows_added": len(new_consolidated_rows),
        "total_verification_rows": new_last_row - 1,
        "total_candidates": len(candidate_summary_rows),
        "total_spend_details": total_spend,
        "output_path": output_path,
    }
