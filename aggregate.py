"""
Builds Candidate Summary, Candidate Calculations, Afika's View and a Total Spend
Details summary from the full Consolidated Results ledger (existing + newly added rows).
"""
from collections import OrderedDict
import pricing_rules as pr

CORE_CATEGORIES = ["Citizenship", "PERSAL", "Criminal", "Qualification", "Matric"]
ALL_COST_CATEGORIES = ["Citizenship", "PERSAL", "Criminal", "Qualification", "Matric",
                        "SAQA-Matric", "SAQA-Tertiary", "UNISA", "MUT"]

CANDIDATE_SUMMARY_HEADERS = [
    "Candidate Name", "South African ID Number",
    "Citizenship Verification", "Citizenship Supplier", "Citizenship Cost (Incl. VAT)",
    "PERSAL Verification", "PERSAL Supplier", "PERSAL Cost (Incl. VAT)",
    "Criminal Verification", "Criminal Supplier", "Criminal Cost (Incl. VAT)",
    "Qualification Verification", "Qualification Supplier", "Qualification Cost (Incl. VAT)",
    "Matric Verification", "Matric Supplier", "Matric Cost (Incl. VAT)",
    "SAQA (Matric/NSC) Verification", "SAQA (Matric/NSC) Cost (Incl. VAT)",
    "SAQA (Tertiary) Verification", "SAQA (Tertiary) Cost (Incl. VAT)",
    "UNISA Verification", "UNISA Cost (Incl. VAT)",
    "MUT Verification", "MUT Cost (Incl. VAT)",
    "Number of Duplicates", "Duplicate Verification Cost (Incl. VAT)",
    "Total Verification Cost (Incl. VAT)", "DEL Charge", "Over / Under", "Notes",
    "All 5 Core Verifications Complete?",
]

CANDIDATE_CALCULATIONS_HEADERS = [
    "Candidate Name", "South African ID Number", "Verification Type", "Supplier",
    "Verification Result", "Cost (Excl. VAT)", "VAT (15%)", "Cost (Incl. VAT)",
    "Duplicate?", "Calculation / Reason", "Candidate Total (Incl. VAT)",
]

AFIKA_VIEW_HEADERS = [
    "Name", "ID",
    "Citizenship MIE", "Citizenship Datanamix", "Citizenship Total", "Citizenship Yes(#)", "Citizenship No",
    "Persal MIE", "Persal DPSA", "Persal Total", "Persal Yes(#)", "Persal No",
    "Matric MIE", "Matric SAQA", "Matric Total", "Matric Yes(#)", "Matric No",
    "Qualification MIE", "Qualification SAQA", "Qualification UNISA", "Qualification MUT",
    "Qualification Total", "Qualification Yes(#)", "Qualification No",
    "Criminal MIE", "Criminal Total", "Criminal Yes(#)", "Criminal No",
    "DEL Cost", "Actual Cost", "Difference",
]


def group_by_candidate(all_rows):
    """all_rows: list of Consolidated Results dict rows (existing + new), in ledger order."""
    grouped = OrderedDict()
    for row in all_rows:
        sa_id = str(row.get("South African ID Number") or "").strip()
        if not sa_id:
            continue
        if sa_id not in grouped:
            grouped[sa_id] = {"name": row.get("Candidate Name"), "rows": []}
        grouped[sa_id]["rows"].append(row)
    return grouped


def build_candidate_summary(all_rows):
    grouped = group_by_candidate(all_rows)
    summary_rows = []
    for sa_id, info in grouped.items():
        rows = info["rows"]
        by_cat = {cat: [r for r in rows if r.get("Category") == cat] for cat in ALL_COST_CATEGORIES}

        def verif(cat):
            n = len(by_cat[cat])
            return f"Yes ({n})" if n else "No"

        def supplier(cat):
            suppliers = sorted({str(r.get("Verification Supplier") or "") for r in by_cat[cat] if r.get("Verification Supplier")})
            return "/".join(suppliers) if suppliers else None

        def cost(cat):
            return round(sum(float(r.get("Cost (Incl. VAT)") or 0) for r in by_cat[cat]), 2)

        dup_rows = [r for r in rows if str(r.get("Duplicate?")) == "Yes"]
        n_dup = len(dup_rows)
        dup_cost = round(sum(float(r.get("Cost (Incl. VAT)") or 0) for r in dup_rows), 2)

        total = round(sum(cost(cat) for cat in ALL_COST_CATEGORIES), 2)
        del_charge = pr.DEL_CHARGE_PER_CANDIDATE
        over_under = round(del_charge - total, 2)

        all_core = all(len(by_cat[c]) > 0 for c in CORE_CATEGORIES)
        notes = f"{n_dup} duplicate verification(s) included in total." if n_dup else None

        summary_rows.append({
            "Candidate Name": info["name"], "South African ID Number": sa_id,
            "Citizenship Verification": verif("Citizenship"), "Citizenship Supplier": supplier("Citizenship"),
            "Citizenship Cost (Incl. VAT)": cost("Citizenship"),
            "PERSAL Verification": verif("PERSAL"), "PERSAL Supplier": supplier("PERSAL"),
            "PERSAL Cost (Incl. VAT)": cost("PERSAL"),
            "Criminal Verification": verif("Criminal"), "Criminal Supplier": supplier("Criminal"),
            "Criminal Cost (Incl. VAT)": cost("Criminal"),
            "Qualification Verification": verif("Qualification"), "Qualification Supplier": supplier("Qualification"),
            "Qualification Cost (Incl. VAT)": cost("Qualification"),
            "Matric Verification": verif("Matric"), "Matric Supplier": supplier("Matric"),
            "Matric Cost (Incl. VAT)": cost("Matric"),
            "SAQA (Matric/NSC) Verification": verif("SAQA-Matric"), "SAQA (Matric/NSC) Cost (Incl. VAT)": cost("SAQA-Matric"),
            "SAQA (Tertiary) Verification": verif("SAQA-Tertiary"), "SAQA (Tertiary) Cost (Incl. VAT)": cost("SAQA-Tertiary"),
            "UNISA Verification": verif("UNISA"), "UNISA Cost (Incl. VAT)": cost("UNISA"),
            "MUT Verification": verif("MUT"), "MUT Cost (Incl. VAT)": cost("MUT"),
            "Number of Duplicates": n_dup, "Duplicate Verification Cost (Incl. VAT)": dup_cost,
            "Total Verification Cost (Incl. VAT)": total, "DEL Charge": del_charge,
            "Over / Under": over_under, "Notes": notes,
            "All 5 Core Verifications Complete?": "Yes" if all_core else "No",
        })
    return summary_rows


def build_candidate_calculations(all_rows):
    grouped = group_by_candidate(all_rows)
    calc_rows = []
    for sa_id, info in grouped.items():
        rows = info["rows"]
        block_total = round(sum(float(r.get("Cost (Incl. VAT)") or 0) for r in rows), 2)
        for r in rows:
            excl = float(r.get("Cost (Excl. VAT - as sourced)") or 0)
            incl = float(r.get("Cost (Incl. VAT)") or 0)
            calc_rows.append({
                "Candidate Name": r.get("Candidate Name"), "South African ID Number": sa_id,
                "Verification Type": r.get("Verification Type"), "Supplier": r.get("Verification Supplier"),
                "Verification Result": r.get("Verification Result"),
                "Cost (Excl. VAT)": excl, "VAT (15%)": round(incl - excl, 2), "Cost (Incl. VAT)": incl,
                "Duplicate?": r.get("Duplicate?"), "Calculation / Reason": r.get("Notes"),
                "Candidate Total (Incl. VAT)": block_total,
                "_block_size": len(rows),
            })
    return calc_rows


def build_afika_view(all_rows):
    grouped = group_by_candidate(all_rows)
    out = []
    for sa_id, info in grouped.items():
        rows = info["rows"]
        by_cat = {cat: [r for r in rows if r.get("Category") == cat] for cat in ALL_COST_CATEGORIES}

        def cost_by_supplier(cat, supplier_substr):
            return round(sum(float(r.get("Cost (Incl. VAT)") or 0) for r in by_cat[cat]
                              if supplier_substr.lower() in str(r.get("Verification Supplier") or "").lower()), 2)

        def cat_cost(cat):
            return round(sum(float(r.get("Cost (Incl. VAT)") or 0) for r in by_cat[cat]), 2)

        cit_mie = cost_by_supplier("Citizenship", "MIE")
        cit_dat = cat_cost("Citizenship") - cit_mie
        per_mie = cost_by_supplier("PERSAL", "MIE")
        per_api = cat_cost("PERSAL") - per_mie
        mat_mie = cat_cost("Matric")
        mat_saqa = cat_cost("SAQA-Matric")
        qual_mie = cat_cost("Qualification")
        qual_saqa = cat_cost("SAQA-Tertiary")
        qual_unisa = cat_cost("UNISA")
        qual_mut = cat_cost("MUT")
        crim_mie = cat_cost("Criminal")

        actual = round(cit_mie + cit_dat + per_mie + per_api + crim_mie
                        + qual_mie + qual_saqa + qual_unisa + qual_mut
                        + mat_mie + mat_saqa, 2)
        del_cost = pr.DEL_CHARGE_PER_CANDIDATE

        out.append({
            "Name": info["name"], "ID": sa_id,
            "Citizenship MIE": cit_mie, "Citizenship Datanamix": cit_dat,
            "Citizenship Total": round(cit_mie + cit_dat, 2),
            "Citizenship Yes(#)": 1 if by_cat["Citizenship"] else 0, "Citizenship No": None,
            "Persal MIE": per_mie, "Persal DPSA": per_api, "Persal Total": round(per_mie + per_api, 2),
            "Persal Yes(#)": 1 if by_cat["PERSAL"] else 0, "Persal No": None,
            "Matric MIE": mat_mie, "Matric SAQA": mat_saqa, "Matric Total": round(mat_mie + mat_saqa, 2),
            "Matric Yes(#)": 1 if (by_cat["Matric"] or by_cat["SAQA-Matric"]) else 0, "Matric No": None,
            "Qualification MIE": qual_mie, "Qualification SAQA": qual_saqa,
            "Qualification UNISA": qual_unisa, "Qualification MUT": qual_mut,
            "Qualification Total": round(qual_mie + qual_saqa + qual_unisa + qual_mut, 2),
            "Qualification Yes(#)": 1 if (by_cat["Qualification"] or by_cat["SAQA-Tertiary"] or by_cat["UNISA"] or by_cat["MUT"]) else 0,
            "Qualification No": None,
            "Criminal MIE": crim_mie, "Criminal Total": crim_mie,
            "Criminal Yes(#)": 1 if by_cat["Criminal"] else 0, "Criminal No": None,
            "DEL Cost": del_cost, "Actual Cost": actual, "Difference": round(del_cost - actual, 2),
        })
    return out


def build_total_spend_details(all_rows, candidate_summary_rows):
    def sum_cost(pred):
        return round(sum(float(r.get("Cost (Incl. VAT)") or 0) for r in all_rows if pred(r)), 2)

    n_candidates = len(candidate_summary_rows)
    n_verifications = len(all_rows)
    n_dup = sum(1 for r in all_rows if str(r.get("Duplicate?")) == "Yes")
    dup_cost = sum_cost(lambda r: str(r.get("Duplicate?")) == "Yes")

    total_spend = round(sum(float(c.get("Total Verification Cost (Incl. VAT)") or 0) for c in candidate_summary_rows), 2)
    total_del = round(sum(float(c.get("DEL Charge") or 0) for c in candidate_summary_rows), 2)
    profit_loss = round(total_del - total_spend, 2)
    n_complete = sum(1 for c in candidate_summary_rows if c.get("All 5 Core Verifications Complete?") == "Yes")

    return {
        "Total Number of Candidates": n_candidates,
        "Total Number of Verifications": n_verifications,
        "Total Duplicate Verifications": n_dup,
        "Total Cost Attributable to Duplicates (Incl. VAT)": dup_cost,
        "Total MIE Spend": sum_cost(lambda r: r.get("Verification Supplier") == "MIE"),
        "Total Datanamics Spend": sum_cost(lambda r: r.get("Verification Supplier") == "Datanamics API"),
        "Total API Spend (PERSAL assumed)": sum_cost(lambda r: r.get("Verification Supplier") == "API"),
        "Total SAQA Spend - Matric/NSC": sum_cost(lambda r: r.get("Category") == "SAQA-Matric"),
        "Total SAQA Spend - Tertiary": sum_cost(lambda r: r.get("Category") == "SAQA-Tertiary"),
        "Total UNISA Spend": sum_cost(lambda r: r.get("Category") == "UNISA"),
        "Total MUT Spend": sum_cost(lambda r: r.get("Category") == "MUT"),
        "Total Spend Across All Suppliers": total_spend,
        "Total DEL Revenue / Charges": total_del,
        "Total Profit / Loss": profit_loss,
        "Number of Candidates With All 5 Core Verifications Complete": n_complete,
    }
