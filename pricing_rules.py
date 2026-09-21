"""
Pricing / business-rule engine for the Verification Cost & Reconciliation workbook.

Every rate and rule here mirrors the "Pricing" sheet and the business rules documented
in its Notes column, as established in the original reconciliation workbook.
"""

VAT_RATE = 0.15

# Institution surcharge keywords, checked in order (first match wins), against the
# free-text MIE "Supplier" field for Qualification-category rows.
# (keyword, breakdown_code, breakdown_label, surcharge_excl_vat)
INSTITUTION_SURCHARGES = [
    ("department of education - colleges", "DOE - Colleges", "Department of Education - Colleges", 82.00),
    ("quality council for trades", "QCTO", "Quality Council for Trades & Occupations", 110.00),
    ("examinations council of lesotho", "Exam Council Lesotho", "Examinations Council of Lesotho", 575.00),
    ("southern african emergency services institute", "SAESI", "Southern African Emergency Services Institute", 150.00),
    ("university of johannesburg", "UJ", "University of Johannesburg", 168.00),
    ("university of pretoria", "UP", "University of Pretoria", 170.00),
    ("witwatersrand", "Wits", "University of Witwatersrand", 100.00),
    ("nelson mandela", "NMMU", "Nelson Mandela Metropolitan University", 120.00),
]

QUALIFICATION_BASE = 46.86
TERTIARY_COURSE_ADD = 7.00

MATRIC_NATIONAL_SECONDARY_FEE = 65.00
MATRIC_SYMBOL_MATCH_FEE = 130.00
MATRIC_BASE = 46.86

CITIZENSHIP_MIE = 27.50
CITIZENSHIP_ASSUMED = 15.90  # Datanamics API, VAT IS applied to this one

PERSAL_MIE = 87.86
PERSAL_ASSUMED = 0.0  # API, free

CRIMINAL_STANDARD = 219.53
CRIMINAL_FINGERPRINT_ZONE = 75.00

SAQA_BULK_FLAT = 58.00       # no VAT
SAQA_INDIVIDUAL_FLAT = 82.00  # no VAT

DEL_CHARGE_PER_CANDIDATE = 593.71  # flat, no VAT


def incl_vat(excl):
    return round(excl * (1 + VAT_RATE), 2)


def match_institution_surcharge(supplier_text):
    """Return (code, label, surcharge_excl_vat) or (None, None, 0.0)."""
    text = (supplier_text or "").lower()
    for keyword, code, label, surcharge in INSTITUTION_SURCHARGES:
        if keyword in text:
            return code, label, surcharge
    return None, None, 0.0


def price_citizenship_mie():
    excl = CITIZENSHIP_MIE
    return excl, incl_vat(excl), "MIE", f"MIE Citizenship flat fee R{excl:.2f} | Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl_vat(excl):.2f}"


def price_citizenship_assumed():
    excl = CITIZENSHIP_ASSUMED
    return excl, incl_vat(excl), "Datanamics API", (
        "No MIE Citizenship record found for this candidate; assumed completed via Datanamics API "
        f"per pricing rules (R{excl:.2f} excl VAT). | Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl_vat(excl):.2f}"
    )


def price_persal_mie():
    excl = PERSAL_MIE
    return excl, incl_vat(excl), "MIE", f"MIE PERSAL flat fee R{excl:.2f} | Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl_vat(excl):.2f}"


def price_persal_assumed():
    excl = PERSAL_ASSUMED
    return excl, incl_vat(excl), "API", (
        "No MIE PERSAL record found for this candidate; assumed completed via API per pricing rules "
        f"(Free, R{excl:.2f})."
    )


def price_criminal():
    excl = round(CRIMINAL_STANDARD + CRIMINAL_FINGERPRINT_ZONE, 2)
    return excl, incl_vat(excl), "MIE (FingerprintZone)", (
        f"MIE Criminal Standard R{CRIMINAL_STANDARD:.2f} + Fingerprint Zone R{CRIMINAL_FINGERPRINT_ZONE:.2f} = "
        f"R{excl:.2f} | Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl_vat(excl):.2f}"
    )


def price_qualification(qualification_kind, supplier_text):
    """
    qualification_kind: 'National Qualifications Register' | 'National Tertiary' | 'Tertiary Course'
    Returns (excl, incl, breakdown_label, notes)
    """
    code, label, surcharge = match_institution_surcharge(supplier_text)
    breakdown = label if label else "Base (no institution surcharge)"

    parts_excl = [QUALIFICATION_BASE]
    notes_parts = [f"Qualification base R{QUALIFICATION_BASE:.2f}"]

    if qualification_kind == "Tertiary Course":
        parts_excl.append(TERTIARY_COURSE_ADD)
        notes_parts.append(f"Tertiary Course Add R{TERTIARY_COURSE_ADD:.2f}")

    if surcharge:
        parts_excl.append(surcharge)
        notes_parts.append(f"{label} institution surcharge R{surcharge:.2f}")

    excl = round(sum(parts_excl), 2)
    incl = incl_vat(excl)
    notes = " + ".join(notes_parts) + f" | Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl:.2f}"
    return excl, incl, breakdown, notes


def price_matric(kind):
    """kind: 'National Secondary' | 'Symbol Match'"""
    if kind == "National Secondary":
        fee = MATRIC_NATIONAL_SECONDARY_FEE
        excl = round(MATRIC_BASE + fee, 2)
        notes = (f"UMALUSI base R{MATRIC_BASE:.2f} + National Secondary Fee R{fee:.2f} = R{excl:.2f} | "
                  f"Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl_vat(excl):.2f}")
        return excl, incl_vat(excl), "National Secondary", notes
    else:
        fee = MATRIC_SYMBOL_MATCH_FEE
        excl = round(MATRIC_BASE + fee, 2)
        notes = (f"UMALUSI base R{MATRIC_BASE:.2f} + Verification/Symbol Match Fee R{fee:.2f} = R{excl:.2f} | "
                  f"Incl. 15% VAT: R{excl:.2f} x 1.15 = R{incl_vat(excl):.2f}")
        return excl, incl_vat(excl), "Symbol Match", notes


def price_saqa(mode):
    """mode: 'bulk' | 'individual'. Flat fee, NO VAT."""
    flat = SAQA_BULK_FLAT if mode == "bulk" else SAQA_INDIVIDUAL_FLAT
    label = "" if mode == "bulk" else " (Individual)"
    notes = (f"SAQA Verification{label} flat fee R{flat:.2f} per verification (per pricing rules). "
              "No VAT grossed up on SAQA fee (matches Afika template).")
    return flat, flat, "SAQA", notes


def price_direct_free(kind, detail_note):
    """kind: 'UNISA' | 'MUT'. Always free."""
    return 0.0, 0.0, detail_note
