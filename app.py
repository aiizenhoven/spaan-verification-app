"""
Spaan Verification Cost & Reconciliation — local app.

Upload result report documents (MIE Captured Credentials, SAQA/NLRD PDFs, UNISA R4C
uploads, MUT candidate files); the app prices each record per the established business
rules, dedupes against the master workbook, and lets you apply the update and browse
every sheet.

Run with:  streamlit run app.py
"""
import os
import tempfile
import datetime

import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth

import parsers
import engine
import aggregate
import workbook_writer

st.set_page_config(page_title="Spaan Verification Cost & Reconciliation", layout="wide")

# ---------------------------------------------------------------------------
# Login gate — credentials come from Streamlit secrets, never from the repo.
# Locally: .streamlit/secrets.toml (gitignored). On Streamlit Community Cloud:
# App settings -> Secrets.
# ---------------------------------------------------------------------------
def _to_plain_dict(value):
    """Recursively convert Streamlit's read-only Secrets mapping to plain dicts,
    since streamlit-authenticator mutates the credentials structure at runtime."""
    if hasattr(value, "items"):
        return {k: _to_plain_dict(v) for k, v in value.items()}
    return value


authenticator = stauth.Authenticate(
    _to_plain_dict(st.secrets["credentials"]),
    st.secrets["cookie"]["name"],
    st.secrets["cookie"]["key"],
    st.secrets["cookie"]["expiry_days"],
)
authenticator.login()

if st.session_state.get("authentication_status") is False:
    st.error("Username or password is incorrect.")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("Please log in to continue.")
    st.stop()

with st.sidebar:
    st.write(f"Signed in as **{st.session_state['name']}**")
    authenticator.logout()

SPAAN_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700&display=swap');

html, body, [class*="css"]  { font-family: 'Montserrat', sans-serif; }
:root {
    --navy: #26384A; --ocean: #3E6A90; --denim: #7CA0C2;
    --shelduck: #E67E30; --cinnamon: #A8461F; --tobacco: #5C3220; --cement: #D9D5CF;
}
h1, h2, h3 { color: var(--navy) !important; font-weight: 700 !important; }
[data-testid="stMetricValue"] { color: var(--navy); font-weight: 700; }
.stButton>button {
    background-color: var(--shelduck); color: white !important; font-weight: 700;
    border-radius: 8px; border: none;
}
.stButton>button:hover { background-color: var(--cinnamon); color: white !important; }
[data-testid="stSidebar"] { background-color: var(--navy); }
[data-testid="stSidebar"] * { color: white !important; }
.spaan-banner {
    background: var(--navy); color: white; padding: 20px 28px; border-radius: 8px;
    margin-bottom: 20px; font-weight: 700; font-size: 26px;
}
.spaan-banner span { color: var(--shelduck); }
</style>
"""
st.markdown(SPAAN_CSS, unsafe_allow_html=True)
st.markdown('<div class="spaan-banner">spa<span>a</span>n // Verification Cost &amp; Reconciliation</div>',
            unsafe_allow_html=True)

if "master_path" not in st.session_state:
    st.session_state.master_path = None
if "output_path" not in st.session_state:
    st.session_state.output_path = None
if "pending_add" not in st.session_state:
    st.session_state.pending_add = None

WORKDIR = tempfile.gettempdir()


def save_upload(uploaded_file):
    path = os.path.join(WORKDIR, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


# ---------------------------------------------------------------------------
# Sidebar: master workbook
# ---------------------------------------------------------------------------
st.sidebar.header("Master workbook")
master_upload = st.sidebar.file_uploader(
    "Verification Cost Analysis & Candidate Reconciliation.xlsx", type=["xlsx"], key="master")
if master_upload is not None:
    st.session_state.master_path = save_upload(master_upload)
    st.session_state.output_path = None

if st.session_state.master_path:
    st.sidebar.success(os.path.basename(st.session_state.master_path))
else:
    st.sidebar.info("Upload the master workbook to begin.")

active_path = st.session_state.output_path or st.session_state.master_path

if not active_path:
    st.stop()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_ingest, tab_browse = st.tabs(["Ingest new reports", "Browse workbook"])

# ---------------------------------------------------------------------------
# Ingest tab
# ---------------------------------------------------------------------------
with tab_ingest:
    st.subheader("Upload result report documents")
    st.caption("Recognized formats: MIE Captured Credentials (.xlsx), SAQA/NLRD report (.pdf), "
               "UNISA R4C upload (.xlsx), MUT Candidates (.xlsx).")

    uploads = st.file_uploader("Drop one or more report files", accept_multiple_files=True,
                                type=["xlsx", "xls", "pdf"])

    saqa_modes = {}
    if uploads:
        pdfs = [u for u in uploads if u.name.lower().endswith(".pdf")]
        for u in pdfs:
            saqa_modes[u.name] = st.radio(
                f"'{u.name}' — is this a bulk or individual SAQA submission?",
                ["bulk", "individual"], horizontal=True, key=f"mode_{u.name}")

    if uploads and st.button("Parse & preview"):
        ledger = engine.load_ledger(active_path)
        all_new, all_skipped, batch_names = [], [], []
        for u in uploads:
            path = save_upload(u)
            batch_names.append(u.name)
            try:
                if u.name.lower().endswith(".pdf"):
                    records, skipped = parsers.parse_saqa_pdf(path, mode=saqa_modes.get(u.name, "bulk"))
                    fmt = "SAQA/NLRD PDF"
                else:
                    (records, skipped), fmt = parsers.detect_and_parse(path)
                st.write(f"**{u.name}** — detected as *{fmt}*, {len(records)} record(s) parsed"
                         + (f", {len(skipped)} skipped" if skipped else ""))
                all_new.extend(records)
                all_skipped.extend(skipped)
            except Exception as e:
                st.error(f"Could not parse '{u.name}': {e}")

        added, already_ingested = engine.merge_batch(ledger, all_new)
        fallback_rows = engine.apply_assumed_fallbacks(ledger + added)

        st.session_state.pending_add = {
            "added": added, "already_ingested": already_ingested,
            "fallback_rows": fallback_rows, "skipped": all_skipped, "batch_names": batch_names,
        }

    pending = st.session_state.pending_add
    if pending:
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("New verification rows", len(pending["added"]))
        c2.metric("Already in workbook (skipped)", len(pending["already_ingested"]))
        c3.metric("Assumed fallback rows", len(pending["fallback_rows"]))
        c4.metric("Unrecognized (not priced)", len(pending["skipped"]))

        new_cost = sum(r["Cost (Incl. VAT)"] for r in pending["added"] + pending["fallback_rows"])
        st.metric("Total new cost added (Incl. VAT)", f"R {new_cost:,.2f}")

        if pending["added"]:
            st.write("**New rows to add:**")
            st.dataframe(pd.DataFrame(pending["added"]), use_container_width=True, height=280)
        if pending["skipped"]:
            st.write("**Unrecognized rows (no pricing rule — not added):**")
            st.dataframe(pd.DataFrame([{"reason": s["reason"]} for s in pending["skipped"]]),
                         use_container_width=True)

        if st.button("Apply to workbook", type="primary"):
            out_path = os.path.join(
                WORKDIR, f"Verification Cost Analysis & Candidate Reconciliation Updated "
                         f"{datetime.date.today().isoformat()}.xlsx")
            result = workbook_writer.write_updated_workbook(
                st.session_state.master_path, out_path,
                pending["added"] + pending["fallback_rows"],
                batch_source_names=pending["batch_names"],
            )
            st.session_state.output_path = out_path
            st.session_state.pending_add = None
            st.success(f"Workbook updated — {result['new_rows_added']} rows added, "
                       f"{result['total_candidates']} candidates, "
                       f"{result['total_verification_rows']} total verification rows.")
            with open(out_path, "rb") as f:
                st.download_button("Download updated workbook", f,
                                    file_name=os.path.basename(out_path))

# ---------------------------------------------------------------------------
# Browse tab
# ---------------------------------------------------------------------------
with tab_browse:
    ledger = engine.load_ledger(active_path)
    candidate_summary_rows = aggregate.build_candidate_summary(ledger)
    total_spend = aggregate.build_total_spend_details(ledger, candidate_summary_rows)

    st.subheader("Total Spend Details")
    cols = st.columns(4)
    cols[0].metric("Candidates", total_spend["Total Number of Candidates"])
    cols[1].metric("Verification rows", total_spend["Total Number of Verifications"])
    cols[2].metric("Total spend (Incl. VAT)", f"R {total_spend['Total Spend Across All Suppliers']:,.2f}")
    cols[3].metric("Profit / Loss vs DEL", f"R {total_spend['Total Profit / Loss']:,.2f}")

    with st.expander("Full spend breakdown"):
        st.table(pd.DataFrame(total_spend.items(), columns=["Metric", "Value"]))

    st.divider()
    st.subheader("Candidate Summary")
    search = st.text_input("Search candidate name or ID")
    df_summary = pd.DataFrame(candidate_summary_rows)
    if search:
        mask = (df_summary["Candidate Name"].astype(str).str.contains(search, case=False, na=False)
                | df_summary["South African ID Number"].astype(str).str.contains(search, na=False))
        df_summary = df_summary[mask]
    st.dataframe(df_summary, use_container_width=True, height=400)

    st.divider()
    st.subheader("Consolidated Results (raw ledger)")
    df_ledger = pd.DataFrame(ledger)
    id_filter = st.text_input("Filter by South African ID Number", key="ledger_filter")
    if id_filter:
        df_ledger = df_ledger[df_ledger["South African ID Number"].astype(str).str.contains(id_filter, na=False)]
    st.dataframe(df_ledger, use_container_width=True, height=400)
