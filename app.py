"""
app.py  —  Forus Toolkit Manager
Streamlit web interface for the Forus Resilience & Support Toolkit.

Run locally:   streamlit run app.py
Deploy to:     Streamlit Community Cloud (connect GitHub repo, set ANTHROPIC_API_KEY secret)
"""

import io
import os
import shutil
import sys
import tempfile
import datetime
from pathlib import Path

import json

import openpyxl
import pandas as pd
import requests
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Forus Toolkit Manager",
    page_icon="🛡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Password gate ──────────────────────────────────────────────────────────────
def _check_password():
    """Returns True if the user has entered the correct password."""
    if st.session_state.get("authenticated"):
        return True
    st.markdown(
        """
        <div style='max-width:380px;margin:80px auto 0;text-align:center'>
            <h2 style='color:#00424D'>🛡 Forus Toolkit Manager</h2>
            <p style='color:#555;margin-bottom:1.5rem'>Enter the access password to continue.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("Password", type="password", label_visibility="collapsed",
                            placeholder="Password")
        if st.button("Login", use_container_width=True, type="primary"):
            correct = st.secrets.get("APP_PASSWORD", "")
            if pwd == correct and correct != "":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password — please try again.")
    st.stop()

_check_password()

# ── Brand colours ─────────────────────────────────────────────────────────────
TEAL  = "#58C5C7"
DARK  = "#00424D"
PINK  = "#ED1651"
MINT  = "#5C9C8E"
LIME  = "#B2C100"
LGREY = "#F5F5F5"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  /* Sidebar brand strip */
  [data-testid="stSidebar"] {{
    background-color: {DARK};
  }}
  [data-testid="stSidebar"] * {{
    color: white !important;
  }}
  [data-testid="stSidebar"] .stRadio label {{
    font-size: 15px;
    padding: 4px 0;
  }}
  /* Page headings */
  h1 {{ color: {DARK}; border-bottom: 3px solid {TEAL}; padding-bottom: 6px; }}
  h2 {{ color: {DARK}; }}
  h3 {{ color: {MINT}; }}
  /* Cards */
  .forus-card {{
    background: {LGREY};
    border-left: 4px solid {TEAL};
    border-radius: 4px;
    padding: 14px 16px;
    margin-bottom: 12px;
  }}
  .forus-card.change   {{ border-color: {PINK}; }}
  .forus-card.approved {{ border-color: {MINT}; background: #E8F5F2; }}
  .forus-card.rejected {{ border-color: #888; background: #F9F9F9; opacity: 0.7; }}
  /* Confidence badges */
  .badge {{
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: bold;
    color: white;
  }}
  .badge.HIGH   {{ background: {MINT}; }}
  .badge.MEDIUM {{ background: {LIME}; color: {DARK}; }}
  .badge.LOW    {{ background: {PINK}; }}
  /* Stat boxes */
  .stat-box {{
    background: white;
    border: 1px solid #DDD;
    border-top: 4px solid {TEAL};
    border-radius: 4px;
    padding: 16px;
    text-align: center;
  }}
  .stat-num {{ font-size: 32px; font-weight: bold; color: {DARK}; }}
  .stat-lbl {{ font-size: 13px; color: #666; margin-top: 4px; }}
</style>
""", unsafe_allow_html=True)


# ── Session state ─────────────────────────────────────────────────────────────
def _init():
    if "sp_path" not in st.session_state:
        local = Path("Forus_Toolkit_Content_DB.xlsx")
        if local.exists():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            shutil.copy(str(local), tmp.name)
            tmp.close()
            st.session_state["sp_path"] = tmp.name
            st.session_state["sp_name"] = local.name
        else:
            st.session_state["sp_path"] = None
            st.session_state["sp_name"] = None
    if "action_log" not in st.session_state:
        st.session_state["action_log"] = []

_init()


# ── Helpers ───────────────────────────────────────────────────────────────────
def sp():
    """Return current spreadsheet path (or None)."""
    return st.session_state.get("sp_path")


def _load_wb(data_only=True):
    return openpyxl.load_workbook(sp(), data_only=data_only)


def _save_wb(wb):
    wb.save(sp())


def _sheet_to_df(wb, sheet_name, header_row=2):
    if sheet_name not in wb.sheetnames:
        return pd.DataFrame()
    ws   = wb[sheet_name]
    hdrs = [c.value for c in ws[header_row]]
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if any(v is not None for v in row):
            rows.append({hdrs[i]: row[i] for i in range(len(hdrs)) if i < len(row)})
    return pd.DataFrame(rows)


def _badge(conf):
    conf = str(conf or "").upper()
    return f'<span class="badge {conf}">{conf}</span>'


# ── Google Drive helpers ──────────────────────────────────────────────────────
def _gdrive_service():
    """Build an authenticated Google Drive service using service-account credentials."""
    creds_json = st.secrets.get("GDRIVE_CREDENTIALS", "")
    if not creds_json:
        return None
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds_dict = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive.file"],
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        st.warning(f"Could not build Drive service: {e}")
        return None


def _load_from_gdrive():
    """Download spreadsheet from Google Drive on first load."""
    if st.session_state.get("sp_path"):
        return
    file_id = st.secrets.get("GDRIVE_FILE_ID", "")
    if not file_id:
        return

    svc = _gdrive_service()
    try:
        if svc:
            # Authenticated download via service account
            from googleapiclient.http import MediaIoBaseDownload
            request = svc.files().get_media(fileId=file_id)
            buf = io.BytesIO()
            dl = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = dl.next_chunk()
            data = buf.getvalue()
        else:
            # Fallback: public download link
            url = f"https://drive.google.com/uc?export=download&id={file_id}"
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            data = r.content

        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.write(data)
        tmp.close()
        st.session_state["sp_path"] = tmp.name
        st.session_state["sp_name"] = "Forus_Toolkit_Content_DB.xlsx (Google Drive)"
    except Exception as e:
        st.warning(f"Could not load spreadsheet from Google Drive: {e}")


def save_to_gdrive():
    """Upload the current spreadsheet back to Google Drive, replacing the original."""
    file_id = st.secrets.get("GDRIVE_FILE_ID", "")
    if not file_id or not sp():
        return False, "No file ID or spreadsheet path configured."
    svc = _gdrive_service()
    if not svc:
        return False, "No service-account credentials found (GDRIVE_CREDENTIALS secret not set)."
    try:
        from googleapiclient.http import MediaFileUpload
        media = MediaFileUpload(
            sp(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False,
        )
        svc.files().update(fileId=file_id, media_body=media).execute()
        return True, "Saved to Google Drive."
    except Exception as e:
        return False, str(e)


_load_from_gdrive()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"## 🛡 Forus Toolkit")
    st.markdown("---")

    # Spreadsheet status
    if sp():
        st.success(f"📊 {st.session_state['sp_name']}")
        with open(sp(), "rb") as f:
            st.download_button(
                "⬇ Download updated spreadsheet",
                f.read(),
                file_name="Forus_Toolkit_Content_DB.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        if st.secrets.get("GDRIVE_CREDENTIALS"):
            if st.button("☁ Save to Google Drive", use_container_width=True):
                ok, msg = save_to_gdrive()
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
    else:
        st.warning("No spreadsheet loaded")
        uploaded = st.file_uploader("Upload spreadsheet (.xlsx)", type=["xlsx"])
        if uploaded:
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            tmp.write(uploaded.read())
            tmp.close()
            st.session_state["sp_path"] = tmp.name
            st.session_state["sp_name"] = uploaded.name
            st.rerun()

    st.markdown("---")
    page = st.radio(
        "Navigation",
        ["📊 Dashboard", "🔍 Check Mechanisms", "📋 Review Queue",
         "✅ Apply Approved", "📄 Generate PDF"],
        label_visibility="collapsed",
    )

    # Recent action log
    if st.session_state["action_log"]:
        st.markdown("---")
        st.markdown("**Recent actions**")
        for entry in st.session_state["action_log"][-5:]:
            st.markdown(f"<small>{entry}</small>", unsafe_allow_html=True)


# ── Guard: require spreadsheet ────────────────────────────────────────────────
def require_spreadsheet():
    if not sp():
        st.warning("Please upload the spreadsheet using the sidebar to get started.")
        st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Dashboard
# ═══════════════════════════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.title("Forus Toolkit — Dashboard")
    require_spreadsheet()

    wb       = _load_wb()
    df_cont  = _sheet_to_df(wb, "CONTENT")
    df_mech  = _sheet_to_df(wb, "MECHANISMS")
    df_rq    = _sheet_to_df(wb, "REVIEW_QUEUE")

    today = datetime.date.today()

    # ── Key stats ─────────────────────────────────────────────────────────────
    total_blocks  = len(df_cont)
    over_limit    = int(df_cont["within_limit"].apply(
        lambda x: str(x or "").startswith("OVER")).sum()) if "within_limit" in df_cont else 0

    pending_rq = 0
    if not df_rq.empty and "status" in df_rq:
        pending_rq = int((df_rq["status"].str.upper() == "PENDING").sum())

    mechs_due = 0
    if not df_mech.empty and "next_verify_due" in df_mech:
        window = today + datetime.timedelta(days=14)
        for _, row in df_mech.iterrows():
            nvd = row.get("next_verify_due")
            if nvd:
                try:
                    if isinstance(nvd, datetime.datetime): nvd = nvd.date()
                    elif not isinstance(nvd, datetime.date): nvd = datetime.date.fromisoformat(str(nvd))
                    if nvd <= window: mechs_due += 1
                except (ValueError, TypeError):
                    pass
            if str(row.get("status","")).upper() == "VERIFY": mechs_due += 1

    c1, c2, c3, c4 = st.columns(4)
    for col, num, label, colour in [
        (c1, total_blocks,  "Content blocks",     TEAL),
        (c2, over_limit,    "Over word limit",     PINK if over_limit else MINT),
        (c3, pending_rq,    "Awaiting review",     PINK if pending_rq else MINT),
        (c4, mechs_due,     "Mechanisms due",      LIME if mechs_due else MINT),
    ]:
        col.markdown(
            f'<div class="stat-box" style="border-top-color:{colour}">'
            f'<div class="stat-num">{num}</div>'
            f'<div class="stat-lbl">{label}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("")

    # ── Mechanisms table ──────────────────────────────────────────────────────
    st.subheader("Mechanisms")
    if not df_mech.empty:
        show_cols = ["mech_id", "mechanism_name", "category", "status",
                     "platform_eligible", "last_verified", "next_verify_due"]
        show_cols = [c for c in show_cols if c in df_mech.columns]
        st.dataframe(df_mech[show_cols], use_container_width=True, hide_index=True)
    else:
        st.info("No MECHANISMS sheet found.")

    # ── Review queue summary ──────────────────────────────────────────────────
    if not df_rq.empty and "status" in df_rq:
        st.subheader("Review Queue")
        status_counts = df_rq["status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        st.dataframe(status_counts, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Check Mechanisms
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🔍 Check Mechanisms":
    st.title("Check Mechanisms")
    st.markdown(
        "Run AI-powered verification against each mechanism entry whose "
        f"next check date is within the next 14 days, or is flagged **VERIFY**. "
        "Results are written to the **REVIEW_QUEUE** tab of the spreadsheet."
    )
    require_spreadsheet()

    wb      = _load_wb()
    df_mech = _sheet_to_df(wb, "MECHANISMS")

    today  = datetime.date.today()
    window = today + datetime.timedelta(days=14)

    # Identify due rows
    due = []
    if not df_mech.empty:
        for _, row in df_mech.iterrows():
            nvd = row.get("next_verify_due")
            is_due = False
            if nvd:
                try:
                    if isinstance(nvd, datetime.datetime): nvd = nvd.date()
                    elif not isinstance(nvd, datetime.date): nvd = datetime.date.fromisoformat(str(nvd))
                    if nvd <= window: is_due = True
                except (ValueError, TypeError):
                    pass
            if str(row.get("status","")).upper() == "VERIFY": is_due = True
            if is_due:
                due.append(row)

    if not due:
        st.success("✓ No mechanisms are currently due for verification.")
    else:
        st.info(f"**{len(due)} mechanism(s)** due for verification:")
        due_df = pd.DataFrame(due)[["mech_id","mechanism_name","category","next_verify_due","status"]]
        st.dataframe(due_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("Run verification")

        api_key = st.text_input(
            "Anthropic API key",
            type="password",
            value=os.environ.get("ANTHROPIC_API_KEY", ""),
            help="Required to call Claude with web search. Set ANTHROPIC_API_KEY env var to pre-fill.",
        )

        if st.button("🔍 Run checks now", type="primary", disabled=not api_key):
            # Import the agent function from generate_toolkit
            sys.path.insert(0, str(Path(__file__).parent))
            try:
                import generate_toolkit as gt
            except ImportError as e:
                st.error(f"Could not import generate_toolkit: {e}")
                st.stop()

            # Patch SPREADSHEET global to point at working copy
            gt.SPREADSHEET = sp()

            progress = st.progress(0, text="Starting…")
            log      = st.empty()
            results  = []

            for i, (_, mrow) in enumerate(pd.DataFrame(due).iterrows()):
                mech_dict = mrow.to_dict()
                mname     = mech_dict.get("mechanism_name", "")
                mid       = mech_dict.get("mech_id", "")
                progress.progress((i) / len(due), text=f"Checking {mid} — {mname}…")

                result = gt.call_ai_agent(mech_dict, api_key)
                results.append((mech_dict, result))

            progress.progress(1.0, text="Done!")

            # Write results to REVIEW_QUEUE
            wb2   = openpyxl.load_workbook(sp())
            ws_m  = wb2["MECHANISMS"]
            ws_rq = wb2["REVIEW_QUEUE"]
            mhdrs = [c.value for c in ws_m[2]]
            mcm   = {h: i for i, h in enumerate(mhdrs) if h}

            max_n = 0
            for row in ws_rq.iter_rows(min_row=3, values_only=True):
                rid = str(row[1] or "")
                if rid.startswith("RQ-"):
                    try: max_n = max(max_n, int(rid[3:]))
                    except ValueError: pass
            next_id = [max_n]

            # Find row index for each mech_id in MECHANISMS
            mech_row_map = {}
            for r_idx, row in enumerate(ws_m.iter_rows(min_row=3, values_only=False), start=3):
                mid = row[0].value
                if mid: mech_row_map[str(mid).strip()] = r_idx

            summary = {"NO_CHANGE": 0, "CHANGE_DETECTED": 0, "UNABLE_TO_VERIFY": 0}
            for mech_dict, result in results:
                mid    = mech_dict.get("mech_id", "")
                name   = mech_dict.get("mechanism_name", "")
                cat    = mech_dict.get("category", "")
                status = result.get("status", "UNABLE_TO_VERIFY")
                conf   = result.get("confidence", "LOW")
                summary[status] = summary.get(status, 0) + 1

                if status == "NO_CHANGE":
                    row_idx = mech_row_map.get(str(mid).strip())
                    if row_idx:
                        gt._update_mech_verified(ws_m, row_idx, mcm, today, cat)

                elif status == "CHANGE_DETECTED":
                    for ch in result.get("changes", []):
                        gt._rq_append(ws_rq, {
                            "date_flagged": str(today), "mech_id": mid,
                            "mechanism_name": name, "category": cat,
                            "change_type": "UPDATED_INFO",
                            "field": ch.get("field",""),
                            "current_value": ch.get("current_value",""),
                            "proposed_value": ch.get("proposed_value",""),
                            "reason": ch.get("reason",""),
                            "source_url": ch.get("source_url",""),
                            "confidence": conf,
                        }, next_id)
                    for nm in result.get("new_mechanisms_found", []):
                        gt._rq_append(ws_rq, {
                            "date_flagged": str(today), "mech_id": mid,
                            "mechanism_name": nm.get("name",""), "category": cat,
                            "change_type": "NEW_ENTRY", "field": "new_mechanism",
                            "current_value": "",
                            "proposed_value": (f"{nm.get('name','')}  |  "
                                               f"{nm.get('organisation','')}  |  "
                                               f"{nm.get('url','')}"),
                            "reason": nm.get("reason",""),
                            "source_url": nm.get("url",""),
                            "confidence": conf,
                        }, next_id)

                else:  # UNABLE_TO_VERIFY
                    gt._rq_append(ws_rq, {
                        "date_flagged": str(today), "mech_id": mid,
                        "mechanism_name": name, "category": cat,
                        "change_type": "UNABLE_TO_VERIFY", "field": "all",
                        "current_value": "", "proposed_value": "",
                        "reason": result.get("notes",""),
                        "source_url": "", "confidence": conf,
                    }, next_id)

            wb2.save(sp())

            # Summary
            st.success(
                f"✓ Done — {summary['NO_CHANGE']} unchanged, "
                f"{summary['CHANGE_DETECTED']} change(s) detected, "
                f"{summary['UNABLE_TO_VERIFY']} unable to verify."
            )
            if summary["CHANGE_DETECTED"] or summary["UNABLE_TO_VERIFY"]:
                st.info("Switch to the **Review Queue** page to review and approve proposed changes.")
            st.session_state["action_log"].append(
                f"{today} — Checked {len(due)} mechanisms"
            )


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Review Queue
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Review Queue":
    st.title("Review Queue")
    require_spreadsheet()

    wb    = _load_wb(data_only=False)
    ws_rq = wb["REVIEW_QUEUE"] if "REVIEW_QUEUE" in wb.sheetnames else None

    if ws_rq is None:
        st.warning("No REVIEW_QUEUE sheet found in spreadsheet.")
        st.stop()

    rq_hdrs = [c.value for c in ws_rq[2]]
    rqcm    = {h: i for i, h in enumerate(rq_hdrs) if h}

    # Filter controls
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        filter_status = st.multiselect(
            "Filter by status",
            ["PENDING", "APPROVED", "REJECTED", "COMPLETED", "COMPLETED_MANUAL"],
            default=["PENDING"],
        )
    with col_f2:
        filter_cat = st.multiselect("Filter by category", ["legal", "emergency-funding", "digital-security"])

    # Load rows
    rows = []
    for r_idx, row in enumerate(ws_rq.iter_rows(min_row=3, values_only=False), start=3):
        rid = row[rqcm.get("review_id", 1)].value
        if not rid:
            continue
        status = str(row[rqcm.get("status", 13)].value or "").strip().upper()
        cat    = str(row[rqcm.get("category", 5)].value or "")
        if filter_status and status not in [s.upper() for s in filter_status]:
            continue
        if filter_cat and cat not in filter_cat:
            continue
        rows.append((r_idx, row, rid, status, cat))

    if not rows:
        st.info("No items match the current filter.")
    else:
        st.markdown(f"**{len(rows)} item(s)** shown")
        st.markdown("")

    action_taken = False

    for r_idx, row, rid, status, cat in rows:
        mech   = str(row[rqcm.get("mechanism_name",  4)].value or "")
        ctyp   = str(row[rqcm.get("change_type",     6)].value or "")
        fld    = str(row[rqcm.get("field",            7)].value or "")
        cur    = str(row[rqcm.get("current_value",    8)].value or "")
        prop   = str(row[rqcm.get("proposed_value",   9)].value or "")
        reason = str(row[rqcm.get("reason",          10)].value or "")
        src    = str(row[rqcm.get("source_url",      11)].value or "")
        conf   = str(row[rqcm.get("confidence",      12)].value or "")
        date_f = str(row[rqcm.get("date_flagged",     2)].value or "")

        card_cls = {
            "PENDING": "forus-card change",
            "APPROVED": "forus-card approved",
            "REJECTED": "forus-card rejected",
        }.get(status, "forus-card")

        st.markdown(
            f'<div class="{card_cls}">'
            f'<strong>{rid}</strong> &nbsp; {_badge(conf)} &nbsp; '
            f'<span style="color:#888;font-size:13px">{ctyp} · {cat} · {date_f}</span><br>'
            f'<strong>{mech}</strong> — <em>{fld}</em>'
            f'</div>',
            unsafe_allow_html=True,
        )

        inner_c1, inner_c2 = st.columns(2)
        with inner_c1:
            st.markdown(f"**Current value**")
            st.markdown(f"<div style='background:#FFF8F8;padding:8px;border-radius:4px;font-size:13px'>{cur or '—'}</div>",
                        unsafe_allow_html=True)
        with inner_c2:
            st.markdown(f"**Proposed value**")
            st.markdown(f"<div style='background:#F8FFF8;padding:8px;border-radius:4px;font-size:13px'>{prop or '—'}</div>",
                        unsafe_allow_html=True)

        if reason:
            st.markdown(f"<small>**Reason:** {reason}</small>", unsafe_allow_html=True)
        if src:
            st.markdown(f"<small>**Source:** <a href='{src}' target='_blank'>{src}</a></small>",
                        unsafe_allow_html=True)

        if status == "PENDING":
            btn_c1, btn_c2, _ = st.columns([1, 1, 4])
            with btn_c1:
                if st.button(f"✅ Approve", key=f"approve_{rid}"):
                    row[rqcm.get("status", 13)].value         = "APPROVED"
                    row[rqcm.get("reviewed_by",   15)].value  = "Forus staff"
                    row[rqcm.get("reviewed_date", 16)].value  = str(datetime.date.today())
                    wb.save(sp())
                    st.session_state["action_log"].append(f"{datetime.date.today()} — Approved {rid}")
                    action_taken = True
                    st.rerun()
            with btn_c2:
                if st.button(f"❌ Reject", key=f"reject_{rid}"):
                    row[rqcm.get("status", 13)].value         = "REJECTED"
                    row[rqcm.get("reviewed_by",   15)].value  = "Forus staff"
                    row[rqcm.get("reviewed_date", 16)].value  = str(datetime.date.today())
                    wb.save(sp())
                    st.session_state["action_log"].append(f"{datetime.date.today()} — Rejected {rid}")
                    action_taken = True
                    st.rerun()

        st.markdown("---")

    if action_taken:
        st.success("✓ Saved. Use **Download updated spreadsheet** in the sidebar to keep your changes.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Apply Approved
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "✅ Apply Approved":
    st.title("Apply Approved Changes")
    st.markdown(
        "Write all **APPROVED** items from the Review Queue back to the "
        "MECHANISMS sheet, and update verification dates."
    )
    require_spreadsheet()

    wb    = _load_wb(data_only=False)
    ws_rq = wb["REVIEW_QUEUE"] if "REVIEW_QUEUE" in wb.sheetnames else None

    if ws_rq is None:
        st.warning("No REVIEW_QUEUE sheet found.")
        st.stop()

    rq_hdrs = [c.value for c in ws_rq[2]]
    rqcm    = {h: i for i, h in enumerate(rq_hdrs) if h}

    approved = []
    for r_idx, row in enumerate(ws_rq.iter_rows(min_row=3, values_only=False), start=3):
        rid = row[rqcm.get("review_id", 1)].value
        if not rid: continue
        if str(row[rqcm.get("status", 13)].value or "").strip().upper() == "APPROVED":
            approved.append((r_idx, row))

    if not approved:
        st.info("No APPROVED items in the Review Queue. Approve items on the **Review Queue** page first.")
        st.stop()

    st.success(f"**{len(approved)} approved item(s)** ready to apply.")

    reviewer = st.text_input("Your name (recorded in spreadsheet)", value="Forus staff")

    if st.button("✅ Apply all approved changes", type="primary"):
        sys.path.insert(0, str(Path(__file__).parent))
        try:
            import generate_toolkit as gt
            gt.SPREADSHEET = sp()
        except ImportError as e:
            st.error(f"Could not import generate_toolkit: {e}")
            st.stop()

        gt.apply_approved(reviewer_name=reviewer or "Forus staff")

        applied = len(approved)
        st.success(
            f"✓ {applied} change(s) applied to MECHANISMS sheet. "
            f"Verification dates updated. Review Queue items marked COMPLETED."
        )
        st.session_state["action_log"].append(
            f"{datetime.date.today()} — Applied {applied} approved change(s)"
        )
        st.info("Use **Download updated spreadsheet** in the sidebar to save your changes, "
                "then regenerate the PDF when ready.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Generate PDF
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "📄 Generate PDF":
    st.title("Generate PDF")
    require_spreadsheet()

    # Check reportlab is available
    try:
        import reportlab
    except ImportError:
        st.error(
            "**reportlab** is not installed. "
            "Run `pip install reportlab` in your environment."
        )
        st.stop()

    sys.path.insert(0, str(Path(__file__).parent))
    try:
        import generate_toolkit as gt
    except ImportError as e:
        st.error(f"Could not import generate_toolkit: {e}")
        st.stop()

    # Patch SPREADSHEET
    gt.SPREADSHEET = sp()

    tab1, tab2 = st.tabs(["Standard build", "Custom member PDF"])

    with tab1:
        st.subheader("Standard build")
        st.markdown(
            "Generates both the **Public** and **Network** versions of the full toolkit PDF."
        )
        col_a, col_b = st.columns(2)
        with col_a:
            build_public  = st.checkbox("Public PDF",  value=True)
        with col_b:
            build_network = st.checkbox("Network PDF (confidential)", value=True)

        if st.button("📄 Build PDF(s)", type="primary"):
            with tempfile.TemporaryDirectory() as tmp_dir:
                v = gt.VERSION
                pub_path = os.path.join(tmp_dir, f"Forus_Toolkit_v{v}_Public.pdf")
                net_path = os.path.join(tmp_dir, f"Forus_Toolkit_v{v}_Network.pdf")
                gt.OUT_PUBLIC  = pub_path
                gt.OUT_NETWORK = net_path

                with st.spinner("Building PDF(s)…"):
                    if build_public:
                        gt.build_pdf(1)
                    if build_network:
                        gt.build_pdf(2)

                dl_c1, dl_c2 = st.columns(2)
                if build_public and os.path.exists(pub_path):
                    with open(pub_path, "rb") as f:
                        dl_c1.download_button(
                            f"⬇ Download Public PDF",
                            f.read(),
                            file_name=f"Forus_Toolkit_v{v}_Public.pdf",
                            mime="application/pdf",
                        )
                if build_network and os.path.exists(net_path):
                    with open(net_path, "rb") as f:
                        dl_c2.download_button(
                            f"⬇ Download Network PDF",
                            f.read(),
                            file_name=f"Forus_Toolkit_v{v}_Network.pdf",
                            mime="application/pdf",
                        )
                st.session_state["action_log"].append(
                    f"{datetime.date.today()} — Generated PDF v{v}"
                )

    with tab2:
        st.subheader("Custom member PDF")
        st.markdown(
            "Generate a personalised PDF for a specific member request "
            "using a row from the **REQUEST_LOG** sheet."
        )

        # Load REQUEST_LOG if it exists
        wb2 = _load_wb()
        req_ids = []
        if "REQUEST_LOG" in wb2.sheetnames:
            ws_rl = wb2["REQUEST_LOG"]
            for row in ws_rl.iter_rows(min_row=5, values_only=True):
                rid = str(row[0] or "").strip()
                if rid and str(row[21] or "").strip().upper() in ("PENDING", ""):
                    name = str(row[2] or "")
                    org  = str(row[3] or "")
                    req_ids.append(f"{rid} — {name}, {org}")
        else:
            st.info("No REQUEST_LOG sheet found. Add one to the spreadsheet to use this feature.")

        if req_ids:
            selected = st.selectbox("Select request", req_ids)
            req_id   = selected.split(" — ")[0].strip() if selected else ""
            access   = st.radio("Access level", ["Public", "Network"])

            if st.button("📄 Build custom PDF", type="primary") and req_id:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    suffix = f"_{req_id}"
                    v = gt.VERSION
                    access_level = 1 if access == "Public" else 2
                    out_path = os.path.join(
                        tmp_dir,
                        f"Forus_Toolkit_v{v}_{'Public' if access_level==1 else 'Network'}{suffix}.pdf"
                    )
                    gt.OUT_PUBLIC  = out_path if access_level == 1 else os.path.join(tmp_dir, "pub.pdf")
                    gt.OUT_NETWORK = out_path if access_level == 2 else os.path.join(tmp_dir, "net.pdf")

                    with st.spinner(f"Building custom PDF for {req_id}…"):
                        gt.build_request_pdf(req_id, access_level=access_level)

                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            st.download_button(
                                f"⬇ Download {req_id} PDF",
                                f.read(),
                                file_name=os.path.basename(out_path),
                                mime="application/pdf",
                            )
                        st.session_state["action_log"].append(
                            f"{datetime.date.today()} — Custom PDF for {req_id}"
                        )
                    else:
                        st.error("PDF generation failed — check that the request ID exists in REQUEST_LOG.")
