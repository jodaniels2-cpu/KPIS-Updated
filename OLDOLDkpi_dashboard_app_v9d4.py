
import os
import re
from io import BytesIO
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from datetime import datetime, date, timedelta

st.set_page_config(page_title="Modality Lewisham — KPI Dashboard (v9d8: XLSX export)",
                   page_icon="🏥", layout="wide", initial_sidebar_state="expanded")

# ---------- Branding ----------
st.markdown("""
<style>
:root { --ml-primary:#005eb8; --ml-accent:#00a3a3; --ml-muted:#f2f7ff; }
.ml-header{background:var(--ml-muted);border:1px solid #e6eefc;padding:14px 18px;border-radius:16px;display:flex;align-items:center;gap:14px;margin-bottom:10px;}
.ml-pill{background:var(--ml-primary);color:#fff;font-weight:700;padding:4px 10px;border-radius:999px;font-size:12px;letter-spacing:.3px;}
.ml-title{margin:0;font-weight:800;font-size:22px;color:#0b2e59;}
.ml-sub{margin:0;color:#345;font-size:13px;}
.stDataFrame{border-radius:12px;overflow:hidden;border:1px solid #eef3ff;}
.stButton > button, .stDownloadButton > button { border-radius:10px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ml-header">
  <span class="ml-pill">Modality Lewisham</span>
  <div>
    <p class="ml-title">A‑Team KPI Dashboard</p>
    <p class="ml-sub">Klinik • Docman • Calls — with XLSX export (all raw sheets + totals)</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ---------- Auth ----------
def _get_app_password():
    try:
        if "auth" in st.secrets and "password" in st.secrets["auth"]:
            return st.secrets["auth"]["password"]
    except Exception:
        pass
    return os.getenv("APP_PASSWORD", None)

def password_gate():
    pw = _get_app_password()
    if not pw:
        st.error("No dashboard password is configured. Set **auth.password** in Secrets or **APP_PASSWORD** env var.")
        st.stop()
    if st.session_state.get("authenticated", False):
        return True
    with st.form("login", clear_on_submit=False):
        st.subheader("🔒 Enter password to continue")
        entered = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Unlock")
    if submitted:
        if entered == pw:
            st.session_state["authenticated"] = True
            st.success("Unlocked")
            return True
        else:
            st.error("Incorrect password")
    st.stop()

password_gate()

# ---------- Helpers ----------
@st.cache_data
def load_table(file):
    if file is None:
        return None
    name = getattr(file, "name", "").lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        file.seek(0)
        try:
            return pd.read_excel(file, engine="openpyxl")
        except Exception:
            file.seek(0); return pd.read_excel(file)
    for kwargs in [{"sep":",","encoding":"utf-8"},{"sep":";","encoding":"utf-8"},{"sep":",","encoding":"latin-1"},{"sep":";","encoding":"latin-1"}]:
        try:
            file.seek(0); return pd.read_csv(file, **kwargs)
        except Exception:
            continue
    file.seek(0); return pd.read_csv(file)

def parse_dt(date_series, time_series=None):
    if time_series is None:
        return pd.to_datetime(date_series, errors="coerce", dayfirst=True, infer_datetime_format=True)
    ds = pd.Series(date_series).astype(str).str.strip()
    ts = pd.Series(time_series).astype(str).str.strip()
    s = (ds + " " + ts).str.replace(r"\s+", " ", regex=True)
    out = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
    if out.notna().mean() < 0.5:
        alt = pd.to_datetime(date_series, errors="coerce", dayfirst=True, infer_datetime_format=True)
        out = out.where(out.notna(), alt)
    return out

def find_col(cols, targets=None, contains_any=None):
    targets = targets or []
    low=[c.lower() for c in cols]
    for t in targets:
        if t.lower() in low: return cols[low.index(t.lower())]
    if contains_any:
        for i,c in enumerate(low):
            if any(tok.lower() in c for tok in contains_any): return cols[i]
    return None

def normalise_name_or_email(x):
    if pd.isna(x): return None
    s = str(x).strip()
    return s.lower() if "@" in s else s

def hour_index_from(df, clamp):
    if df.empty:
        return []
    if clamp:
        return list(range(8,19))
    vals = sorted(df["when"].dt.hour.dropna().unique().tolist())
    return vals if vals else list(range(0,24))

def surname_key(s: str):
    if s is None: return ""
    s2 = str(s).strip()
    if "@" in s2:
        local = s2.split("@",1)[0]
        parts = re.split(r"[._\-+]+", local)
        return parts[-1].lower() if parts else s2.lower()
    parts = s2.split()
    return parts[-1].lower() if parts else s2.lower()

def sort_people_by_surname(people):
    return sorted(people, key=surname_key)

# ---------- Sidebar (global filters) ----------
with st.sidebar:
    st.header("Global filters")
    today=date.today()
    default_start=today - timedelta(days=today.weekday())
    default_end=default_start + timedelta(days=4)
    start=st.date_input("Start date", value=default_start)
    end=st.date_input("End date", value=default_end)
    if end<start: st.error("End date must be on or after start date"); st.stop()
    clamp = st.checkbox("Limit to 08:00–18:30", value=True)
    only_answered = st.checkbox("Calls: only include answered/connected", value=True)
    st.subheader("Upload files (all three required)")
    case_file=st.file_uploader("Klinik Case Counts (CSV/XLSX)", type=["csv","xlsx","xls"])
    doc_file=st.file_uploader("Docman Tasks (CSV/XLSX)", type=["csv","xlsx","xls"])
    call_file=st.file_uploader("Telephone Calls Export (CSV/XLSX)", type=["csv","xlsx","xls"])

case_df=load_table(case_file) if case_file else None
doc_df=load_table(doc_file) if doc_file else None
call_df=load_table(call_file) if call_file else None
if any(x is None for x in (case_df, doc_df, call_df)):
    st.info("Upload **all three** files to continue."); st.stop()

# ---------- Column mapping per source ----------
# Klinik
case_staff_col = find_col(case_df.columns, ["last_archived_by"], contains_any=["archived by","staff","user"])
case_date_col  = find_col(case_df.columns, ["last_archived_date"], contains_any=["archived date","date"])
case_time_col  = find_col(case_df.columns, ["last_archived_time"], contains_any=["archived time","time"])
case_unit_col  = find_col(case_df.columns, ["unit_closed_in","unit"], contains_any=["unit","closed in"])
if any(v is None for v in [case_staff_col, case_date_col, case_unit_col]):
    st.error("Klinik file: couldn't auto-detect required columns."); st.stop()

# Docman
doc_user_col = find_col(doc_df.columns, ["User","Completed User"], contains_any=["user","completed"])
doc_dt_col   = find_col(doc_df.columns, ["Date and Time of Event"], contains_any=["date and time of event","completed","datetime","date","time"])
if doc_dt_col is None:
    for c in doc_df.columns:
        try:
            test = pd.to_datetime(doc_df[c], errors="coerce")
            if test.notna().mean()>0.7:
                doc_dt_col=c; break
        except Exception:
            pass
if any(v is None for v in [doc_user_col, doc_dt_col]):
    st.error("Docman file: couldn't auto-detect required columns."); st.stop()

# Calls
call_user_answered_col = find_col(call_df.columns, ["User Name","Answered","Answered By","Agent"], contains_any=["user name","answered","agent","owner","user"])
call_caller_col        = find_col(call_df.columns, ["Caller Name","Caller"], contains_any=["caller","callback"])
call_type_col          = find_col(call_df.columns, ["Call Type","Type"], contains_any=["type","call type","group","callback","cb"])
call_outcome_col       = find_col(call_df.columns, ["Outcome"], contains_any=["outcome","result","status","disposition"])
call_start_col         = find_col(call_df.columns, ["Start Time","Start","StartDateTime","Start Datetime","Call Start Time","Call Started"], contains_any=["start time","start","begin"])
if any(v is None for v in [call_user_answered_col, call_caller_col, call_type_col, call_start_col]):
    st.error("Calls file: couldn't auto-detect required columns."); st.stop()

# ---------- Build per-source events ----------
start_dt = datetime.combine(start, datetime.min.time())
end_dt = datetime.combine(end, datetime.max.time())

# Klinik (picker from date-range only)
case_dt_all = parse_dt(case_df[case_date_col], case_df[case_time_col] if case_time_col in case_df.columns else None) if case_time_col else parse_dt(case_df[case_date_col])
klinik_all = pd.DataFrame({
    "when": case_dt_all,
    "person": case_df[case_staff_col].map(normalise_name_or_email),
    "unit": case_df[case_unit_col].astype(str).replace(["nan","None",""], np.nan).fillna("Unknown")
}).dropna(subset=["when","person"])
klinik_all = klinik_all[(klinik_all["when"]>=pd.Timestamp(start_dt)) & (klinik_all["when"]<=pd.Timestamp(end_dt))]
klinik = klinik_all.copy()
if clamp:
    klinik = klinik[((klinik["when"].dt.hour > 8) | ((klinik["when"].dt.hour==8) & (klinik["when"].dt.minute>=0))) &
                    ((klinik["when"].dt.hour < 18) | ((klinik["when"].dt.hour==18) & (klinik["when"].dt.minute<=30)))]

# Docman
doc_dt_all = parse_dt(doc_df[doc_dt_col], None)
docman_all = pd.DataFrame({
    "when": doc_dt_all,
    "person": doc_df[doc_user_col].map(normalise_name_or_email)
}).dropna(subset=["when","person"])
docman_all = docman_all[(docman_all["when"]>=pd.Timestamp(start_dt)) & (docman_all["when"]<=pd.Timestamp(end_dt))]
docman = docman_all.copy()
if clamp:
    docman = docman[((docman["when"].dt.hour > 8) | ((docman["when"].dt.hour==8) & (docman["when"].dt.minute>=0))) &
                    ((docman["when"].dt.hour < 18) | ((docman["when"].dt.hour==18) & (docman["when"].dt.minute<=30)))]

# Calls
call_dt_all = parse_dt(call_df[call_start_col], None)
type_lower = call_df[call_type_col].astype(str).str.lower()
is_callback = type_lower.str.contains("callback|call back|cb", na=False)
who_all = np.where(is_callback, call_df[call_caller_col].astype(str), call_df[call_user_answered_col].astype(str))

calls_all = pd.DataFrame({
    "when": call_dt_all,
    "person": pd.Series(who_all).map(normalise_name_or_email),
    "kind": np.where(is_callback, "Callback","Group"),
    "outcome": call_df[call_outcome_col].astype(str).str.lower() if call_outcome_col else ""
}).dropna(subset=["when","person"])
calls_all = calls_all[(calls_all["when"]>=pd.Timestamp(start_dt)) & (calls_all["when"]<=pd.Timestamp(end_dt))]
calls = calls_all.copy()
if only_answered and call_outcome_col is not None:
    answered_mask = calls["outcome"].str.contains("answer|answered|connect|connected|complete|completed|handled|finished|resolved|success|ok", na=False)
    calls = calls[answered_mask]
if clamp:
    calls = calls[((calls["when"].dt.hour > 8) | ((calls["when"].dt.hour==8) & (calls["when"].dt.minute>=0))) &
                  ((calls["when"].dt.hour < 18) | ((calls["when"].dt.hour==18) & (calls["when"].dt.minute<=30)))]

# ---------- Surname-sorted staff lists ----------
def sorted_unique(series):
    people = series.dropna().unique().tolist()
    return sort_people_by_surname(people)

k_users = sorted_unique(klinik_all["person"])
d_users = sorted_unique(docman_all["person"])
c_users = sorted_unique(calls_all["person"])
people_all = sort_people_by_surname(sorted(set(k_users) | set(d_users) | set(c_users)))

# ---------- Utility to build period totals for selected staff ----------
def build_period_totals(selected_people):
    combined_period = pd.concat([
        klinik.assign(source="Klinik"),
        docman.assign(source="Docman"),
        calls.assign(source="Calls")
    ], ignore_index=True, sort=False)
    selp = combined_period[combined_period["person"].isin(selected_people)].copy()
    if selp.empty:
        return (pd.DataFrame(), pd.DataFrame())

    # By day
    selp["date"] = selp["when"].dt.date
    by_day = (selp.assign(kind_full=np.where(selp["source"]=="Calls", "Calls_"+selp.get("kind","").astype(str), selp["source"]))
                    .assign(kind_full=lambda d: d["kind_full"].replace({"Calls_nan":"Calls"}))
                    .groupby(["date","kind_full"]).size()
                    .unstack(fill_value=0).reset_index())
    for col in ["Klinik","Docman","Calls_Group","Calls_Callback","Calls"]:
        if col not in by_day.columns: by_day[col]=0
    by_day["Total"] = by_day.get("Klinik",0)+by_day.get("Docman",0)+by_day.get("Calls_Group",0)+by_day.get("Calls_Callback",0)+by_day.get("Calls",0)
    by_day["Date"] = pd.to_datetime(by_day["date"]).dt.strftime("%a %d %b")

    # By hour across entire period
    selp["hour"] = selp["when"].dt.hour
    hours = hour_index_from(selp, clamp)
    by_hour = (selp.assign(kind_full=np.where(selp["source"]=="Calls", "Calls_"+selp.get("kind","").astype(str), selp["source"]))
                    .assign(kind_full=lambda d: d["kind_full"].replace({"Calls_nan":"Calls"}))
                    .groupby(["hour","kind_full"]).size()
                    .unstack(fill_value=0).reindex(index=hours, fill_value=0).reset_index())
    for col in ["Klinik","Docman","Calls_Group","Calls_Callback","Calls"]:
        if col not in by_hour.columns: by_hour[col]=0
    by_hour.rename(columns={"hour":"Hour"}, inplace=True)
    by_hour["Hour"] = by_hour["Hour"].apply(lambda h: f"{int(h):02d}:00")
    by_hour["Total"] = by_hour.get("Klinik",0)+by_hour.get("Docman",0)+by_hour.get("Calls_Group",0)+by_hour.get("Calls_Callback",0)+by_hour.get("Calls",0)

    return (by_day, by_hour)

# ---------- UI sections (compact) ----------
st.markdown("### Combined export area")
default_sel = []
for key in ["klin_picker","doc_picker","call_picker"]:
    val = st.session_state.get(key)
    if val and val not in default_sel: default_sel.append(val)
if not default_sel:
    default_sel = people_all[:1] if people_all else []

picks = st.multiselect("Staff to include in XLSX (searchable)", options=people_all, default=default_sel)

# Weekly Units Closed In — combined (for XLSX)
def weekly_units_for(kl_df):
    if kl_df.empty:
        return pd.DataFrame()
    df = kl_df.copy()
    df["week_start"] = df["when"].dt.to_period("W-MON").apply(lambda p: p.start_time.date())
    wk = (df.groupby(["week_start","unit"]).size()
                    .reset_index(name="Cases")
                    .pivot(index="week_start", columns="unit", values="Cases")
                    .fillna(0).astype(int))
    wk.index = pd.to_datetime(wk.index).strftime("%d %b %Y (Mon)")
    return wk.reset_index().rename(columns={"index":"Week"})

if st.button("Build Excel preview tables"):
    st.success("Tables prepared below. Use the Download XLSX button to save everything.")

# Build all tables for export (uses current filters)
combined_all = pd.concat([
    klinik.assign(source="Klinik"),
    docman.assign(source="Docman"),
    calls.assign(source="Calls")
], ignore_index=True, sort=False)

by_day_sel, by_hour_sel = build_period_totals(picks) if picks else (pd.DataFrame(), pd.DataFrame())
wk_units_combined = weekly_units_for(klinik_all[klinik_all["person"].isin(picks)]) if picks else pd.DataFrame()

# Show previews
with st.expander("Preview: Period totals by day (selected staff)", expanded=False):
    if by_day_sel.empty:
        st.info("No data for selected staff.")
    else:
        st.dataframe(by_day_sel, use_container_width=True, height=260)

with st.expander("Preview: Period totals by hour (selected staff)", expanded=False):
    if by_hour_sel.empty:
        st.info("No data for selected staff.")
    else:
        st.dataframe(by_hour_sel, use_container_width=True, height=260)

with st.expander("Preview: Weekly Units closed in — combined (selected staff)", expanded=False):
    if wk_units_combined.empty:
        st.info("No Klinik cases for selected staff.")
    else:
        st.dataframe(wk_units_combined, use_container_width=True)

# ---------- Build XLSX in-memory and offer download ----------
def to_excel_bytes():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Raw (all staff) — respect current filters
        ka = klinik_all.copy(); ka.sort_values("when", inplace=True)
        da = docman_all.copy(); da.sort_values("when", inplace=True)
        ca = calls_all.copy();  ca.sort_values("when", inplace=True)
        ka.to_excel(writer, sheet_name="Klinik_Period_All", index=False)
        da.to_excel(writer, sheet_name="Docman_Period_All", index=False)
        ca.to_excel(writer, sheet_name="Calls_Period_All", index=False)

        # Raw (selected staff only)
        if picks:
            ks = klinik_all[klinik_all["person"].isin(picks)].copy()
            ds = docman_all[docman_all["person"].isin(picks)].copy()
            cs = calls_all[calls_all["person"].isin(picks)].copy()
            ks.to_excel(writer, sheet_name="Klinik_Selected", index=False)
            ds.to_excel(writer, sheet_name="Docman_Selected", index=False)
            cs.to_excel(writer, sheet_name="Calls_Selected", index=False)

        # Totals for whole period (selected staff)
        if not by_day_sel.empty:
            by_day_sel.to_excel(writer, sheet_name="Totals_By_Day_Selected", index=False)
        if not by_hour_sel.empty:
            by_hour_sel.to_excel(writer, sheet_name="Totals_By_Hour_Selected", index=False)

        # Weekly units (selected staff combined)
        if not wk_units_combined.empty:
            wk_units_combined.to_excel(writer, sheet_name="Weekly_Units_Selected", index=False)

        # Combined hourly by task for a chosen day is dynamic in UI; skip in export to avoid confusion.
        # You can add a day picker export later if needed.

    output.seek(0)
    return output.getvalue()

xlsx_bytes = to_excel_bytes()
st.download_button(
    "⬇️ Download full KPI workbook (XLSX)",
    data=xlsx_bytes,
    file_name="Modality_Lewisham_KPI_v9d8.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Includes raw Klinik/Docman/Calls (all + selected staff) and all totals."
)

