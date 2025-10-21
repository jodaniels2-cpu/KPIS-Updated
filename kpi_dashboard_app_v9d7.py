import os
import re
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from datetime import datetime, date, timedelta

st.set_page_config(page_title="Modality Lewisham — KPI Dashboard (v9d7: period totals + weekly units + surname sort)",
                   page_icon="🏥", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
<style>
:root { --ml-primary:#005eb8; --ml-accent:#00a3a3; --ml-muted:#f2f7ff; }
.ml-header{background:var(--ml-muted);border:1px solid #e6eefc;padding:14px 18px;border-radius:16px;display:flex;align-items:center;gap:14px;margin-bottom:10px;}
.ml-pill{background:var(--ml-primary);color:#fff;font-weight:700;padding:4px 10px;border-radius:999px;font-size:12px;letter-spacing:.3px;}
.ml-title{margin:0;font-weight:800;font-size:22px;color:#0b2e59;}
.ml-sub{margin:0;color:#345;font-size:13px;}
.stDataFrame{border-radius:12px;overflow:hidden;border:1px solid #eef3ff;}
.stButton > button, .stDownloadButton > button { border-radius:10px; }
</style>
""", unsafe_allow_html=True,
)

st.markdown(
    """
<div class=\"ml-header\">
  <span class=\"ml-pill\">Modality Lewisham</span>
  <div>
    <p class=\"ml-title\">A‑Team KPI Dashboard</p>
    <p class=\"ml-sub\">Klinik • Docman • Calls — combined & drilldowns</p>
  </div>
</div>
""", unsafe_allow_html=True,
)

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

# Mapping
case_staff_col = find_col(case_df.columns, ["last_archived_by"], contains_any=["archived by","staff","user"])
case_date_col  = find_col(case_df.columns, ["last_archived_date"], contains_any=["archived date","date"])
case_time_col  = find_col(case_df.columns, ["last_archived_time"], contains_any=["archived time","time"])
case_unit_col  = find_col(case_df.columns, ["unit_closed_in","unit"], contains_any=["unit","closed in"])
if any(v is None for v in [case_staff_col, case_date_col, case_unit_col]):
    st.error("Klinik file: couldn't auto-detect required columns."); st.stop()

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

call_user_answered_col = find_col(call_df.columns, ["User Name","Answered","Answered By","Agent"], contains_any=["user name","answered","agent","owner","user"])
call_caller_col        = find_col(call_df.columns, ["Caller Name","Caller"], contains_any=["caller","callback"])
call_type_col          = find_col(call_df.columns, ["Call Type","Type"], contains_any=["type","call type","group","callback","cb"])
call_outcome_col       = find_col(call_df.columns, ["Outcome"], contains_any=["outcome","result","status","disposition"])
call_start_col         = find_col(call_df.columns, ["Start Time","Start","StartDateTime","Start Datetime","Call Start Time","Call Started"], contains_any=["start time","start","begin"])
if any(v is None for v in [call_user_answered_col, call_caller_col, call_type_col, call_start_col]):
    st.error("Calls file: couldn't auto-detect required columns."); st.stop()

# Datasets
start_dt = datetime.combine(start, datetime.min.time())
end_dt = datetime.combine(end, datetime.max.time())

# Klinik
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

# Sorted lists
def sorted_unique(series):
    people = series.dropna().unique().tolist()
    return sort_people_by_surname(people)

k_users = sorted_unique(klinik_all["person"])
d_users = sorted_unique(docman_all["person"])
c_users = sorted_unique(calls_all["person"])
people_all = sort_people_by_surname(sorted(set(k_users) | set(d_users) | set(c_users)))

# Klinik
st.markdown("## Klinik")
if not k_users:
    st.info("No Klinik users found in the selected **date range**.")
else:
    k_user = st.selectbox("Klinik staff member (searchable)", k_users, key="klin_picker")
    kdf = klinik[klinik["person"]==k_user].copy()
    kdf_all = klinik_all[klinik_all["person"]==k_user].copy()
    if kdf_all.empty:
        st.warning("No Klinik cases for this person in the selected date range.")
    else:
        st.caption("Week summary — per day")
        tmp = kdf_all.copy(); tmp["date_only"] = tmp["when"].dt.date
        w = (tmp.groupby("date_only").size().reset_index(name="Klinik Cases")
              .assign(Date=lambda d: pd.to_datetime(d["date_only"]).dt.strftime("%a %d %b"))
              .drop(columns=["date_only"]))
        total = pd.DataFrame([{"Date":"Total", "Klinik Cases":int(w["Klinik Cases"].sum())}])
        st.dataframe(pd.concat([w,total], ignore_index=True), use_container_width=True, height=200)
        valid_days = sorted(kdf_all["when"].dt.date.unique().tolist())
        day_label_map = {d: pd.Timestamp(d).strftime("%A %d %B") for d in valid_days}
        pick = st.selectbox("Choose day", [day_label_map[d] for d in valid_days], key="klin_day")
        day_dt = [d for d,l in day_label_map.items() if l==pick][0]
        day_rows = kdf[kdf["when"].dt.date==day_dt]
        if day_rows.empty:
            st.warning("No Klinik rows for this person on that day with current filters. Try unticking **Limit to 08:00–18:30**.")
        else:
            tmp = day_rows.copy(); tmp["hour"] = tmp["when"].dt.hour
            hours = hour_index_from(tmp, clamp)
            k_pivot = tmp.groupby("hour").size().reindex(hours, fill_value=0).reset_index()
            k_pivot.columns = ["Hour","Klinik Cases"]
            k_pivot["Hour"] = k_pivot["Hour"].apply(lambda h: f"{int(h):02d}:00")
            st.dataframe(k_pivot, use_container_width=True, height=260)
            with st.expander("Units closed in (this day)", expanded=True):
                unit_tbl = (day_rows["unit"].value_counts().reset_index())
                unit_tbl.columns = ["Unit","Cases"]
                st.dataframe(unit_tbl, use_container_width=True)
            with st.expander("Weekly Units closed in — selected user", expanded=False):
                dfw = kdf_all.copy()
                if not dfw.empty:
                    dfw["week_start"] = dfw["when"].dt.to_period("W-MON").apply(lambda p: p.start_time.date())
                    week_unit = (dfw.groupby(["week_start","unit"]).size()
                                   .reset_index(name="Cases")
                                   .pivot(index="week_start", columns="unit", values="Cases")
                                   .fillna(0).astype(int))
                    week_unit.index = pd.to_datetime(week_unit.index).strftime("%d %b %Y (Mon)")
                    st.dataframe(week_unit, use_container_width=True)

st.markdown("---")

# Docman
st.markdown("## Docman")
if not d_users:
    st.info("No Docman users found in the selected **date range**.")
else:
    d_user = st.selectbox("Docman staff member (searchable)", d_users, key="doc_picker")
    ddf = docman[docman["person"]==d_user].copy()
    ddf_all = docman_all[docman_all["person"]==d_user].copy()
    if ddf_all.empty:
        st.warning("No Docman tasks for this person in the selected date range.")
    else:
        st.caption("Week summary — per day")
        tmp = ddf_all.copy(); tmp["date_only"] = tmp["when"].dt.date
        w = (tmp.groupby("date_only").size().reset_index(name="Docman Completed")
              .assign(Date=lambda d: pd.to_datetime(d["date_only"]).dt.strftime("%a %d %b"))
              .drop(columns=["date_only"]))
        total = pd.DataFrame([{"Date":"Total", "Docman Completed":int(w["Docman Completed"].sum())}])
        st.dataframe(pd.concat([w,total], ignore_index=True), use_container_width=True, height=200)
        valid_days = sorted(ddf_all["when"].dt.date.unique().tolist())
        day_label_map = {d: pd.Timestamp(d).strftime("%A %d %B") for d in valid_days}
        pick = st.selectbox("Choose day", [day_label_map[d] for d in valid_days], key="doc_day")
        day_dt = [d for d,l in day_label_map.items() if l==pick][0]
        day_rows = ddf[ddf["when"].dt.date==day_dt]
        if day_rows.empty:
            st.warning("No Docman rows for this person on that day with current filters. Try unticking **Limit to 08:00–18:30**.")
        else:
            tmp = day_rows.copy(); tmp["hour"] = tmp["when"].dt.hour
            hours = hour_index_from(tmp, clamp)
            d_pivot = tmp.groupby("hour").size().reindex(hours, fill_value=0).reset_index()
            d_pivot.columns = ["Hour","Docman Completed"]
            d_pivot["Hour"] = d_pivot["Hour"].apply(lambda h: f"{int(h):02d}:00")
            st.dataframe(d_pivot, use_container_width=True, height=260)

st.markdown("---")

# Calls
st.markdown("## Calls")
if not c_users:
    st.info("No Calls users found in the selected **date range**.")
else:
    c_user = st.selectbox("Calls staff member (searchable)", c_users, key="call_picker",
                          help="For callbacks, 'person' is the caller name; for inbound, it's the 'User Name'.")
    cdf = calls[calls["person"]==c_user].copy()
    cdf_all = calls_all[calls_all["person"]==c_user].copy()
    if cdf_all.empty:
        st.warning("No calls for this person in the selected date range.")
    else:
        st.caption("Week summary — per day")
        day_counts = (cdf_all.assign(label=lambda d: np.where(d["kind"]=="Callback","Callback","Group"))
                      .groupby([cdf_all["when"].dt.date,"label"]).size()
                      .unstack(fill_value=0).reset_index())
        if "Group" not in day_counts.columns: day_counts["Group"]=0
        if "Callback" not in day_counts.columns: day_counts["Callback"]=0
        day_counts["Calls Total"] = day_counts["Group"] + day_counts["Callback"]
        day_counts.rename(columns={"when":"Date"}, inplace=True)
        day_counts["Date"] = pd.to_datetime(day_counts["Date"]).dt.strftime("%a %d %b")
        total_row = pd.DataFrame([{"Date":"Total","Group":int(day_counts["Group"].sum()),
                                   "Callback":int(day_counts["Callback"].sum()), "Calls Total":int(day_counts["Calls Total"].sum())}])
        st.dataframe(pd.concat([day_counts[["Date","Group","Callback","Calls Total"]], total_row], ignore_index=True),
                     use_container_width=True, height=240)
        valid_days = sorted(cdf_all["when"].dt.date.unique().tolist())
        day_label_map = {d: pd.Timestamp(d).strftime("%A %d %B") for d in valid_days}
        pick = st.selectbox("Choose day", [day_label_map[d] for d in valid_days], key="call_day")
        day_dt = [d for d,l in day_label_map.items() if l==pick][0]
        day_rows = cdf[cdf["when"].dt.date==day_dt]
        if day_rows.empty:
            if only_answered:
                st.warning("No calls on that day with current filters. Try unticking **Calls: only include answered/connected** or **Limit to 08:00–18:30**.")
            else:
                st.warning("No calls on that day with current filters. Try unticking **Limit to 08:00–18:30**.")
        else:
            tmp = day_rows.copy(); tmp["hour"] = tmp["when"].dt.hour
            hours = hour_index_from(tmp, clamp)
            c_pivot = (tmp.groupby(["hour","kind"]).size()
                         .unstack(fill_value=0).reindex(index=hours, fill_value=0).reset_index())
            if "Callback" not in c_pivot.columns: c_pivot["Callback"]=0
            if "Group" not in c_pivot.columns: c_pivot["Group"]=0
            c_pivot.rename(columns={"hour":"Hour"}, inplace=True)
            c_pivot["Hour"] = c_pivot["Hour"].apply(lambda h: f"{int(h):02d}:00")
            c_pivot["Total"] = c_pivot["Callback"] + c_pivot["Group"]
            st.dataframe(c_pivot[["Hour","Group","Callback","Total"]], use_container_width=True, height=260)
            fig = go.Figure()
            fig.add_bar(x=c_pivot["Hour"], y=c_pivot["Group"], name="Group")
            fig.add_bar(x=c_pivot["Hour"], y=c_pivot["Callback"], name="Callback")
            fig.update_layout(barmode="stack", title="Calls by hour (selected day)")
            st.plotly_chart(fig, use_container_width=True)

# Combined day picker + activity type
st.markdown("---")
st.markdown("## Combined hourly table — by activity type")
default_sel = []
for key in ["klin_picker","doc_picker","call_picker"]:
    val = st.session_state.get(key)
    if val and val not in default_sel: default_sel.append(val)
if not default_sel:
    default_sel = people_all[:1] if people_all else []

picks = st.multiselect("Staff to include (searchable)", options=people_all, default=default_sel,
                       help="Counts by hour for Klinik / Docman / Calls (Group + Callback) on the chosen day.")
if not picks:
    st.warning("Choose at least one staff member above.")
else:
    combined_all = pd.concat([
        klinik_all.assign(source="Klinik"),
        docman_all.assign(source="Docman"),
        calls_all.assign(source="Calls")
    ], ignore_index=True, sort=False)
    combined_all = combined_all[combined_all["person"].isin(picks)]
    if combined_all.empty:
        st.info("No rows for selected staff in date range.")
    else:
        valid_days = sorted(combined_all["when"].dt.date.unique().tolist())
        day_label_map = {d: pd.Timestamp(d).strftime("%A %d %B") for d in valid_days}
        pick_day = st.selectbox("Choose day", [day_label_map[d] for d in valid_days], key="combined_day_v9d7")
        day_dt = [d for d,l in day_label_map.items() if l==pick_day][0]

        combined = pd.concat([
            klinik.assign(source="Klinik"),
            docman.assign(source="Docman"),
            calls.assign(source="Calls")
        ], ignore_index=True, sort=False)
        sel = combined[(combined["person"].isin(picks)) & (combined["when"].dt.date == day_dt)].copy()
        if sel.empty:
            st.info("No activity for the selected staff on that day with current filters.")
        else:
            sel["hour"] = sel["when"].dt.hour
            hours = hour_index_from(sel, clamp)
            base = (sel[sel["source"].isin(["Klinik","Docman"])]
                    .groupby(["hour","source"]).size()
                    .unstack(fill_value=0)
                    .reindex(index=hours, fill_value=0)
                    .reset_index())
            for col in ["Klinik","Docman"]:
                if col not in base.columns: base[col]=0
            calls_sel = sel[sel["source"]=="Calls"].copy()
            if not calls_sel.empty and "kind" in calls_sel.columns:
                calls_by = (calls_sel.groupby(["hour","kind"]).size()
                                          .unstack(fill_value=0)
                                          .reindex(index=hours, fill_value=0)
                                          .reset_index())
            else:
                calls_by = pd.DataFrame({"hour": hours, "Callback":0, "Group":0})
            for c in ["Callback","Group"]:
                if c not in calls_by.columns: calls_by[c]=0
            out = pd.merge(base, calls_by, on="hour", how="outer").fillna(0)
            out.rename(columns={"hour":"Hour","Group":"Calls_Group","Callback":"Calls_Callback"}, inplace=True)
            out = out.sort_values("Hour")
            out["Hour"] = out["Hour"].apply(lambda h: f"{int(h):02d}:00")
            out["Total"] = out[["Klinik","Docman","Calls_Group","Calls_Callback"]].sum(axis=1).astype(int)
            for c in ["Klinik","Docman","Calls_Group","Calls_Callback","Total"]:
                out[c] = out[c].astype(int)
            st.dataframe(out[["Hour","Klinik","Docman","Calls_Group","Calls_Callback","Total"]], use_container_width=True, height=360)
            fig = go.Figure()
            fig.add_bar(x=out["Hour"], y=out["Klinik"], name="Klinik")
            fig.add_bar(x=out["Hour"], y=out["Docman"], name="Docman")
            fig.add_bar(x=out["Hour"], y=out["Calls_Group"], name="Calls — Group")
            fig.add_bar(x=out["Hour"], y=out["Calls_Callback"], name="Calls — Callback")
            fig.update_layout(barmode="stack", title=f"Hourly activity by type — {pick_day}")
            st.plotly_chart(fig, use_container_width=True)
            csv = out[["Hour","Klinik","Docman","Calls_Group","Calls_Callback","Total"]].to_csv(index=False).encode("utf-8")
            st.download_button("Download hourly activity by type (CSV)", data=csv, file_name="hourly_activity_by_type.csv", mime="text/csv")

    # Period totals + Weekly Units
    st.markdown("### Totals for the whole period (selected staff)")
    combined_period = pd.concat([
        klinik.assign(source="Klinik"),
        docman.assign(source="Docman"),
        calls.assign(source="Calls")
    ], ignore_index=True, sort=False)
    selp = combined_period[combined_period["person"].isin(picks)].copy()
    if selp.empty:
        st.info("No activity for the selected staff in this period with current filters.")
    else:
        # Totals by day
        selp["date"] = selp["when"].dt.date
        by_day = (selp.assign(kind_full=np.where(selp["source"]=="Calls", "Calls_"+selp.get("kind","").astype(str), selp["source"]))
                        .assign(kind_full=lambda d: d["kind_full"].replace({"Calls_nan":"Calls"}))
                        .groupby(["date","kind_full"]).size()
                        .unstack(fill_value=0).reset_index())
        for col in ["Klinik","Docman","Calls_Group","Calls_Callback","Calls"]:
            if col not in by_day.columns: by_day[col]=0
        by_day["Total"] = by_day.get("Klinik",0)+by_day.get("Docman",0)+by_day.get("Calls_Group",0)+by_day.get("Calls_Callback",0)+by_day.get("Calls",0)
        by_day["Date"] = pd.to_datetime(by_day["date"]).dt.strftime("%a %d %b")
        cols_day = ["Date","Klinik","Docman","Calls_Group","Calls_Callback"]
        if "Calls" in by_day.columns: cols_day.append("Calls")
        cols_day.append("Total")
        st.dataframe(by_day[cols_day], use_container_width=True)

        # Totals by hour across entire period
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
        cols_hr = ["Hour","Klinik","Docman","Calls_Group","Calls_Callback"]
        if "Calls" in by_hour.columns: cols_hr.append("Calls")
        cols_hr.append("Total")
        st.dataframe(by_hour[cols_hr], use_container_width=True)

        # Weekly Units Closed In — combined
        st.markdown("### Weekly Units closed in — combined (selected staff)")
        klinik_picks = klinik_all[klinik_all["person"].isin(picks)].copy()
        if klinik_picks.empty:
            st.info("No Klinik cases for the selected staff in this period.")
        else:
            klinik_picks["week_start"] = klinik_picks["when"].dt.to_period("W-MON").apply(lambda p: p.start_time.date())
            wk = (klinik_picks.groupby(["week_start","unit"]).size()
                                .reset_index(name="Cases")
                                .pivot(index="week_start", columns="unit", values="Cases")
                                .fillna(0).astype(int))
            wk.index = pd.to_datetime(wk.index).strftime("%d %b %Y (Mon)")
            st.dataframe(wk, use_container_width=True)
