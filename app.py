import streamlit as st
import pandas as pd
from datetime import date, time, datetime
from report_generator import generate_report, create_excel, detect_issues

st.set_page_config(
    page_title="Work Time Report",
    page_icon="⏱",
    layout="wide",
)

st.markdown("""
<style>
    [data-testid="stMetricValue"] { font-size: 2rem; font-weight: 700; }
    div.stDownloadButton > button {
        background: #1F4E79; color: white; font-weight: 600;
        font-size: 1.05rem; border-radius: 8px; padding: 0.6rem 1rem;
    }
    div.stDownloadButton > button:hover { background: #2E75B6; }
</style>
""", unsafe_allow_html=True)

st.title("⏱ Work Time Report")
st.caption("Upload → Review issues → Manual adjustments → Generate → Download")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙ Shift Settings")

    night_start = st.time_input(
        "Night Shift Start",
        value=time(0, 0),
        help="Night shift (default 00:00–09:00). Check-ins up to 2h before "
             "the start (e.g. 23:40) count for the next day's night shift.",
    )
    am_start = st.time_input(
        "AM Shift Start",
        value=time(5, 30),
        help="Fixed start time applied when first check-in is before 12:00",
    )
    pm_start = st.time_input(
        "PM Shift Start",
        value=time(13, 0),
        help="Fixed start time applied when first check-in is at or after 12:00",
    )
    grace_minutes = st.number_input(
        "Late grace period (minutes)",
        min_value=0, max_value=30, value=0, step=1,
        help="0 = anyone clocking in 1 minute after shift start is LATE. "
             "E.g. 5 = up to 5 min late is still counted as on time.",
    )
    break_hours = st.number_input(
        "Break Duration (hours)",
        min_value=0.0, max_value=4.0,
        value=1.0, step=0.5,
        help="Deducted from every day's total",
    )

    st.divider()
    st.markdown(f"""
**Shift logic** (by first check-in)
- Night → start **{night_start.strftime('%H:%M')}**
- AM (before 12:00) → start **{am_start.strftime('%H:%M')}**
- PM (12:00 or later) → start **{pm_start.strftime('%H:%M')}**

**Lateness**
- Clock in on time / early → counted from shift start
- Clock in later than start{f' + {grace_minutes} min' if grace_minutes else ''} → **LATE**, counted from actual clock-in

**Formula**
`Last Clock Out − Counted From − {break_hours}h break`

**Overtime** threshold: 40 h / week
""")
    st.divider()
    st.caption("v3.1 — Night shift + late arrivals")

# ── File upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload clock machine Excel file (.xlsx)",
    type=["xlsx"],
)

if not uploaded_file:
    st.info("Upload a file to get started.")
    with st.expander("Expected column format"):
        st.markdown("""
| Column | Example values |
|---|---|
| **PunchCardTime** | `2026-06-01 07:02:35` |
| **EmployeeName** | `John Doe` |
| **WarehouseEntry/ExitIdentifier** | `Check-in` / `Check-out` or `0` / `1` |
        """)
    st.stop()

st.success(f"✓  **{uploaded_file.name}** loaded")

with st.expander("Preview raw data (first 20 rows)"):
    try:
        df_raw = pd.read_excel(uploaded_file)
        st.dataframe(df_raw.head(20), use_container_width=True)
        st.caption("Detected columns: " + ", ".join(f"`{c}`" for c in df_raw.columns))
        uploaded_file.seek(0)
    except Exception as exc:
        st.error(f"Cannot read file: {exc}")

st.divider()

# ── Auto issue detection (once per file) ──────────────────────────────────────
file_key = f"{uploaded_file.name}|{night_start}"
if st.session_state.get("_file_key") != file_key:
    try:
        issues, employees = detect_issues(
            uploaded_file,
            night_start_hour=night_start.hour,
            night_start_min=night_start.minute,
        )
        uploaded_file.seek(0)
        if st.session_state.get("_file_name") != uploaded_file.name:
            st.session_state.manual_entries = []   # new file → clear manual entries
        st.session_state._file_name     = uploaded_file.name
        st.session_state._file_key      = file_key
        st.session_state._issues        = issues
        st.session_state._employees     = employees
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

issues    = st.session_state._issues
employees = st.session_state._employees

# ── Issues panel ──────────────────────────────────────────────────────────────
if issues:
    st.subheader(f"⚠ Issues Found — {len(issues)} incomplete record(s)")
    st.dataframe(
        pd.DataFrame(issues),
        use_container_width=True,
        hide_index=True,
    )
    st.caption("Add the missing entries below before generating the report.")
else:
    st.success("✓ All records have complete check-in / check-out pairs")

st.divider()

# ── Manual adjustments ────────────────────────────────────────────────────────
st.subheader("✏ Manual Adjustments")
st.caption("Add missing clock-in or clock-out entries. They will be merged with the file data before calculation. "
           "For a night-shift check-in before midnight (e.g. 23:50), pick the actual calendar date of that check-in.")

with st.form("add_manual_entry", clear_on_submit=True):
    c1, c2, c3, c4 = st.columns([3, 2, 2, 2])
    sel_emp  = c1.selectbox("Employee",  employees if employees else ["—"])
    sel_date = c2.date_input("Date",     value=date.today())
    sel_type = c3.selectbox("Type",      ["Check-in", "Check-out"])
    sel_time = c4.time_input("Time",     value=time(6, 0))
    add_btn  = st.form_submit_button("➕ Add Entry", use_container_width=True)

if add_btn and employees:
    st.session_state.manual_entries.append({
        "employee": sel_emp,
        "datetime": datetime.combine(sel_date, sel_time),
        "type":     sel_type,
    })
    st.rerun()

if st.session_state.manual_entries:
    st.write(f"**{len(st.session_state.manual_entries)} manual entr(ies) added:**")

    header = st.columns([3, 2, 2, 2, 1])
    header[0].markdown("**Employee**")
    header[1].markdown("**Date**")
    header[2].markdown("**Time**")
    header[3].markdown("**Type**")
    header[4].markdown("**Del**")

    to_delete = []
    for i, entry in enumerate(st.session_state.manual_entries):
        c1, c2, c3, c4, c5 = st.columns([3, 2, 2, 2, 1])
        c1.write(entry["employee"])
        c2.write(str(entry["datetime"].date()))
        c3.write(entry["datetime"].strftime("%H:%M"))
        c4.write(entry["type"])
        if c5.button("🗑", key=f"del_{i}", help="Delete"):
            to_delete.append(i)

    if to_delete:
        for i in sorted(to_delete, reverse=True):
            st.session_state.manual_entries.pop(i)
        st.rerun()
else:
    st.caption("No manual entries added.")

st.divider()

# ── Generate report ───────────────────────────────────────────────────────────
if st.button("Generate Report", type="primary", use_container_width=True):

    uploaded_file.seek(0)

    with st.spinner("Processing punch-card data…"):
        try:
            daily_df, weekly_df, monthly_df, total_hours, skipped = generate_report(
                uploaded_file,
                am_start_hour=am_start.hour,
                am_start_min=am_start.minute,
                pm_start_hour=pm_start.hour,
                pm_start_min=pm_start.minute,
                break_hours=break_hours,
                manual_entries=st.session_state.manual_entries,
                night_start_hour=night_start.hour,
                night_start_min=night_start.minute,
                grace_minutes=int(grace_minutes),
            )
        except ValueError as exc:
            st.error(str(exc))
            st.stop()
        except Exception as exc:
            st.error(f"Unexpected error: {exc}")
            with st.expander("Details"):
                st.exception(exc)
            st.stop()

    st.success("Report generated successfully!")

    late_count = int((daily_df["Status"] == "LATE").sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Records",       len(daily_df))
    c2.metric("Employees",     daily_df["Employee"].nunique())
    c3.metric("Days",          daily_df["Date"].nunique())
    c4.metric("Total Hours",   f"{total_hours:.1f} h")
    c5.metric("Late Arrivals", late_count)

    if skipped:
        with st.expander(f"⚠ {len(skipped)} record(s) still incomplete — excluded from report"):
            for s in skipped:
                st.text(s)

    def _highlight_late(row):
        if row["Status"] == "LATE":
            return ["background-color: #FFE4B5; color: #9C0006; font-weight: 600"] * len(row)
        return [""] * len(row)

    st.subheader("Daily Report")
    if late_count:
        st.warning(f"⚠  {late_count} late arrival(s) — highlighted in orange")
    st.dataframe(daily_df.style.apply(_highlight_late, axis=1),
                 use_container_width=True, hide_index=True)

    st.subheader("Weekly Summary")
    ot_count = (weekly_df["Overtime Hours"] != "—").sum()
    if ot_count:
        st.warning(f"⚠  {ot_count} employee-week(s) exceed 40 h")
    st.dataframe(weekly_df, use_container_width=True, hide_index=True)

    st.subheader("Monthly Summary")
    mot_count = (monthly_df["Overtime Hours"] != "—").sum()
    if mot_count:
        st.warning(f"⚠  {mot_count} employee-month(s) exceed 160 h")
    st.dataframe(monthly_df, use_container_width=True, hide_index=True)

    st.divider()
    excel_buf = create_excel(daily_df, weekly_df, monthly_df)
    filename  = f"worktime_report_{date.today().strftime('%Y%m%d')}.xlsx"

    st.download_button(
        "⬇  Download Excel Report  (Daily + Late + Weekly + Monthly)",
        data=excel_buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
