import pandas as pd
from datetime import datetime, time, timedelta
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# Night-shift check-ins made up to this many minutes BEFORE the night shift
# starts (e.g. 23:40 for a 00:00 start) are counted on the next day's shift.
NIGHT_EARLY_CHECKIN_WINDOW_MIN = 120


def _find_column(df, *keyword_groups):
    for keywords in keyword_groups:
        for col in df.columns:
            col_lower = str(col).lower()
            if any(kw in col_lower for kw in keywords):
                return col
    return None


def _detect_columns(df):
    time_col = _find_column(df,
        ["punchcardtime", "punch_card_time"],
        ["punchcard", "punch_card"],
        ["timestamp", "datetime"],
        ["time"],
    )
    entry_exit_col = _find_column(df,
        ["warehouseentry", "exitidentifier"],
        ["identifier", "entry/exit", "entryexit"],
        ["entry", "exit"],
        ["checkin", "check-in", "check_in"],
    )
    employee_col = _find_column(df,
        ["employeename", "employee_name"],
        ["employee"],
        ["name", "worker", "staff"],
    )
    return time_col, entry_exit_col, employee_col


def _validate_columns(df, time_col, entry_exit_col, employee_col):
    missing = []
    if not time_col:       missing.append("PunchCard Time")
    if not entry_exit_col: missing.append("Entry/Exit Identifier")
    if not employee_col:   missing.append("Employee Name")
    if missing:
        available = ", ".join(str(c) for c in df.columns)
        raise ValueError(
            f"Could not detect column(s): {', '.join(missing)}\n"
            f"Available columns: {available}"
        )


def _is_checkin(value):
    val = str(value).lower().strip()
    return "check-in" in val or "checkin" in val or val == "0"


def _is_checkout(value):
    val = str(value).lower().strip()
    return "check-out" in val or "checkout" in val or val == "1"


def _fmt_td(td):
    total_sec = int(td.total_seconds())
    sign = "-" if total_sec < 0 else ""
    total_sec = abs(total_sec)
    h, rem = divmod(total_sec, 3600)
    m, s = divmod(rem, 60)
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"


def _fmt_hours(decimal_hours):
    """Convert decimal hours to 'Xh Ym' string (e.g. 2.5 → '2h 30m')."""
    if decimal_hours <= 0:
        return "—"
    total_min = round(decimal_hours * 60)
    h, m = divmod(total_min, 60)
    return f"{h}h {m:02d}m"


def _night_start_offset(night_start_hour, night_start_min):
    """Night shift start in minutes relative to midnight of the shift date.
    00:00 → 0, 23:00 → -60 (starts the evening before)."""
    mins = night_start_hour * 60 + night_start_min
    return mins - 1440 if mins >= 720 else mins


def _assign_shift_dates(df, time_col, entry_exit_col,
                        night_start_hour=0, night_start_min=0):
    """Adds a '_Date' column = the date of the shift each punch belongs to.
    Check-ins made shortly before a night shift that starts around midnight
    (e.g. 23:45 for a 00:00 start) are moved to the next day."""
    df[time_col] = pd.to_datetime(df[time_col])
    df["_Date"] = df[time_col].dt.date

    rollover = _night_start_offset(night_start_hour, night_start_min) \
        - NIGHT_EARLY_CHECKIN_WINDOW_MIN
    if rollover < 0:                     # window reaches into the previous evening
        rollover_min_of_day = rollover + 1440
        mins = df[time_col].dt.hour * 60 + df[time_col].dt.minute
        mask = df[entry_exit_col].apply(_is_checkin) & (mins >= rollover_min_of_day)
        if mask.any():
            df["_Date"] = df["_Date"].astype(object)
            df.loc[mask, "_Date"] = (
                df.loc[mask, time_col] + pd.Timedelta(days=1)
            ).dt.date
    return df


def detect_issues(input_file, night_start_hour=0, night_start_min=0):
    """Quick parse — returns (issues_list, sorted_employees_list)."""
    df = pd.read_excel(input_file)
    time_col, entry_exit_col, employee_col = _detect_columns(df)
    _validate_columns(df, time_col, entry_exit_col, employee_col)

    df = _assign_shift_dates(df, time_col, entry_exit_col,
                             night_start_hour, night_start_min)

    issues = []
    for (day, emp), group in df.groupby(["_Date", employee_col]):
        has_in  = group[entry_exit_col].apply(_is_checkin).any()
        has_out = group[entry_exit_col].apply(_is_checkout).any()
        if not has_in:
            issues.append({"Date": str(day), "Employee": emp, "Problem": "Missing Check-in"})
        elif not has_out:
            issues.append({"Date": str(day), "Employee": emp, "Problem": "Missing Check-out"})

    employees = sorted(df[employee_col].dropna().unique().tolist())
    return issues, employees


def generate_report(
    input_file,
    am_start_hour=5, am_start_min=30,
    pm_start_hour=13, pm_start_min=0,
    break_hours=1.0,
    manual_entries=None,
    night_start_hour=0, night_start_min=0,
    grace_minutes=0,
):
    df = pd.read_excel(input_file)
    time_col, entry_exit_col, employee_col = _detect_columns(df)
    _validate_columns(df, time_col, entry_exit_col, employee_col)

    # Merge manual entries into the dataframe before processing
    if manual_entries:
        manual_rows = [
            {
                time_col:       pd.Timestamp(e["datetime"]),
                entry_exit_col: e["type"],
                employee_col:   e["employee"],
            }
            for e in manual_entries
        ]
        df = pd.concat([df, pd.DataFrame(manual_rows)], ignore_index=True)

    df = _assign_shift_dates(df, time_col, entry_exit_col,
                             night_start_hour, night_start_min)

    # Shift starts in minutes from midnight of the shift date
    night_rel = _night_start_offset(night_start_hour, night_start_min)
    am_rel    = am_start_hour * 60 + am_start_min
    pm_rel    = pm_start_hour * 60 + pm_start_min
    night_am_boundary = (night_rel + am_rel) / 2   # e.g. 00:00 & 06:00 → 03:00
    am_pm_boundary    = 12 * 60                    # check-in ≥ 12:00 → PM

    break_label = f"{int(break_hours)}h" if break_hours == int(break_hours) else f"{break_hours}h"
    total_col   = f"Total Time (-{break_label} break)"

    rows, skipped = [], []

    for (day, emp), group in df.groupby(["_Date", employee_col]):
        group = group.sort_values(time_col)

        checkins  = group[group[entry_exit_col].apply(_is_checkin)]
        checkouts = group[group[entry_exit_col].apply(_is_checkout)]

        if checkins.empty:
            skipped.append(f"{day}  /  {emp}  →  no check-in found")
            continue
        if checkouts.empty:
            skipped.append(f"{day}  /  {emp}  →  no check-out found")
            continue

        day_start = datetime.combine(day, time(0, 0))
        first_in  = checkins[time_col].min().to_pydatetime()
        last_out  = checkouts[time_col].max().to_pydatetime()
        in_rel    = (first_in - day_start).total_seconds() / 60

        if in_rel < night_am_boundary:
            shift, start_rel = "Night", night_rel
        elif in_rel < am_pm_boundary:
            shift, start_rel = "AM", am_rel
        else:
            shift, start_rel = "PM", pm_rel
        start_dt = day_start + timedelta(minutes=start_rel)

        # Lateness — compared at minute level: 06:00:59 = on time, 06:01 = late
        first_in_min = first_in.replace(second=0, microsecond=0)
        late_by_min  = int((first_in_min - start_dt).total_seconds() // 60)
        is_late      = late_by_min > grace_minutes
        counted_from = first_in_min if is_late else start_dt

        if last_out <= counted_from:
            skipped.append(f"{day}  /  {emp}  →  check-out ({last_out:%H:%M}) is before shift start")
            continue

        total_delta     = last_out - counted_from - timedelta(hours=break_hours)
        total_hours_val = total_delta.total_seconds() / 3600

        in_label = first_in.strftime("%H:%M:%S")
        if first_in.date() < day:
            in_label += " (prev day)"

        rows.append({
            "Date":           day,
            "Employee":       emp,
            "Shift":          shift,
            "Shift Start":    start_dt.strftime("%H:%M"),
            "First Clock In": in_label,
            "Status":         "LATE" if is_late else "On time",
            "Late By":        _fmt_hours(late_by_min / 60) if is_late else "—",
            "Counted From":   counted_from.strftime("%H:%M"),
            "Last Clock Out": last_out.strftime("%H:%M:%S"),
            total_col:        _fmt_td(total_delta),
            "_total_hours":   round(total_hours_val, 4),
            "_late":          int(is_late),
            "_late_min":      late_by_min if is_late else 0,
        })

    if not rows:
        raise ValueError(
            "No valid records found.\n"
            "Make sure the file contains Check-in and Check-out entries."
        )

    daily = (
        pd.DataFrame(rows)
        .sort_values(["Date", "Employee"])
        .reset_index(drop=True)
    )

    # ── Weekly summary ────────────────────────────────────────────────────────
    dates = pd.to_datetime(daily["Date"])
    iso   = dates.dt.isocalendar()
    daily["_year"]       = iso.year.astype(int)
    daily["_week"]       = iso.week.astype(int)
    daily["_week_start"] = dates - pd.to_timedelta(dates.dt.dayofweek, unit="D")

    weekly = (
        daily
        .groupby(["_year", "_week", "_week_start", "Employee"])
        .agg(total=("_total_hours", "sum"),
             late_days=("_late", "sum"),
             late_min=("_late_min", "sum"))
        .reset_index()
    )
    weekly.columns = ["Year", "Week", "Week Start", "Employee",
                      "Total Hours", "Late Days", "_late_min"]
    weekly["Total Hours"]    = weekly["Total Hours"].round(2)
    weekly["Regular Hours"]  = weekly["Total Hours"].apply(lambda x: min(x, 40))
    weekly["Overtime Hours"] = weekly["Total Hours"].apply(lambda x: max(0.0, x - 40))
    weekly["Total Hours"]    = weekly["Total Hours"].apply(_fmt_hours)
    weekly["Regular Hours"]  = weekly["Regular Hours"].apply(_fmt_hours)
    weekly["Overtime Hours"] = weekly["Overtime Hours"].apply(_fmt_hours)
    weekly["Total Late"]     = (weekly["_late_min"] / 60).apply(_fmt_hours)
    weekly["Week Start"]     = weekly["Week Start"].dt.date
    weekly = weekly.sort_values(["Year", "Week", "Employee"]).reset_index(drop=True)

    daily_export  = daily[["Date", "Employee", "Shift", "Shift Start",
                           "First Clock In", "Status", "Late By", "Counted From",
                           "Last Clock Out", total_col]].copy()
    weekly_export = weekly[["Week Start", "Year", "Week", "Employee",
                            "Total Hours", "Regular Hours", "Overtime Hours",
                            "Late Days", "Total Late"]].copy()

    # ── Monthly summary ───────────────────────────────────────────────────────
    dates_m = pd.to_datetime(daily["Date"])
    daily["_month_num"]  = dates_m.dt.month
    daily["_month_name"] = dates_m.dt.strftime("%B %Y")

    monthly = (
        daily
        .groupby(["_year", "_month_num", "_month_name", "Employee"])
        .agg(days_worked=("_total_hours", "count"),
             total_hours=("_total_hours", "sum"),
             late_days=("_late", "sum"),
             late_min=("_late_min", "sum"))
        .reset_index()
    )
    monthly["total_hours"]    = monthly["total_hours"].round(2)
    monthly["regular_hours"]  = monthly["total_hours"].apply(lambda x: min(x, 160))
    monthly["overtime_hours"] = monthly["total_hours"].apply(lambda x: max(0.0, x - 160))
    monthly["total_hours"]    = monthly["total_hours"].apply(_fmt_hours)
    monthly["regular_hours"]  = monthly["regular_hours"].apply(_fmt_hours)
    monthly["overtime_hours"] = monthly["overtime_hours"].apply(_fmt_hours)
    monthly["total_late"]     = (monthly["late_min"] / 60).apply(_fmt_hours)
    monthly = monthly.sort_values(["_year", "_month_num", "Employee"]).reset_index(drop=True)

    monthly_export = monthly.rename(columns={
        "_month_name":    "Month",
        "_year":          "Year",
        "days_worked":    "Days Worked",
        "total_hours":    "Total Hours",
        "regular_hours":  "Regular Hours",
        "overtime_hours": "Overtime Hours",
        "late_days":      "Late Days",
        "total_late":     "Total Late",
    })[["Month", "Year", "Employee", "Days Worked", "Total Hours",
        "Regular Hours", "Overtime Hours", "Late Days", "Total Late"]].copy()

    return daily_export, weekly_export, monthly_export, daily["_total_hours"].sum(), skipped


# ── Excel export ──────────────────────────────────────────────────────────────

_HEADER_FILL   = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
_HEADER_FONT   = Font(color="FFFFFF", bold=True)
_OVERTIME_FILL = PatternFill(start_color="FFD7D7", end_color="FFD7D7", fill_type="solid")
_LATE_FILL     = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
_LATE_FONT     = Font(color="C00000", bold=True)
_CENTER        = Alignment(horizontal="center", vertical="center")


def _style_sheet(ws, df, overtime_col=None, late_status_col=None, late_count_col=None):
    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = _HEADER_FILL
        cell.font      = _HEADER_FONT
        cell.alignment = _CENTER

    # Overtime rows → light red
    if overtime_col:
        ot_idx = df.columns.get_loc(overtime_col) + 1
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=ot_idx).value
            if val not in (None, "—", "", 0):
                for ci in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=ci).fill = _OVERTIME_FILL

    # Late rows (daily) → orange row, red bold "LATE"
    if late_status_col:
        st_idx = df.columns.get_loc(late_status_col) + 1
        late_by_idx = df.columns.get_loc("Late By") + 1 if "Late By" in df.columns else None
        for row_idx in range(2, ws.max_row + 1):
            if ws.cell(row=row_idx, column=st_idx).value == "LATE":
                for ci in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=ci).fill = _LATE_FILL
                ws.cell(row=row_idx, column=st_idx).font = _LATE_FONT
                if late_by_idx:
                    ws.cell(row=row_idx, column=late_by_idx).font = _LATE_FONT

    # Late counts (weekly / monthly) → highlight cell when > 0
    if late_count_col:
        lc_idx = df.columns.get_loc(late_count_col) + 1
        cols_to_mark = [lc_idx]
        if "Total Late" in df.columns:
            cols_to_mark.append(df.columns.get_loc("Total Late") + 1)
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=lc_idx).value
            if isinstance(val, (int, float)) and val > 0:
                for ci in cols_to_mark:
                    c = ws.cell(row=row_idx, column=ci)
                    c.fill = _LATE_FILL
                    c.font = _LATE_FONT

    for col_idx, col_name in enumerate(df.columns, 1):
        max_len = len(str(col_name))
        for row_idx in range(2, ws.max_row + 1):
            v = ws.cell(row=row_idx, column=col_idx).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20


def create_excel(daily_df, weekly_df, monthly_df):
    late_df = daily_df[daily_df["Status"] == "LATE"].reset_index(drop=True)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        daily_df.to_excel(writer,   sheet_name="Daily Report",    index=False)
        late_df.to_excel(writer,    sheet_name="Late Arrivals",   index=False)
        weekly_df.to_excel(writer,  sheet_name="Weekly Summary",  index=False)
        monthly_df.to_excel(writer, sheet_name="Monthly Summary", index=False)
        _style_sheet(writer.sheets["Daily Report"],    daily_df,  late_status_col="Status")
        _style_sheet(writer.sheets["Late Arrivals"],   late_df,   late_status_col="Status")
        _style_sheet(writer.sheets["Weekly Summary"],  weekly_df,
                     overtime_col="Overtime Hours", late_count_col="Late Days")
        _style_sheet(writer.sheets["Monthly Summary"], monthly_df,
                     overtime_col="Overtime Hours", late_count_col="Late Days")
    buf.seek(0)
    return buf
