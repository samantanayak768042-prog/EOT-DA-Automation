"""
===============================================================================
  PRIMAVERA P6 EOT (Extension of Time) AUTOMATION – PURE SQLITE VERSION
===============================================================================
This script performs Extension of Time (EOT) analysis for Primavera P6 by
manipulating the SQLite database directly. It copies projects, inserts delay
activities, maps predecessor relationships, and calculates the early/late
start and finish dates mathematically using a custom Critical Path Method (CPM)
engine in Python.

Outputs a single consolidated Excel report (in the standardized Milestone vs
Delay format) displaying the impacted dates for all specified activities
imported from an Excel file (or matching the standard milestone pattern).

PREREQUISITES
──────────────
    pip install openpyxl pandas

HOW TO RUN
──────────
     python db_eot_automation.py
==============================================================================="""

import argparse
import base64
import logging
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import uuid
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG & PATHS
# ─────────────────────────────────────────────────────────────────────────────

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB_PATH = r"C:\Users\91943\Downloads\T3.db"
OUTPUT_FILE = os.path.join(_SCRIPT_DIR, "EOTR.xlsx")
LOG_FILE = os.path.join(_SCRIPT_DIR, "eot_process_logs.log")

# Regex to identify milestone task codes like MS-01, MS 02, T3-CA-MS-01, etc.
MS_PATTERN = re.compile(r"MS[-_ ]?(\d+)", re.IGNORECASE)

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────


def configure_logging(log_file: str):
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        encoding="utf-8",
        force=True,
    )


def log(msg: str):
    print(msg)
    logging.info(msg)


def err(msg: str):
    print(f"[ERROR] {msg}")
    logging.error(msg)


# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE GENERATOR (Consolidated from standalone generate_template.py)
# ─────────────────────────────────────────────────────────────────────────────


def create_template(output_path: str = None) -> str:
    """Generates the sample Delay Events input Excel template."""
    data = {
        "Delay Event ID": ["DE1001", "DE1002"],
        "Description of Event": [
            "Delay in Handover of Land at Zone A",
            "Drawing approval delay for Foundation Works",
        ],
        "Notification Date": ["2026-06-01", "2026-06-15"],
        "Delay in days": [15, 30],
        "Activity Impacted": ["T3-ACT-1010", "T3-ACT-1020, T3-ACT-1030"],
        "Relationship": ["FS", "SS"],
        "Lag": [0, 5],
        "Primary Constraint": ["Start On or After", "Start On or After"],
    }
    df = pd.DataFrame(data)
    if not output_path:
        output_path = os.path.join(_SCRIPT_DIR, "delay_events_template.xlsx")
    df.to_excel(output_path, index=False)
    print(f"Template successfully generated at: {output_path}")
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# ARGUMENT PARSING
# ─────────────────────────────────────────────────────────────────────────────


def parse_args():
    parser = argparse.ArgumentParser(
        description="Primavera P6 Pure SQLite EOT automation (No GUI)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--excel", "-e", required=False,
        help="Path to the Excel file containing delay events",
    )
    parser.add_argument(
        "--source-project", "-s", required=False,
        help="Exact baseline project name or ID (PROJ_SHORT_NAME) in P6",
    )
    parser.add_argument(
        "--db-path", "-d", default=DEFAULT_DB_PATH,
        help="Path to the Primavera SQLite database file (.db)",
    )
    parser.add_argument(
        "--output", "-o", default=OUTPUT_FILE,
        help="Optional explicit path for the consolidated output Excel file",
    )
    parser.add_argument(
        "--log-file", default=LOG_FILE,
        help="Path for the automation log file",
    )
    parser.add_argument(
        "--default-remarks", default="",
        help="Default remarks to apply for every delay event",
    )
    parser.add_argument(
        "--milestone-mode", choices=["standard", "custom"], default=None,
        help="Activity/Milestone selection mode: standard (Level 1 MS-XX) or custom (imported Excel list/pattern)",
    )
    parser.add_argument(
        "--milestone-config", "--activities-excel", "-a", dest="milestone_config", default=None,
        help="Path to the Excel file containing the list of specified activities (or custom milestone pattern)",
    )
    parser.add_argument(
        "--generate-template", nargs="?", const="", default=None,
        help="Generate a sample delay events Excel template and exit",
    )
    parser.add_argument(
        "--self-test", action="store_true",
        help="Run built-in regression test verification suite and exit",
    )
    return parser.parse_args()


# ─────────────────────────────────────────────────────────────────────────────
# SQLITE DATA MANIPULATION HELPERS
# ─────────────────────────────────────────────────────────────────────────────


def generate_p6_guid():
    u = uuid.uuid4()
    return base64.b64encode(u.bytes).decode("utf-8").rstrip("=")


def allocate_keys(conn, key_name, count):
    if count <= 0:
        return []
    cur = conn.cursor()
    cur.execute("SELECT KEY_SEQ_NUM FROM NEXTKEY WHERE KEY_NAME = ?", (key_name,))
    row = cur.fetchone()
    if not row:
        raise ValueError(f"Key name {key_name} not found in NEXTKEY")
    start_key = row[0]
    next_key = start_key + count
    cur.execute("UPDATE NEXTKEY SET KEY_SEQ_NUM = ? WHERE KEY_NAME = ?", (next_key, key_name))
    return list(range(start_key, next_key))


def insert_row(cur, table, data_dict):
    keys = list(data_dict.keys())
    values = [data_dict[k] for k in keys]
    placeholders = ", ".join(["?"] * len(keys))
    sql = f"INSERT INTO {table} ({', '.join(keys)}) VALUES ({placeholders})"
    cur.execute(sql, values)


def cleanup_project_tables(cur, proj_id):
    """Safely removes all records for a simulated project ID across P6 tables."""
    try:
        cur.execute("DELETE FROM SUMTASK WHERE PROJ_ID = ?", (proj_id,))
    except sqlite3.Error:
        pass
    try:
        cur.execute("DELETE FROM SCHEDOPTIONS WHERE PROJ_ID = ?", (proj_id,))
    except sqlite3.Error:
        pass
    try:
        cur.execute("DELETE FROM TASKRSRC WHERE PROJ_ID = ?", (proj_id,))
    except sqlite3.Error:
        pass
    cur.execute("DELETE FROM TASKPRED WHERE PROJ_ID = ? OR PRED_PROJ_ID = ?", (proj_id, proj_id))
    cur.execute("DELETE FROM TASK WHERE PROJ_ID = ?", (proj_id,))
    cur.execute("DELETE FROM PROJWBS WHERE PROJ_ID = ?", (proj_id,))
    try:
        cur.execute("DELETE FROM CALENDAR WHERE PROJ_ID = ? AND CLNDR_TYPE = 'CA_Proj'", (proj_id,))
    except sqlite3.Error:
        pass
    try:
        cur.execute("DELETE FROM PROJSHAR WHERE PROJ_ID = ?", (proj_id,))
    except sqlite3.Error:
        pass
    try:
        cur.execute("DELETE FROM PROJPROP WHERE PROJ_ID = ?", (proj_id,))
    except sqlite3.Error:
        pass
    cur.execute("DELETE FROM PROJECT WHERE PROJ_ID = ?", (proj_id,))


def parse_p6_datetime(date_str, default_hour=8):
    """Parses P6 SQLite date/datetime strings across varying formats."""
    if not date_str:
        return None
    if isinstance(date_str, datetime):
        return date_str
    clean = str(date_str).strip().split(".")[0]
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(clean, fmt)
            if fmt == "%Y-%m-%d":
                dt = dt.replace(hour=default_hour, minute=0, second=0, microsecond=0)
            return dt
        except ValueError:
            continue
    dt = pd.to_datetime(clean).to_pydatetime()
    return dt


# ─────────────────────────────────────────────────────────────────────────────
# CALENDAR RESOLUTION & UTILITIES
# ─────────────────────────────────────────────────────────────────────────────


def load_and_resolve_calendars(conn):
    cur = conn.cursor()
    cur.execute("SELECT CLNDR_ID, CLNDR_NAME, BASE_CLNDR_ID, CLNDR_DATA FROM CALENDAR")
    raw_calendars = {}
    for cid, name, bid, data in cur.fetchall():
        raw_calendars[cid] = {
            "name": name,
            "base_id": bid,
            "data": data.decode("utf-8", errors="ignore") if isinstance(data, bytes) else data,
        }

    def parse_calendar_data(data):
        if not data:
            return {"working_weekdays": [2, 3, 4, 5, 6], "exceptions": set()}
        working_weekdays = []
        for day in range(1, 8):
            if f"{day}()()" not in data:
                working_weekdays.append(day)
        exceptions = set()
        d_numbers = re.findall(r"d\|(\d+)", data)
        excel_base = datetime(1899, 12, 30)
        for num in d_numbers:
            dt = excel_base + timedelta(days=int(num))
            exceptions.add(dt.date())
        return {
            "working_weekdays": working_weekdays,
            "exceptions": exceptions,
        }

    resolved = {}

    def resolve_calendar(cid, visited=None):
        if cid in resolved:
            return resolved[cid]
        if visited is None:
            visited = set()
        if cid in visited:
            return {"working_weekdays": [2, 3, 4, 5, 6], "exceptions": set()}
        visited.add(cid)

        raw = raw_calendars.get(cid)
        if not raw:
            return {"working_weekdays": [2, 3, 4, 5, 6], "exceptions": set()}
        parsed = parse_calendar_data(raw["data"])
        base_info = None
        if raw["base_id"] and raw["base_id"] != cid:
            base_info = resolve_calendar(raw["base_id"], visited)
        all_exceptions = set(parsed["exceptions"])
        if base_info:
            all_exceptions.update(base_info["exceptions"])
        if parsed["working_weekdays"]:
            working_weekdays = parsed["working_weekdays"]
        elif base_info:
            working_weekdays = base_info["working_weekdays"]
        else:
            working_weekdays = [2, 3, 4, 5, 6]
        res = {
            "working_weekdays": working_weekdays,
            "exceptions": all_exceptions,
        }
        resolved[cid] = res
        return res

    for cid in raw_calendars:
        resolve_calendar(cid)
    return resolved


def is_workday(dt, clndr_id, resolved_calendars):
    cal_info = resolved_calendars.get(clndr_id) if resolved_calendars else None
    if not cal_info:
        return dt.weekday() < 5
    py_wd = dt.weekday()
    p6_wd = 1 if py_wd == 6 else py_wd + 2
    if p6_wd not in cal_info["working_weekdays"]:
        return False
    if dt.date() in cal_info["exceptions"]:
        return False
    return True


def add_workdays_cal(start_dt, days, clndr_id, resolved_calendars):
    dt = start_dt
    if days > 0:
        added = 0
        while added < days:
            dt += timedelta(days=1)
            if is_workday(dt, clndr_id, resolved_calendars):
                added += 1
    elif days < 0:
        subtracted = 0
        while subtracted < abs(days):
            dt -= timedelta(days=1)
            if is_workday(dt, clndr_id, resolved_calendars):
                subtracted += 1
    return dt


def get_next_workday_start_cal(dt, clndr_id, resolved_calendars):
    res = dt.replace(hour=8, minute=0, second=0, microsecond=0)
    res += timedelta(days=1)
    while not is_workday(res, clndr_id, resolved_calendars):
        res += timedelta(days=1)
    return res


def add_workdays(start_dt, days):
    dt = start_dt
    if days > 0:
        added = 0
        while added < days:
            dt += timedelta(days=1)
            if dt.weekday() < 5:  # Monday-Friday
                added += 1
    elif days < 0:
        subtracted = 0
        while subtracted < abs(days):
            dt -= timedelta(days=1)
            if dt.weekday() < 5:
                subtracted += 1
    return dt


def map_constraint_type(cstr_str: str) -> str:
    if not cstr_str:
        return "CS_MSOA"
    val = str(cstr_str).strip().lower()
    if "start on or after" in val or "start no earlier than" in val or "snet" in val:
        return "CS_MSOA"
    if "start on or before" in val or "start no later than" in val or "snlt" in val:
        return "CS_MSOB"
    if "finish on or after" in val or "finish no earlier than" in val or "fnet" in val:
        return "CS_MFOA"
    if "finish on or before" in val or "finish no later than" in val or "fnlt" in val:
        return "CS_MFOB"
    if "must start on" in val or "mso" in val:
        return "CS_MSO"
    if "must finish on" in val or "mfo" in val:
        return "CS_MFO"
    return "CS_MSOA"


def duplicate_project_in_db(conn, source_short_name, new_short_name):
    cur = conn.cursor()

    # 1. Find source project ID
    cur.execute(
        "SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
        (source_short_name,),
    )
    row = cur.fetchone()
    if not row:
        raise ValueError(f"Source project '{source_short_name}' not found in database")
    source_proj_id = row[0]

    # Check if destination project already exists. If so, delete it first to get a clean slate.
    cur.execute("SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ?", (new_short_name,))
    exists_row = cur.fetchone()
    if exists_row:
        proj_id_to_del = exists_row[0]
        log(f"  Project '{new_short_name}' already exists in database. Recreating for clean slate...")
        try:
            cur.execute("DELETE FROM SUMTASK WHERE PROJ_ID = ?", (proj_id_to_del,))
        except sqlite3.Error:
            pass
        cur.execute("DELETE FROM TASKPRED WHERE PROJ_ID = ?", (proj_id_to_del,))
        cur.execute("DELETE FROM TASK WHERE PROJ_ID = ?", (proj_id_to_del,))
        cur.execute("DELETE FROM PROJWBS WHERE PROJ_ID = ?", (proj_id_to_del,))
        cur.execute("DELETE FROM PROJSHAR WHERE PROJ_ID = ?", (proj_id_to_del,))
        cur.execute("DELETE FROM PROJPROP WHERE PROJ_ID = ?", (proj_id_to_del,))
        cur.execute("DELETE FROM PROJECT WHERE PROJ_ID = ?", (proj_id_to_del,))

    log(f"  [DB] Replicating project: {source_short_name} (ID: {source_proj_id}) -> {new_short_name}...")

    # 2. Allocate project ID
    new_proj_id = allocate_keys(conn, "project_proj_id", 1)[0]

    # 3. Copy PROJECT row (preserving ORIG_PROJ_ID / SOURCE_PROJ_ID as in baseline so P6 shows it directly on Refresh)
    cur.execute("SELECT * FROM PROJECT WHERE PROJ_ID = ?", (source_proj_id,))
    proj_row = cur.fetchone()
    cols = [c[0] for c in cur.description]
    proj_data = dict(zip(cols, proj_row))
    proj_data["PROJ_ID"] = new_proj_id
    proj_data["PROJ_SHORT_NAME"] = new_short_name
    proj_data["GUID"] = generate_p6_guid()
    if "CREATE_DATE" in proj_data:
        proj_data["CREATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if "UPDATE_DATE" in proj_data:
        proj_data["UPDATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    insert_row(cur, "PROJECT", proj_data)

    # 4. Copy PROJPROP rows
    cur.execute("SELECT * FROM PROJPROP WHERE PROJ_ID = ?", (source_proj_id,))
    prop_cols = [c[0] for c in cur.description]
    prop_rows = cur.fetchall()
    for r in prop_rows:
        prop_data = dict(zip(prop_cols, r))
        prop_data["PROJ_ID"] = new_proj_id
        insert_row(cur, "PROJPROP", prop_data)

    # 5. Copy PROJSHAR rows
    cur.execute("SELECT * FROM PROJSHAR WHERE PROJ_ID = ?", (source_proj_id,))
    shar_cols = [c[0] for c in cur.description]
    shar_rows = cur.fetchall()
    for r in shar_rows:
        shar_data = dict(zip(shar_cols, r))
        shar_data["PROJ_ID"] = new_proj_id
        insert_row(cur, "PROJSHAR", shar_data)

    # 6. Copy WBS rows
    cur.execute("SELECT * FROM PROJWBS WHERE PROJ_ID = ?", (source_proj_id,))
    wbs_rows = cur.fetchall()
    wbs_cols = [c[0] for c in cur.description]

    # Allocate new WBS IDs, including one extra for the "Delay Events" WBS node
    new_wbs_ids = allocate_keys(conn, "projwbs_wbs_id", len(wbs_rows) + 1)
    delay_wbs_id = new_wbs_ids[-1]

    wbs_id_map = {}
    for old_wbs_row, new_wbs_id in zip(wbs_rows, new_wbs_ids[:-1]):
        old_wbs_id = old_wbs_row[wbs_cols.index("WBS_ID")]
        wbs_id_map[old_wbs_id] = new_wbs_id

    new_root_wbs_id = None
    root_wbs_obs_id = None

    for old_wbs_row in wbs_rows:
        wbs_data = dict(zip(wbs_cols, old_wbs_row))
        old_id = wbs_data["WBS_ID"]
        wbs_data["WBS_ID"] = wbs_id_map[old_id]
        wbs_data["PROJ_ID"] = new_proj_id
        wbs_data["GUID"] = generate_p6_guid()
        if "CREATE_DATE" in wbs_data:
            wbs_data["CREATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if "UPDATE_DATE" in wbs_data:
            wbs_data["UPDATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        old_parent = wbs_data["PARENT_WBS_ID"]
        if old_parent in wbs_id_map:
            wbs_data["PARENT_WBS_ID"] = wbs_id_map[old_parent]

        if wbs_data["PROJ_NODE_FLAG"] == "Y":
            wbs_data["WBS_SHORT_NAME"] = new_short_name
            wbs_data["WBS_NAME"] = f"{new_short_name} - {old_wbs_row[wbs_cols.index('WBS_NAME')]}"
            wbs_data["SUM_DATA_FLAG"] = "Y"
            new_root_wbs_id = wbs_data["WBS_ID"]
            root_wbs_obs_id = wbs_data.get("OBS_ID")

        insert_row(cur, "PROJWBS", wbs_data)

    # Insert new "Delay Events" WBS node under the root project WBS node
    new_wbs_data = {
        "WBS_ID": delay_wbs_id,
        "PROJ_ID": new_proj_id,
        "OBS_ID": root_wbs_obs_id,
        "SEQ_NUM": 10000,
        "EST_WT": 1.0,
        "PROJ_NODE_FLAG": "N",
        "SUM_DATA_FLAG": "N",
        "STATUS_CODE": "WS_Open",
        "WBS_SHORT_NAME": "DE",
        "WBS_NAME": "Delay Events",
        "PARENT_WBS_ID": new_root_wbs_id,
        "GUID": generate_p6_guid(),
        "CREATE_DATE": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "UPDATE_DATE": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "CREATE_USER": "admin",
        "UPDATE_USER": "admin",
    }
    new_wbs_data = {col: new_wbs_data[col] for col in wbs_cols if col in new_wbs_data}
    insert_row(cur, "PROJWBS", new_wbs_data)

    # 7. Copy TASK rows
    cur.execute("SELECT * FROM TASK WHERE PROJ_ID = ?", (source_proj_id,))
    task_rows = cur.fetchall()
    task_cols = [c[0] for c in cur.description]

    new_task_ids = allocate_keys(conn, "task_task_id", len(task_rows))
    task_id_map = {}
    for old_task_row, new_task_id in zip(task_rows, new_task_ids):
        old_task_id = old_task_row[task_cols.index("TASK_ID")]
        task_id_map[old_task_id] = new_task_id

    for old_task_row in task_rows:
        task_data = dict(zip(task_cols, old_task_row))
        old_tid = task_data["TASK_ID"]
        task_data["TASK_ID"] = task_id_map[old_tid]
        task_data["PROJ_ID"] = new_proj_id
        if "CREATE_DATE" in task_data:
            task_data["CREATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if "UPDATE_DATE" in task_data:
            task_data["UPDATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        old_wbs = task_data["WBS_ID"]
        if old_wbs in wbs_id_map:
            task_data["WBS_ID"] = wbs_id_map[old_wbs]

        task_data["GUID"] = generate_p6_guid()
        insert_row(cur, "TASK", task_data)

    # 8. Copy TASKPRED rows
    cur.execute("SELECT * FROM TASKPRED WHERE PROJ_ID = ? AND TASK_PRED_ID != 40924", (source_proj_id,))
    pred_rows = cur.fetchall()
    pred_cols = [c[0] for c in cur.description]

    new_pred_ids = allocate_keys(conn, "taskpred_task_pred_id", len(pred_rows))
    for old_pred_row, new_pred_id in zip(pred_rows, new_pred_ids):
        pred_data = dict(zip(pred_cols, old_pred_row))
        pred_data["TASK_PRED_ID"] = new_pred_id
        pred_data["PROJ_ID"] = new_proj_id
        pred_data["PRED_PROJ_ID"] = new_proj_id
        if "CREATE_DATE" in pred_data:
            pred_data["CREATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if "UPDATE_DATE" in pred_data:
            pred_data["UPDATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        old_tid = pred_data["TASK_ID"]
        old_pred_tid = pred_data["PRED_TASK_ID"]

        if old_tid in task_id_map:
            pred_data["TASK_ID"] = task_id_map[old_tid]
        if old_pred_tid in task_id_map:
            pred_data["PRED_TASK_ID"] = task_id_map[old_pred_tid]

        insert_row(cur, "TASKPRED", pred_data)

    return new_proj_id, delay_wbs_id


def add_delay_event_to_project(
    conn, proj_id, delay_wbs_id, delay_row, resolved_cols, source_proj_id, resolved_calendars=None
):
    cur = conn.cursor()

    # Retrieve columns for TASK and TASKPRED tables
    cur.execute("SELECT * FROM TASK LIMIT 1")
    task_cols = [c[0] for c in cur.description]
    cur.execute("SELECT * FROM TASKPRED LIMIT 1")
    pred_cols = [c[0] for c in cur.description]

    # 1. Parse successor codes
    try:
        raw_val = delay_row[resolved_cols["Activity Impacted"]]
        if pd.isna(raw_val):
            raise ValueError("Activity Impacted is empty")
        raw_impacted = str(raw_val).strip()
        impacted_act_codes = [act.strip() for act in re.split(r"[\n\r,;]+", raw_impacted) if act.strip()]
        if not impacted_act_codes:
            raise ValueError(f"No impacted activity codes found in row: '{raw_impacted}'")

        # Verify all impacted activities exist and retrieve their baseline task details
        imp_data_list = []
        for code in impacted_act_codes:
            cur.execute("SELECT * FROM TASK WHERE PROJ_ID = ? AND TASK_CODE = ?", (source_proj_id, code))
            imp_row = cur.fetchone()
            if not imp_row:
                raise ValueError(f"Impacted activity '{code}' not found in baseline project ID {source_proj_id}")
            imp_data_list.append(dict(zip(task_cols, imp_row)))
    except Exception as e:
        raise ValueError(f"[Activity Impacted] {e}")

    # Use first impacted activity to copy default settings (like Calendar ID)
    imp_data = imp_data_list[0]

    # Retrieve the baseline project's default calendar ID
    cur.execute("SELECT CLNDR_ID FROM PROJECT WHERE PROJ_ID = ?", (source_proj_id,))
    proj_row = cur.fetchone()
    baseline_clndr_id = proj_row[0] if (proj_row and proj_row[0]) else imp_data["CLNDR_ID"]

    delay_id = str(delay_row[resolved_cols["Delay Event ID"]]).strip()
    delay_name = str(delay_row[resolved_cols["Description of Event"]]).strip()

    try:
        raw_dt = delay_row[resolved_cols["Notification Date"]]
        if pd.isna(raw_dt) or str(raw_dt).strip() == "":
            raise ValueError("Notification Date is empty")
        start_dt = pd.to_datetime(raw_dt)
        if pd.isna(start_dt):
            raise ValueError(f"Could not parse date: {raw_dt}")
    except Exception as e:
        raise ValueError(f"[Notification Date] {e}")

    try:
        raw_dur = delay_row[resolved_cols["Delay in days"]]
        if pd.isna(raw_dur) or str(raw_dur).strip() == "":
            raise ValueError("Delay in days is empty")
        duration_days = int(float(raw_dur))
    except Exception as e:
        raise ValueError(f"[Delay in days] {e}")
    duration_hours = max(8.0, float(duration_days * 8.0))

    if resolved_calendars is None:
        resolved_calendars = load_and_resolve_calendars(conn)
    if duration_days > 0:
        finish_dt = add_workdays_cal(
            start_dt, duration_days - 1, baseline_clndr_id, resolved_calendars
        ).replace(hour=17, minute=0, second=0)
    else:
        finish_dt = start_dt.replace(hour=17, minute=0, second=0)

    primary_cstr_col = resolved_cols.get("Primary Constraint")
    try:
        if primary_cstr_col and primary_cstr_col in delay_row and pd.notna(delay_row[primary_cstr_col]):
            cstr_type = map_constraint_type(delay_row[primary_cstr_col])
        else:
            cstr_type = "CS_MSOA"
    except Exception as e:
        raise ValueError(f"[Primary Constraint] {e}")

    try:
        raw_rel = delay_row[resolved_cols["Relationship"]]
        if pd.isna(raw_rel) or str(raw_rel).strip() == "":
            raise ValueError("Relationship is empty")
        rel_clean = str(raw_rel).strip().upper()
        if rel_clean.startswith("PR_"):
            rel_clean = rel_clean[3:]
        if rel_clean not in ("FS", "SS", "FF", "SF"):
            raise ValueError(f"Unsupported relationship type '{raw_rel}'. Expected FS, SS, FF, or SF.")
        rel_type = "PR_" + rel_clean
    except Exception as e:
        raise ValueError(f"[Relationship] {e}")

    try:
        lag_col = resolved_cols.get("Lag")
        if lag_col and lag_col in delay_row and pd.notna(delay_row[lag_col]) and str(delay_row[lag_col]).strip() != "":
            lag_days = float(delay_row[lag_col])
        else:
            lag_days = 0.0
    except Exception as e:
        raise ValueError(f"[Lag] {e}")
    lag_hours = lag_days * 8.0

    # 2. Allocate keys: 1 task ID and len(impacted_act_codes) taskpred IDs
    delay_task_id = allocate_keys(conn, "task_task_id", 1)[0]
    delay_pred_ids = allocate_keys(conn, "taskpred_task_pred_id", len(impacted_act_codes))

    log(
        f"  [DB] Inserting delay activity '{delay_id}' ({duration_days} days = {duration_hours} hrs) under WBS 'Delay Events'..."
    )

    # 3. Create TASK row
    delay_data = {col: None for col in task_cols}
    delay_data["TASK_ID"] = delay_task_id
    delay_data["PROJ_ID"] = proj_id
    delay_data["WBS_ID"] = delay_wbs_id
    delay_data["CLNDR_ID"] = baseline_clndr_id
    delay_data["PHYS_COMPLETE_PCT"] = 0.0
    delay_data["REV_FDBK_FLAG"] = "N"
    delay_data["EST_WT"] = 1.0
    delay_data["LOCK_PLAN_FLAG"] = "N"
    delay_data["AUTO_COMPUTE_ACT_FLAG"] = "N"
    delay_data["COMPLETE_PCT_TYPE"] = "CP_Drtn"
    delay_data["TASK_TYPE"] = "TT_Task"
    delay_data["DURATION_TYPE"] = "DT_FixedDUR2"
    delay_data["STATUS_CODE"] = "TK_NotStart"
    delay_data["TASK_CODE"] = delay_id
    delay_data["TASK_NAME"] = delay_name
    delay_data["TARGET_DRTN_HR_CNT"] = duration_hours
    delay_data["REMAIN_DRTN_HR_CNT"] = duration_hours
    delay_data["CSTR_TYPE"] = cstr_type
    delay_data["CSTR_DATE"] = start_dt.strftime("%Y-%m-%d 08:00:00")
    delay_data["TARGET_START_DATE"] = start_dt.strftime("%Y-%m-%d 08:00:00")
    delay_data["TARGET_END_DATE"] = finish_dt.strftime("%Y-%m-%d 17:00:00")
    delay_data["EARLY_START_DATE"] = start_dt.strftime("%Y-%m-%d 08:00:00")
    delay_data["EARLY_END_DATE"] = finish_dt.strftime("%Y-%m-%d 17:00:00")
    delay_data["GUID"] = generate_p6_guid()
    delay_data["PRIORITY_TYPE"] = "PT_Normal"
    delay_data["CREATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    delay_data["UPDATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    delay_data["CREATE_USER"] = "admin"
    delay_data["UPDATE_USER"] = "admin"
    delay_data["CONTROL_UPDATES_FLAG"] = "N"
    delay_data["SCP_PCT_COMPLETE"] = 0.0

    delay_data = {col: delay_data[col] for col in task_cols if col in delay_data}
    insert_row(cur, "TASK", delay_data)

    # 4. Create Predecessor Relationship(s)
    try:
        for succ_code, delay_pred_id in zip(impacted_act_codes, delay_pred_ids):
            # Find the TASK row in the target project
            cur.execute("SELECT * FROM TASK WHERE PROJ_ID = ? AND TASK_CODE = ?", (proj_id, succ_code))
            succ_row = cur.fetchone()
            if not succ_row:
                raise ValueError(f"Successor task code '{succ_code}' not found in target project ID {proj_id}")
            succ_data = dict(zip(task_cols, succ_row))

            log(
                f"  [DB] Adding predecessor relationship: {delay_id} -> {succ_code} [{rel_type}, Lag: {lag_days}d = {lag_hours}h]"
            )

            pred_data = {col: None for col in pred_cols}
            pred_data["TASK_PRED_ID"] = delay_pred_id
            pred_data["TASK_ID"] = succ_data["TASK_ID"]  # successor
            pred_data["PRED_TASK_ID"] = delay_task_id  # predecessor
            pred_data["PROJ_ID"] = proj_id
            pred_data["PRED_PROJ_ID"] = proj_id
            pred_data["PRED_TYPE"] = rel_type
            pred_data["LAG_HR_CNT"] = lag_hours
            pred_data["CREATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pred_data["UPDATE_DATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pred_data["CREATE_USER"] = "admin"
            pred_data["UPDATE_USER"] = "admin"

            pred_data = {col: pred_data[col] for col in pred_cols if col in pred_data}
            insert_row(cur, "TASKPRED", pred_data)
    except Exception as e:
        raise ValueError(f"[Activity Impacted] {e}")


# ─────────────────────────────────────────────────────────────────────────────
# CPM SCHEDULER MATHEMATICAL ENGINE
# ─────────────────────────────────────────────────────────────────────────────


def run_cpm_scheduling(conn, proj_id, resolved_calendars=None):
    cur = conn.cursor()

    # 1. Get project start date
    cur.execute("SELECT PLAN_START_DATE FROM PROJECT WHERE PROJ_ID = ?", (proj_id,))
    proj_start_str = cur.fetchone()[0]
    proj_start_dt = parse_p6_datetime(proj_start_str, default_hour=8)
    proj_start_dt = proj_start_dt.replace(hour=8, minute=0, second=0, microsecond=0)

    # Load resolved calendars
    if resolved_calendars is None:
        resolved_calendars = load_and_resolve_calendars(conn)

    # 2. Load all tasks in project
    cur.execute(
        "SELECT TASK_ID, TASK_CODE, TARGET_DRTN_HR_CNT, CSTR_TYPE, CSTR_DATE, CLNDR_ID FROM TASK WHERE PROJ_ID = ?",
        (proj_id,),
    )
    tasks = {}
    for r in cur.fetchall():
        tid, code, dur_hours, cstr_type, cstr_date_str, clndr_id = r
        dur_days = int(dur_hours / 8.0) if dur_hours else 0

        cstr_date = None
        if cstr_date_str:
            cstr_date = parse_p6_datetime(cstr_date_str, default_hour=8)

        tasks[tid] = {
            "id": tid,
            "code": code,
            "dur_days": dur_days,
            "cstr_type": cstr_type,
            "cstr_date": cstr_date,
            "clndr_id": clndr_id,
            "early_start": None,
            "early_finish": None,
        }

    # 3. Load all predecessors (relationships)
    cur.execute("SELECT TASK_ID, PRED_TASK_ID, PRED_TYPE, LAG_HR_CNT FROM TASKPRED WHERE PROJ_ID = ?", (proj_id,))
    preds = {tid: [] for tid in tasks}
    successors = {tid: [] for tid in tasks}
    for r in cur.fetchall():
        succ_id, pred_id, pred_type, lag_hours = r
        if succ_id not in tasks or pred_id not in tasks:
            continue
        lag_days = int(lag_hours / 8.0) if lag_hours else 0
        preds[succ_id].append((pred_id, pred_type, lag_days))
        successors[pred_id].append(succ_id)

    # 4. Topological sort (Kahn's Algorithm)
    in_degree = {tid: len(preds[tid]) for tid in tasks}
    queue = [tid for tid in tasks if in_degree[tid] == 0]
    topo_order = []

    while queue:
        u = queue.pop(0)
        topo_order.append(u)
        for v in successors[u]:
            in_degree[v] -= 1
            if in_degree[v] == 0:
                queue.append(v)

    if len(topo_order) != len(tasks):
        log(f"  [CPM] Circular dependencies loop warning. Topo sort processed {len(topo_order)} of {len(tasks)} tasks.")
        remaining = [tid for tid in tasks if tid not in topo_order]
        topo_order.extend(remaining)

    # 5. Forward Pass to calculate Early Dates
    for tid in topo_order:
        t = tasks[tid]
        est = proj_start_dt
        while not is_workday(est, t["clndr_id"], resolved_calendars):
            est += timedelta(days=1)

        for pred_id, pred_type, lag_days in preds[tid]:
            p = tasks[pred_id]
            if p["early_finish"] is None:
                continue
            succ_clndr = t["clndr_id"]

            # Lag is in calendar days (7-day calendar)
            if pred_type == "PR_FS":
                # If predecessor finishes at or before 8:00 AM (milestone/0 dur), successor starts same day
                if p["early_finish"].hour <= 8:
                    candidate_start = p["early_finish"] + timedelta(days=lag_days)
                else:
                    candidate_start = p["early_finish"] + timedelta(days=1 + lag_days)
                while not is_workday(candidate_start, succ_clndr, resolved_calendars):
                    candidate_start += timedelta(days=1)
                if candidate_start > est:
                    est = candidate_start
            elif pred_type == "PR_SS":
                candidate_start = p["early_start"] + timedelta(days=lag_days)
                while not is_workday(candidate_start, succ_clndr, resolved_calendars):
                    candidate_start += timedelta(days=1)
                if candidate_start > est:
                    est = candidate_start
            elif pred_type == "PR_FF":
                candidate_finish = p["early_finish"] + timedelta(days=lag_days)
                candidate_start = add_workdays_cal(
                    candidate_finish, -(t["dur_days"] - 1), succ_clndr, resolved_calendars
                )
                candidate_start = candidate_start.replace(hour=8, minute=0, second=0)
                while not is_workday(candidate_start, succ_clndr, resolved_calendars):
                    candidate_start += timedelta(days=1)
                if candidate_start > est:
                    est = candidate_start
            elif pred_type == "PR_SF":
                candidate_finish = p["early_start"] + timedelta(days=lag_days)
                candidate_start = add_workdays_cal(
                    candidate_finish, -(t["dur_days"] - 1), succ_clndr, resolved_calendars
                )
                candidate_start = candidate_start.replace(hour=8, minute=0, second=0)
                while not is_workday(candidate_start, succ_clndr, resolved_calendars):
                    candidate_start += timedelta(days=1)
                if candidate_start > est:
                    est = candidate_start

        # Apply constraint
        if t["cstr_type"] in ("CS_SNET", "CS_MSOA") and t["cstr_date"]:
            if t["cstr_date"] > est:
                est = t["cstr_date"]
        elif t["cstr_type"] == "CS_MSO" and t["cstr_date"]:
            est = t["cstr_date"]

        while not is_workday(est, t["clndr_id"], resolved_calendars):
            est += timedelta(days=1)

        # Calculate early finish
        if t["dur_days"] > 0:
            ef = add_workdays_cal(est, t["dur_days"] - 1, t["clndr_id"], resolved_calendars).replace(
                hour=17, minute=0, second=0
            )
        else:
            ef = est.replace(hour=8, minute=0, second=0)

        # Apply finish constraints (which might push early start/finish further out)
        if t["cstr_type"] in ("CS_FNET", "CS_MFOA") and t["cstr_date"]:
            if t["cstr_date"] > ef:
                ef = t["cstr_date"]
                if t["dur_days"] > 0:
                    est = add_workdays_cal(ef, -(t["dur_days"] - 1), t["clndr_id"], resolved_calendars).replace(
                        hour=8, minute=0, second=0
                    )
                else:
                    est = ef.replace(hour=8, minute=0, second=0)
        elif t["cstr_type"] in ("CS_MFO", "CS_MFOB") and t["cstr_date"]:
            ef = t["cstr_date"]
            if t["dur_days"] > 0:
                est = add_workdays_cal(ef, -(t["dur_days"] - 1), t["clndr_id"], resolved_calendars).replace(
                    hour=8, minute=0, second=0
                )
            else:
                est = ef.replace(hour=8, minute=0, second=0)

        t["early_start"] = est
        t["early_finish"] = ef

    # 6. Update early dates back to database
    max_finish = proj_start_dt
    for tid, t in tasks.items():
        es_str = t["early_start"].strftime("%Y-%m-%d %H:%M:%S")
        ef_str = t["early_finish"].strftime("%Y-%m-%d %H:%M:%S")

        if t["early_finish"] > max_finish:
            max_finish = t["early_finish"]

        cur.execute(
            """
            UPDATE TASK SET 
                EARLY_START_DATE = ?, EARLY_END_DATE = ?, 
                LATE_START_DATE = ?, LATE_END_DATE = ?, 
                TARGET_START_DATE = ?, TARGET_END_DATE = ?,
                RESTART_DATE = ?, REEND_DATE = ?,
                REM_LATE_START_DATE = ?, REM_LATE_END_DATE = ?
            WHERE TASK_ID = ?""",
            (es_str, ef_str, es_str, ef_str, es_str, ef_str, es_str, ef_str, es_str, ef_str, tid),
        )

    # Update PROJECT scheduled end date and last schedule date in database
    max_finish_str = max_finish.strftime("%Y-%m-%d %H:%M:%S")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        UPDATE PROJECT SET 
            SCD_END_DATE = ?, 
            LAST_SCHEDULE_DATE = ? 
        WHERE PROJ_ID = ?""",
        (max_finish_str, now_str, proj_id),
    )

    # Also update SUMTASK to write Actual Start and Actual Finish so they display in P6 Projects View
    cur.execute("SELECT WBS_ID FROM PROJWBS WHERE PROJ_ID = ? AND PROJ_NODE_FLAG = 'Y'", (proj_id,))
    wbs_row = cur.fetchone()
    if wbs_row:
        root_wbs_id = wbs_row[0]
        try:
            cur.execute("SELECT COUNT(*) FROM SUMTASK WHERE PROJ_ID = ? AND WBS_ID = ?", (proj_id, root_wbs_id))
            exists = cur.fetchone()[0] > 0

            start_str = proj_start_dt.strftime("%Y-%m-%d %H:%M:%S")

            if exists:
                cur.execute(
                    """
                    UPDATE SUMTASK SET 
                        ACT_START_DATE = ?, 
                        ACT_END_DATE = ? 
                    WHERE PROJ_ID = ? AND WBS_ID = ?""",
                    (start_str, max_finish_str, proj_id, root_wbs_id),
                )
            else:
                cur.execute("PRAGMA table_info(SUMTASK)")
                cols = [col_info[1] for col_info in cur.fetchall()]
                if cols:
                    insert_data = {
                        "WBS_ID": root_wbs_id,
                        "PROJ_ID": proj_id,
                        "ACT_START_DATE": start_str,
                        "ACT_END_DATE": max_finish_str,
                        "CREATE_USER": "admin",
                        "UPDATE_USER": "admin",
                    }
                    insert_data = {k: v for k, v in insert_data.items() if k in cols}
                    placeholders = ", ".join(["?"] * len(insert_data))
                    fields = ", ".join(insert_data.keys())
                    sql = f"INSERT INTO SUMTASK ({fields}) VALUES ({placeholders})"
                    cur.execute(sql, list(insert_data.values()))
        except sqlite3.Error:
            pass

    return max_finish


# ─────────────────────────────────────────────────────────────────────────────
# RESULTS RETRIEVAL & EXCEL PARSING HELPERS
# ─────────────────────────────────────────────────────────────────────────────


def get_project_finish_date_from_db(db_path: str, project_id: int) -> str:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    chosen_date = None
    try:
        cur.execute("SELECT SCD_END_DATE FROM PROJECT WHERE PROJ_ID = ?", (project_id,))
        row = cur.fetchone()
        db_date_str = row[0] if row else None

        cur.execute("SELECT MAX(EARLY_END_DATE) FROM TASK WHERE PROJ_ID = ?", (project_id,))
        task_max_row = cur.fetchone()
        task_date_str = task_max_row[0] if task_max_row else None

        chosen_date = db_date_str or task_date_str
        if not chosen_date:
            return "N/A"

        dt = parse_p6_datetime(chosen_date)
        if dt:
            return dt.strftime("%d-%b-%y")
        return str(chosen_date)
    except Exception:
        return str(chosen_date) if chosen_date else "N/A"
    finally:
        conn.close()


DELAY_COLUMN_ALIASES = {
    "Delay Event ID": ["Delay Event ID", "Delay ID", "Event ID", "Activity ID"],
    "Description of Event": [
        "Description of Event",
        "Delay Event Name",
        "Event Name",
        "Description",
        "Event Description",
        "Name",
    ],
    "Notification Date": [
        "Notification Date",
        "Notification-Start Date",
        "Primary Constraint Date",
        "Start Date",
    ],
    "Delay in days": ["Delay in days", "Original Duration", "Duration"],
    "Activity Impacted": [
        "Activity Impacted",
        "Successor ID",
        "Impacting Activity ID",
        "Successor Activity ID",
    ],
    "Relationship": ["Relationship", "Relationship Type", "Predecessor Relationship"],
    "Lag": ["Lag", "Lag days", "Lag in days"],
    "Primary Constraint": ["Primary Constraint", "Constraint Type", "Constraint"],
}


def _parse_single_delay_sheet(df_raw, _log):
    """Attempts to detect headers and resolve required columns on a single DataFrame."""
    df_delays = df_raw.copy()
    df_delays.columns = [str(c).strip() for c in df_delays.columns]

    candidates = [(-1, [str(c).strip() for c in df_delays.columns])]
    for idx, row in df_delays.head(15).iterrows():
        row_vals = [str(val).strip() for val in row.values if pd.notna(val)]
        candidates.append((idx, row_vals))

    best_idx = -1
    best_score = -1
    for idx, cols_list in candidates:
        score = 0
        for key, alias_list in DELAY_COLUMN_ALIASES.items():
            if key in ["Lag", "Primary Constraint"]:
                continue
            matched = False
            for alias in alias_list:
                for col in cols_list:
                    if str(col).strip().lower() == alias.lower():
                        matched = True
                        break
                if matched:
                    break
            if matched:
                score += 1
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_idx != -1:
        new_columns = [
            str(c).strip() if pd.notna(c) and str(c).strip() != "" else f"Unnamed_{i}"
            for i, c in enumerate(df_delays.iloc[best_idx].values)
        ]
        df_delays.columns = new_columns
        df_delays = df_delays.iloc[best_idx + 1 :].reset_index(drop=True)

    resolved_cols = {}
    missing_required = []

    for key, alias_list in DELAY_COLUMN_ALIASES.items():
        matched_col = None
        for alias in alias_list:
            for col in df_delays.columns:
                if str(col).strip().lower() == alias.lower():
                    matched_col = col
                    break
            if matched_col:
                break
        if not matched_col:
            for alias in alias_list:
                for col in df_delays.columns:
                    col_clean = str(col).strip().lower()
                    if alias.lower() in col_clean or col_clean in alias.lower():
                        matched_col = col
                        break
                if matched_col:
                    break

        if key not in ["Lag", "Primary Constraint"]:
            if not matched_col:
                missing_required.append(key)
        resolved_cols[key] = matched_col

    if not missing_required and resolved_cols.get("Delay Event ID"):
        df_delays = df_delays.dropna(subset=[resolved_cols["Delay Event ID"]]).copy()

    return df_delays, resolved_cols, best_idx, missing_required


def resolve_delay_dataframe(excel_path, logger=None, error_logger=None):
    """
    Reads the delay events Excel file, auto-detects the header row (falling back
    to subsequent sheets if the first sheet is an empty template), maps flexible
    column aliases in priority order, and returns (df_delays, resolved_cols, best_idx).
    """
    _log = logger or log
    _err = error_logger or err

    _log(f"Reading delay events: {excel_path}")
    first_result = None

    with pd.ExcelFile(excel_path) as xl:
        sheet_frames = [(name, xl.parse(name)) for name in xl.sheet_names]

    for sheet_idx, (sheet_name, df_raw) in enumerate(sheet_frames):
        df_delays, resolved_cols, best_idx, missing_required = _parse_single_delay_sheet(df_raw, _log)
        if sheet_idx == 0:
            first_result = (df_delays, resolved_cols, best_idx, missing_required)
        if not missing_required and len(df_delays) > 0:
            if sheet_idx > 0:
                _log(f"  [Excel] First sheet was empty; loaded sheet '{sheet_name}' ({len(df_delays)} row(s)).")
            if best_idx != -1:
                _log(f"  [Excel] Detected header row at index {best_idx + 2} in Excel sheet.")
            return df_delays, resolved_cols, best_idx

    df_delays, resolved_cols, best_idx, missing_required = first_result
    if best_idx != -1:
        _log(f"  [Excel] Detected header row at index {best_idx + 2} in Excel sheet.")
    if missing_required:
        _err(f"Excel is missing required columns or recognizable aliases: {missing_required}")
        _err(f"Available columns: {list(df_delays.columns)}")
        raise ValueError(f"Excel is missing required columns: {missing_required}")

    return df_delays, resolved_cols, best_idx


# ─────────────────────────────────────────────────────────────────────────────
# CORE PROCESSING (extracted from main() so GUI and CLI can both call it)
# ─────────────────────────────────────────────────────────────────────────────


def process_delay_events(
    excel_path,
    db_path,
    source_project,
    default_remarks="",
    logger=None,
    error_logger=None,
    output_log_excel=None,
):
    """
    Reads the delay-events Excel file, replicates the baseline project for
    each delay event, inserts the delay activity, runs CPM scheduling, and
    returns (results, skipped_rows).

    Note: Separate completion-log Excel file creation has been standardized into
    the single consolidated Milestone vs Delay Excel output generated by
    run_milestone_vs_delay(). If output_log_excel is explicitly passed by an
    external caller, it will still write df_out there for backward compatibility.
    """
    _log = logger or log
    _err = error_logger or err

    df_delays, resolved_cols, best_idx = resolve_delay_dataframe(
        excel_path, logger=_log, error_logger=_err
    )

    # 3. Modify SQLite database and run calculations
    _log(f"\nProcessing SQLite database at {db_path}...")

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(
        "SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
        (source_project,),
    )
    row = cur.fetchone()
    if not row:
        conn.close()
        _err(f"Baseline project '{source_project}' not found in database.")
        raise ValueError(f"Baseline project '{source_project}' not found in database.")
    source_proj_id = row[0]
    resolved_calendars = load_and_resolve_calendars(conn)
    conn.close()

    results = []
    skipped_rows = []

    for idx, row in df_delays.iterrows():
        delay_id = str(row[resolved_cols["Delay Event ID"]]).strip()
        delay_name = str(row[resolved_cols["Description of Event"]]).strip()

        copy_id = delay_id
        excel_row_num = idx + (best_idx + 3 if best_idx != -1 else 2)

        print("\n" + "-" * 72)
        _log(f"Replicating baseline project: {source_project} -> {copy_id}")
        _log(f"Project Copy Code: {copy_id}")
        print("-" * 72)

        conn = None
        try:
            conn = sqlite3.connect(db_path)
            new_proj_id, delay_wbs_id = duplicate_project_in_db(conn, source_project, copy_id)
            conn.commit()

            _log(f"  [DB] Inserting delay activity '{delay_id}' into '{copy_id}'...")
            add_delay_event_to_project(
                conn, new_proj_id, delay_wbs_id, row, resolved_cols, source_proj_id, resolved_calendars
            )
            conn.commit()

            _log(f"  [CPM] Running critical path calculation on {copy_id}...")
            run_cpm_scheduling(conn, new_proj_id, resolved_calendars)
            conn.commit()

            finish_date = get_project_finish_date_from_db(db_path, new_proj_id)
            _log(f"  Calculated project finish date: {finish_date}")

            lag_col = resolved_cols.get("Lag")
            lag_val = (
                float(row[lag_col])
                if (lag_col and lag_col in row and pd.notna(row[lag_col]) and str(row[lag_col]).strip() != "")
                else 0.0
            )
            results.append(
                {
                    "Delay Event ID": delay_id,
                    "Delay Event Name": delay_name,
                    "Impacting Activity": str(row[resolved_cols["Activity Impacted"]]).strip(),
                    "Relationship Type": str(row[resolved_cols["Relationship"]]).strip(),
                    "Lag (days)": lag_val,
                    "New Finish Date": finish_date,
                    "Observation": default_remarks if default_remarks else "",
                }
            )
        except Exception as exc:
            _err(f"Discrepancy / error at Excel row {excel_row_num} (Delay ID: {delay_id}): {exc}")
            skipped_rows.append(
                {
                    "excel_row": excel_row_num,
                    "delay_id": delay_id,
                    "delay_name": delay_name,
                    "error": str(exc),
                }
            )
            if conn:
                try:
                    conn.rollback()
                except Exception:
                    pass
                try:
                    cur = conn.cursor()
                    cur.execute("SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ?", (copy_id,))
                    exists_row = cur.fetchone()
                    if exists_row:
                        proj_id_to_del = exists_row[0]
                        cleanup_project_tables(cur, proj_id_to_del)
                        conn.commit()
                        _log(f"  [DB] Successfully cleaned up project copy '{copy_id}' from database due to error.")
                except Exception as clean_exc:
                    _err(f"Failed to clean up project copy '{copy_id}' after error: {clean_exc}")
        finally:
            if conn:
                conn.close()

    _log("\nDatabase processing complete.")

    # Optional backward-compatibility hook if an explicit output_log_excel path was passed programmatically
    if results and output_log_excel:
        df_out = pd.DataFrame(results)
        df_out.to_excel(output_log_excel, index=False)
        _log(f"\nResults written to -> {output_log_excel}")
    elif not results:
        _err("\nNo results gathered.")

    # Output summary of skipped rows
    print("\n" + "=" * 72)
    print("  SUMMARY OF SKIPPED ROWS WITH DISCREPANCIES")
    print("=" * 72)
    if skipped_rows:
        _log(f"Total rows skipped: {len(skipped_rows)}")
        for skip in skipped_rows:
            _log(f"  - Excel Row {skip['excel_row']} (Delay ID: {skip['delay_id']}): {skip['error']}")
    else:
        _log("No rows were skipped. All rows processed successfully.")
    print("=" * 72 + "\n")

    return results, skipped_rows


# ─────────────────────────────────────────────────────────────────────────────
# SPECIFIED ACTIVITIES IMPORTER & CONSOLIDATED REPORT GENERATOR
# ─────────────────────────────────────────────────────────────────────────────


class SpecifiedActivityList:
    """
    Represents an ordered list of user-specified activities imported from a
    single Excel file to be reported in the consolidated EOT output.
    Also provides .search() for regex duck-typing compatibility.
    """

    def __init__(self, items, source_path=""):
        # items: list of dicts {"task_code": str, "milestone_id": str|None, "description": str|None}
        self.items = items
        self.source_path = source_path
        self._upper_set = {str(it["task_code"]).strip().upper() for it in items if it.get("task_code")}

    def search(self, code):
        if str(code).strip().upper() in self._upper_set:
            return re.match(r"(.*)", str(code))
        return None


_ACTIVITY_CODE_HEADERS = {
    "task code",
    "task_code",
    "activity id",
    "activity_id",
    "activity code",
    "activity_code",
    "activity",
    "task id",
    "task_id",
    "milestone code",
    "milestone_code",
    "code",
    "activity impacted",
    "impacting activity",
    "successor id",
    "successor activity id",
}

_MILESTONE_ID_HEADERS = {
    "milestone id",
    "milestone_id",
    "milestone",
    "ms id",
    "ms_id",
    "id",
    "sn",
    "s.no",
    "s.no.",
    "sl. no.",
    "sl.no",
    "sr. no.",
    "sr.no",
}

_ACTIVITY_DESC_HEADERS = {
    "milestone description",
    "activity name",
    "task name",
    "description",
    "activity description",
    "task description",
    "description of event",
    "delay event name",
    "event name",
    "name",
    "title",
    "milestone name",
}


def load_ms_pattern_from_excel(path):
    """
    Imports an Excel file specifying which activities (or milestone pattern)
    should be included in the consolidated EOT output report.

    Supports:
    1. A structured Excel list/table of specified activities (with headers such as
       'Activity ID', 'Task Code', 'Milestone ID', 'Milestone Description', etc.,
       including tables offset by title/banner rows).
    2. A headerless single-column or multi-column Excel list of activity codes.
    3. Legacy 2-cell regex configuration ('Pattern' / 'Milestone Pattern' | '<regex>').

    Returns either a SpecifiedActivityList or a compiled re.Pattern (or None if invalid/empty).
    """
    if not path or not os.path.exists(path):
        return None
    try:
        with pd.ExcelFile(path) as xl:
            sheet_frames = [(name, xl.parse(name, header=None)) for name in xl.sheet_names]

        for sheet_name, df in sheet_frames:
            if df is None or df.empty:
                continue

            # 1. Check if this sheet is a legacy 2-cell regex pattern configuration
            for r_idx in range(min(15, len(df))):
                for c_idx in range(df.shape[1]):
                    cell_val = df.iloc[r_idx, c_idx]
                    if pd.isna(cell_val):
                        continue
                    val = str(cell_val).strip().lower()
                    if val in ("pattern", "milestone pattern"):
                        if c_idx + 1 < df.shape[1]:
                            raw_pat = df.iloc[r_idx, c_idx + 1]
                            if pd.notna(raw_pat):
                                pattern_val = str(raw_pat).strip()
                                if pattern_val and pattern_val.lower() != "nan":
                                    log(f"  [Milestone Config] Found custom pattern: '{pattern_val}'")
                                    try:
                                        return re.compile(pattern_val, re.IGNORECASE)
                                    except re.error as e:
                                        err(f"Invalid regex pattern in Excel: {pattern_val}. Error: {e}")
                                        return None

            # 2. Check for a structured table of specified activities with column headers
            header_row_idx = None
            code_col = None
            ms_id_col = None
            desc_col = None

            for r_idx in range(min(15, len(df))):
                row_code_col = None
                row_ms_col = None
                row_desc_col = None
                for c_idx in range(df.shape[1]):
                    raw_val = df.iloc[r_idx, c_idx]
                    if pd.isna(raw_val):
                        continue
                    clean_h = str(raw_val).strip().lower()
                    if row_code_col is None and clean_h in _ACTIVITY_CODE_HEADERS:
                        row_code_col = c_idx
                    elif row_ms_col is None and clean_h in _MILESTONE_ID_HEADERS:
                        row_ms_col = c_idx
                    elif row_desc_col is None and clean_h in _ACTIVITY_DESC_HEADERS:
                        row_desc_col = c_idx

                if row_code_col is not None or (row_ms_col is not None and row_desc_col is not None):
                    header_row_idx = r_idx
                    code_col = row_code_col if row_code_col is not None else row_ms_col
                    ms_id_col = row_ms_col if row_ms_col != code_col else None
                    desc_col = row_desc_col
                    break

            items = []
            if header_row_idx is not None and code_col is not None:
                for r_idx in range(header_row_idx + 1, len(df)):
                    raw_code = df.iloc[r_idx, code_col]
                    if pd.isna(raw_code):
                        continue
                    code_str = str(raw_code).strip()
                    if not code_str or code_str.lower() == "nan":
                        continue

                    ms_id_str = None
                    if ms_id_col is not None and pd.notna(df.iloc[r_idx, ms_id_col]):
                        val_ms = str(df.iloc[r_idx, ms_id_col]).strip()
                        if val_ms and val_ms.lower() != "nan":
                            ms_id_str = val_ms

                    desc_str = None
                    if desc_col is not None and pd.notna(df.iloc[r_idx, desc_col]):
                        val_desc = str(df.iloc[r_idx, desc_col]).strip()
                        if val_desc and val_desc.lower() != "nan":
                            desc_str = val_desc

                    for part in re.split(r"[\n\r,;]+", code_str):
                        part = part.strip()
                        if part:
                            items.append(
                                {
                                    "task_code": part,
                                    "milestone_id": ms_id_str,
                                    "description": desc_str,
                                }
                            )
            else:
                # 3. Headerless activity list fallback (e.g. single column of Activity IDs starting in Row 1)
                non_empty_cols = [
                    c for c in range(df.shape[1]) if df.iloc[:, c].dropna().astype(str).str.strip().ne("").any()
                ]
                if non_empty_cols:
                    first_col = non_empty_cols[0]
                    # Check if first column is just serial numbers (1, 2, 3...) and second column has codes
                    col0_vals = [
                        str(v).strip() for v in df.iloc[:, first_col].dropna().values if str(v).strip() != ""
                    ]
                    is_col0_serial = bool(col0_vals) and all(
                        re.match(r"^\d+(\.0+)?$", v) for v in col0_vals
                    )
                    if is_col0_serial and len(non_empty_cols) >= 2:
                        code_col = non_empty_cols[1]
                        ms_id_col = first_col
                        desc_col = non_empty_cols[2] if len(non_empty_cols) >= 3 else None
                    else:
                        code_col = first_col
                        ms_id_col = None
                        desc_col = non_empty_cols[1] if len(non_empty_cols) >= 2 else None

                    for r_idx in range(len(df)):
                        raw_code = df.iloc[r_idx, code_col]
                        if pd.isna(raw_code):
                            continue
                        code_str = str(raw_code).strip()
                        if not code_str or code_str.lower() == "nan":
                            continue

                        ms_id_str = None
                        if ms_id_col is not None and pd.notna(df.iloc[r_idx, ms_id_col]):
                            val_ms = str(df.iloc[r_idx, ms_id_col]).strip()
                            if val_ms and val_ms.lower() != "nan":
                                ms_id_str = val_ms

                        desc_str = None
                        if desc_col is not None and pd.notna(df.iloc[r_idx, desc_col]):
                            val_desc = str(df.iloc[r_idx, desc_col]).strip()
                            if val_desc and val_desc.lower() != "nan":
                                desc_str = val_desc

                        for part in re.split(r"[\n\r,;]+", code_str):
                            part = part.strip()
                            if part:
                                items.append(
                                    {
                                        "task_code": part,
                                        "milestone_id": ms_id_str,
                                        "description": desc_str,
                                    }
                                )

            if items:
                log(
                    f"  [Activities Config] Loaded {len(items)} specified activity item(s) from '{os.path.basename(path)}' (sheet: '{sheet_name}')."
                )
                return SpecifiedActivityList(items, source_path=path)

    except Exception as e:
        err(f"Failed to read activities/milestone config from Excel: {e}")
    return None


def format_date_ms(date_str):
    """Formats P6 database datetime strings into a cleaner Excel-friendly format."""
    if not date_str or str(date_str).strip() == "" or pd.isna(date_str):
        return "N/A"
    dt = parse_p6_datetime(date_str)
    if dt:
        return dt.strftime("%d-%b-%y")
    return str(date_str)


def get_unique_output_path(base_dir, base_name="after_impact_MS_completion", ext=".xlsx"):
    """Generates an incremental filename using current computer date and R0, R1, etc. suffix."""
    date_str = datetime.now().strftime("%y%m%d")
    counter = 0
    while True:
        candidate = os.path.join(base_dir, f"{date_str}_{base_name}_R{counter}{ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def prompt_save_path(default_dir, default_filename, title):
    """Opens a native file dialogue popup asking where to save the Excel output.
    Falls back to the default path in batch/headless environments where tkinter is unavailable."""
    default_full_path = os.path.join(default_dir, default_filename)
    if os.environ.get("EOT_BATCH_MODE") == "1":
        return default_full_path
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        save_path = filedialog.asksaveasfilename(
            initialdir=default_dir,
            initialfile=default_filename,
            title=title,
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx",
            parent=root,
        )
        root.destroy()
        return save_path
    except Exception:
        return default_full_path


def extract_delay_ids_from_excel(excel_path):
    """Reads the Excel delay sheet and returns a list of unique delay project codes."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found at: {excel_path}")

    aliases = DELAY_COLUMN_ALIASES["Delay Event ID"]
    first_error = None

    with pd.ExcelFile(excel_path) as xl:
        sheet_frames = [(name, xl.parse(name)) for name in xl.sheet_names]

    for sheet_name, df_delays in sheet_frames:
        df_delays.columns = [str(c).strip() for c in df_delays.columns]

        candidates = [(-1, [str(c).strip() for c in df_delays.columns])]
        for idx, row in df_delays.head(15).iterrows():
            row_vals = [str(val).strip() for val in row.values if pd.notna(val)]
            candidates.append((idx, row_vals))

        best_idx = -1
        best_score = -1
        for idx, cols_list in candidates:
            score = 0
            for col in cols_list:
                col_clean = str(col).strip().lower()
                for alias in aliases:
                    if col_clean == alias.lower():
                        score += 1
                        break
            if score > best_score:
                best_score = score
                best_idx = idx

        if best_idx != -1:
            new_columns = [
                str(c).strip() if pd.notna(c) and str(c).strip() != "" else f"Unnamed_{i}"
                for i, c in enumerate(df_delays.iloc[best_idx].values)
            ]
            df_delays.columns = new_columns
            df_delays = df_delays.iloc[best_idx + 1 :].reset_index(drop=True)

        id_col = None
        for alias in aliases:
            for col in df_delays.columns:
                if str(col).strip().lower() == alias.lower():
                    id_col = col
                    break
            if id_col:
                break

        if not id_col:
            for alias in aliases:
                for col in df_delays.columns:
                    col_clean = str(col).strip().lower()
                    if alias.lower() in col_clean or col_clean in alias.lower():
                        id_col = col
                        break
                if id_col:
                    break

        if not id_col:
            if first_error is None:
                first_error = ValueError("Excel file is missing a recognizable 'Delay Event ID' column.")
            continue

        df_delays = df_delays.dropna(subset=[id_col]).copy()
        ids = []
        seen = set()
        for val in df_delays[id_col].values:
            s = str(val).strip()
            if s and s not in seen:
                seen.add(s)
                ids.append(s)
        if ids:
            return ids

    if first_error:
        raise first_error
    return []


def retrieve_project_milestones(cur, baseline_proj_id, ms_pattern=None):
    """
    Retrieves the target activities/milestones from the baseline project:
    - If ms_pattern is a SpecifiedActivityList (imported from an Excel file of specified
      activities), resolves every listed activity in the exact order specified.
    - Otherwise, applies the 3-level milestone fallback (Level 1: Regex pattern ->
      Level 2: Milestone task types -> Level 3: All tasks).
    """
    cur.execute(
        """
        SELECT TASK_ID, TASK_CODE, TASK_NAME, EARLY_END_DATE, TASK_TYPE
        FROM TASK
        WHERE PROJ_ID = ?
    """,
        (baseline_proj_id,),
    )
    baseline_tasks = cur.fetchall()

    # Mode A: User-imported Excel list of specified activities
    if isinstance(ms_pattern, SpecifiedActivityList) or hasattr(ms_pattern, "items"):
        by_code_exact = {}
        by_code_upper = {}
        by_name_upper = {}
        by_ms_norm = {}

        for tid, code, name, end_date, ttype in baseline_tasks:
            c_str = str(code).strip()
            entry = (tid, c_str, name, end_date, ttype)
            by_code_exact[c_str] = entry
            by_code_upper[c_str.upper()] = entry
            if name and str(name).strip():
                by_name_upper[str(name).strip().upper()] = entry
            m = MS_PATTERN.search(c_str)
            if m and m.group(1) and m.group(1).isdigit():
                norm_key = f"MS-{int(m.group(1)):02d}"
                by_ms_norm.setdefault(norm_key.upper(), entry)

        milestones = []
        seen_codes = set()
        for idx, item in enumerate(ms_pattern.items):
            raw_code = str(item.get("task_code") or "").strip()
            if not raw_code:
                continue

            matched_entry = (
                by_code_exact.get(raw_code)
                or by_code_upper.get(raw_code.upper())
                or by_ms_norm.get(raw_code.upper())
                or by_name_upper.get(raw_code.upper())
            )
            if not matched_entry:
                m_try = MS_PATTERN.search(raw_code)
                if m_try and m_try.group(1) and m_try.group(1).isdigit():
                    matched_entry = by_ms_norm.get(f"MS-{int(m_try.group(1)):02d}")

            if matched_entry:
                _, db_code, db_name, end_date, _ = matched_entry
                if db_code.upper() in seen_codes:
                    continue
                seen_codes.add(db_code.upper())

                if item.get("milestone_id"):
                    normalized_id = str(item["milestone_id"]).strip()
                else:
                    m = MS_PATTERN.search(db_code)
                    if m and len(m.groups()) >= 1 and m.group(1) and m.group(1).isdigit():
                        normalized_id = f"MS-{int(m.group(1)):02d}"
                    else:
                        normalized_id = db_code

                milestones.append(
                    {
                        "task_code": db_code,
                        "task_name": db_name if (db_name and str(db_name).strip()) else (item.get("description") or db_code),
                        "normalized_id": normalized_id,
                        "ms_num": idx,
                        "baseline_date": end_date,
                    }
                )
            else:
                if raw_code.upper() in seen_codes:
                    continue
                seen_codes.add(raw_code.upper())
                print(f"[WARNING] Specified activity '{raw_code}' not found in baseline project.")
                milestones.append(
                    {
                        "task_code": raw_code,
                        "task_name": item.get("description") or "Activity not found in baseline",
                        "normalized_id": item.get("milestone_id") or raw_code,
                        "ms_num": idx,
                        "baseline_date": None,
                    }
                )

        return milestones

    # Mode B: Regex pattern with 3-level fallback
    if ms_pattern is None:
        ms_pattern = MS_PATTERN

    milestones = []
    # Level 1: Find tasks matching custom/standard pattern
    for tid, code, name, end_date, ttype in baseline_tasks:
        match = ms_pattern.search(str(code))
        if match:
            if len(match.groups()) >= 1 and match.group(1) and match.group(1).isdigit():
                ms_num = int(match.group(1))
                normalized_id = f"MS-{ms_num:02d}"
            else:
                ms_num = 99999
                normalized_id = code
            milestones.append(
                {
                    "task_code": code,
                    "task_name": name,
                    "normalized_id": normalized_id,
                    "ms_num": ms_num,
                    "baseline_date": end_date,
                }
            )

    # Level 2 Fallback: If no MS-XX pattern tasks, look for other milestone task types
    if not milestones:
        milestone_types = ("TT_MileStart", "TT_MileEnd", "TT_FinMile", "TT_StartMile")
        for tid, code, name, end_date, ttype in baseline_tasks:
            if ttype in milestone_types or "milestone" in str(name).lower() or "milestone" in str(code).lower():
                milestones.append(
                    {
                        "task_code": code,
                        "task_name": name,
                        "normalized_id": code,
                        "ms_num": 99999,
                        "baseline_date": end_date,
                    }
                )
        if milestones:
            print("[INFO] No tasks matching pattern 'MS-XX' found. Found other milestone tasks in baseline project.")

    # Level 3 Fallback: If still no milestones at all, use all tasks in the project
    if not milestones:
        print("[INFO] No milestones found in baseline project. Generating report with all project tasks instead.")
        for tid, code, name, end_date, ttype in baseline_tasks:
            milestones.append(
                {
                    "task_code": code,
                    "task_name": name,
                    "normalized_id": code,
                    "ms_num": 99999,
                    "baseline_date": end_date,
                }
            )

    milestones.sort(key=lambda x: (x["ms_num"], str(x["task_code"])))
    return milestones


def build_milestone_workbook(db_path, source_project, milestones, delay_dates, title_text):
    """Builds and returns the styled openpyxl Workbook in the Combined EOT output format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Milestone vs Delay"
    ws.views.sheetView[0].showGridLines = True

    total_cols = 4 + len(delay_dates)
    last_col_letter = get_column_letter(total_cols)

    # Title Banner (Rows 1 & 2 merged)
    ws.merge_cells(f"A1:{last_col_letter}2")
    title_cell = ws["A1"]
    title_cell.value = title_text
    title_cell.font = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for r in range(1, 3):
        ws.row_dimensions[r].height = 20
        for c in range(1, total_cols + 1):
            ws.cell(row=r, column=c).fill = title_fill

    # Metadata (Row 3)
    ws.merge_cells(f"A3:{last_col_letter}3")
    meta_cell = ws["A3"]
    meta_cell.value = (
        f"Database: {db_path}   |   Baseline Project: {source_project}   |   "
        f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    meta_cell.font = Font(name="Segoe UI", size=9, italic=True, color="555555")
    meta_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 20

    # Table Headers (Row 5)
    headers = ["Milestone ID", "Task Code", "Milestone Description", f"Baseline ({source_project})"]
    active_delays = list(delay_dates.keys())
    headers.extend(active_delays)

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[5].height = 26
    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(row=5, column=col_idx)
        cell.value = text
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_left if col_idx == 3 else align_center

    thin_side = Side(style="thin", color="D9D9D9")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    baseline_fill = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type="solid")
    regular_font = Font(name="Segoe UI", size=10, color="000000")
    bold_font = Font(name="Segoe UI", size=10, bold=True, color="000000")

    current_row = 6
    for idx, ms in enumerate(milestones):
        ws.row_dimensions[current_row].height = 20
        use_zebra = idx % 2 == 1

        c1 = ws.cell(row=current_row, column=1, value=ms["normalized_id"])
        c1.alignment = align_center

        c2 = ws.cell(row=current_row, column=2, value=ms["task_code"])
        c2.alignment = align_center

        c3 = ws.cell(row=current_row, column=3, value=ms["task_name"])
        c3.alignment = align_left

        c4 = ws.cell(row=current_row, column=4, value=format_date_ms(ms["baseline_date"]))
        c4.alignment = align_center
        c4.font = bold_font
        c4.fill = baseline_fill

        for d_idx, d_id in enumerate(active_delays, 5):
            d_date_raw = delay_dates[d_id].get(ms["task_code"], None)
            val = format_date_ms(d_date_raw)
            cd = ws.cell(row=current_row, column=d_idx, value=val)
            cd.alignment = align_center
            if use_zebra:
                cd.fill = zebra_fill
            if val != "N/A" and val != format_date_ms(ms["baseline_date"]):
                cd.font = bold_font

        for c_idx in range(1, total_cols + 1):
            cell = ws.cell(row=current_row, column=c_idx)
            cell.border = border_all
            if c_idx != 4 and cell.font != bold_font:
                cell.font = regular_font
                if use_zebra and c_idx not in (3, 4):
                    cell.fill = zebra_fill
            if use_zebra and c_idx == 3:
                cell.fill = zebra_fill

        current_row += 1

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row < 5:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    return wb


def run_milestone_vs_delay(db_path, excel_path, source_project, ms_pattern=None, output_path=None):
    print("=" * 72)
    print("  PRIMAVERA P6 MILESTONE VS DELAY REPORT GENERATOR")
    print("=" * 72)

    db_path = db_path.strip("'\"")
    excel_path = excel_path.strip("'\"")

    if not os.path.exists(db_path):
        print(f"[ERROR] Primavera SQLite database not found at: {db_path}")
        return

    # 1. Read Delay IDs from Excel
    print(f"\nReading delay IDs from: {excel_path}")
    try:
        delay_ids = extract_delay_ids_from_excel(excel_path)
        print(f"Loaded {len(delay_ids)} delay event(s): {delay_ids}")
    except Exception as e:
        print(f"[ERROR] Failed to read delay events from Excel: {e}")
        return

    # 2. Connect to SQLite and fetch project IDs
    print(f"Connecting to database: {db_path}")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    try:
        cur.execute(
            "SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
            (source_project,),
        )
        row = cur.fetchone()
        if not row:
            print(f"[ERROR] Baseline project '{source_project}' not found in database.")
            return
        baseline_proj_id = row[0]
        print(f"Found baseline project '{source_project}' (ID: {baseline_proj_id})")

        # 3. Retrieve specified activities or milestones from baseline project
        milestones = retrieve_project_milestones(cur, baseline_proj_id, ms_pattern)
        print(f"Found {len(milestones)} item(s) to report on from baseline project.")

        # 4. Extract dates for each delay project copy
        delay_dates = {}
        for d_id in delay_ids:
            cur.execute(
                "SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
                (d_id,),
            )
            proj_row = cur.fetchone()
            if not proj_row:
                print(f"[WARNING] Project copy for delay ID '{d_id}' not found in database. Skipping...")
                continue
            d_proj_id = proj_row[0]

            cur.execute("SELECT TASK_CODE, EARLY_END_DATE FROM TASK WHERE PROJ_ID = ?", (d_proj_id,))
            task_dates = dict(cur.fetchall())
            delay_dates[d_id] = task_dates

    except Exception as e:
        print(f"[ERROR] Database query failed: {e}")
        return
    finally:
        conn.close()

    # 5. Build the single consolidated Excel report using openpyxl (Combined EOT structure)
    print("\nGenerating consolidated report...")
    wb = build_milestone_workbook(
        db_path,
        source_project,
        milestones,
        delay_dates,
        "P6 Milestone Finish Dates vs Delay Event Impacts",
    )

    if not output_path:
        excel_dir = os.path.dirname(os.path.abspath(excel_path))
        output_path = get_unique_output_path(excel_dir, base_name="after_impact_MS_completion")

    try:
        wb.save(output_path)
        print("\nSuccess! Consolidated Milestone vs Delay report generated at:")
        print(f"   -> {output_path}")
        print("=" * 72)
    except PermissionError:
        print(f"\n[ERROR] Permission denied when writing to {output_path}.")
        print("Please close the Excel file in Excel and run the script again.")
        sys.exit(1)


def resolve_milestone_pattern_interactive(args):
    """
    Resolves the specified activities list (or milestone regex pattern) from CLI flags
    or interactive user prompt. If --milestone-config / --activities-excel is provided,
    automatically loads the specified activities from that file.
    """
    milestone_mode = args.milestone_mode
    milestone_config = args.milestone_config

    if milestone_config and milestone_mode is None:
        milestone_mode = "custom"

    if milestone_mode == "standard":
        log("Using standard Level 1 milestone pattern (MS-XX).")
        return re.compile(r"MS[-_ ]?(\d+)", re.IGNORECASE)
    elif milestone_mode == "custom":
        if not milestone_config or not os.path.exists(milestone_config):
            err(f"Activities / milestone configuration file not found: {milestone_config}")
            sys.exit(1)
        log(f"Loading specified activities / milestone pattern from Excel: {milestone_config}")
        ms_pattern = load_ms_pattern_from_excel(milestone_config)
        if ms_pattern is None:
            err("No valid activities list or milestone pattern found in configuration file.")
            sys.exit(1)
        return ms_pattern

    # Interactive fallback when neither --milestone-mode nor --milestone-config is specified
    try:
        proj_name_to_check = args.source_project if args.source_project else "MAHSR T3-CWP-R6-1"
        print(f"Are you working on '{proj_name_to_check}' with standard MS-XX milestones? (yes/no): ", end="", flush=True)
        is_mahsr = None
        if os.environ.get("EOT_BATCH_MODE") != "1":
            try:
                import tkinter as tk
                from tkinter import messagebox

                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                is_yes = messagebox.askyesno(
                    "Project Verification",
                    f"Use standard MS-XX milestone pattern for '{proj_name_to_check}'?\n(Click 'No' to import an Excel file with a list of specified activities)",
                    parent=root,
                )
                root.destroy()
                is_mahsr = "yes" if is_yes else "no"
                print(is_mahsr)
            except Exception:
                is_mahsr = None

        if is_mahsr is None:
            is_mahsr = input().strip().lower() or "yes"

        if is_mahsr in ["yes", "y"]:
            log("Using standard Level 1 milestone pattern (MS-XX).")
            return re.compile(r"MS[-_ ]?(\d+)", re.IGNORECASE)
        else:
            ms_pattern = None
            while ms_pattern is None:
                print("Enter the Excel file path containing the list of specified activities: ", end="", flush=True)
                ms_config_path = None
                if os.environ.get("EOT_BATCH_MODE") != "1":
                    try:
                        import tkinter as tk
                        from tkinter import filedialog

                        root = tk.Tk()
                        root.withdraw()
                        root.attributes("-topmost", True)
                        ms_config_path = filedialog.askopenfilename(
                            title="Select Specified Activities Excel File",
                            filetypes=[("Excel files", "*.xlsx *.xls")],
                            parent=root,
                        )
                        root.destroy()
                    except Exception:
                        ms_config_path = None

                if ms_config_path is None:
                    ms_config_path = input().strip().strip("'\"")

                if not ms_config_path or not os.path.exists(ms_config_path):
                    print("None (Selection cancelled)")
                    err("No valid activities Excel file selected.")
                    sys.exit(1)

                print(ms_config_path)
                log(f"Loading specified activities / milestone pattern from Excel: {ms_config_path}")
                ms_pattern = load_ms_pattern_from_excel(ms_config_path)
                if ms_pattern is None:
                    print("[ERROR] Could not extract activities or pattern from the selected Excel file.")
            return ms_pattern
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(0)


def show_popup(success=True, skipped_rows=None):
    if os.environ.get("EOT_AUTO_NO_POPUP") == "1" or os.environ.get("EOT_BATCH_MODE") == "1":
        sys.exit(0 if success else 1)

    sys.stdout.flush()
    sys.stderr.flush()
    try:
        import time
        import tkinter as tk

        time.sleep(1.0)
        root = tk.Tk()
        root.title("Process Status")

        if skipped_rows:
            root.geometry("600x400")
        else:
            root.geometry("350x150")

        root.configure(bg="#F0F0F0")
        root.resizable(True, True)
        root.attributes("-topmost", True)

        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"+{x}+{y}")

        if success:
            if skipped_rows:
                msg = "The process completed with some discrepancies/skipped rows."
                color = "#E65100"
            else:
                msg = "The process is completed successfully."
                color = "#2e7d32"
        else:
            msg = "Error occurred, please check console."
            color = "#c62828"

        label = tk.Label(
            root,
            text=msg,
            fg=color,
            bg="#F0F0F0",
            font=("Segoe UI", 12, "bold"),
            wrap=550,
        )
        label.pack(pady=(15, 10), padx=10)

        if skipped_rows:
            inst_lbl = tk.Label(
                root,
                text="Skipped Rows (you can select and copy the text below):",
                fg="#555555",
                bg="#F0F0F0",
                font=("Segoe UI", 9, "italic"),
            )
            inst_lbl.pack(anchor="w", padx=20)

            text_frame = tk.Frame(root, bg="#F0F0F0")
            text_frame.pack(fill="both", expand=True, padx=20, pady=5)

            scrollbar = tk.Scrollbar(text_frame)
            scrollbar.pack(side="right", fill="y")

            text_box = tk.Text(
                text_frame,
                wrap="word",
                yscrollcommand=scrollbar.set,
                font=("Consolas", 10),
                bg="white",
                fg="black",
                height=10,
            )
            text_box.pack(side="left", fill="both", expand=True)
            scrollbar.config(command=text_box.yview)

            text_content = ""
            for skip in skipped_rows:
                text_content += f"- Row {skip['excel_row']} (Delay ID: {skip['delay_id']}): {skip['error']}\n"

            text_box.insert("1.0", text_content.strip())
            text_box.config(state="disabled")

        btn = tk.Button(
            root,
            text="Exit",
            command=root.destroy,
            bg="#28A745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            width=10,
            bd=0,
            padx=10,
            pady=5,
        )
        btn.pack(pady=10)

        root.focus_force()
        root.mainloop()
    except Exception:
        pass

    sys.exit(0 if success else 1)


def run_self_test():
    """Built-in smoke & regression verification suite (--self-test)."""
    print("Running P6 EOT Toolkit built-in self-test suite...")
    tmp = tempfile.mkdtemp(prefix="eot_selftest_")
    try:
        os.environ["EOT_AUTO_NO_POPUP"] = "1"
        os.environ["EOT_BATCH_MODE"] = "1"
        tpl = create_template(os.path.join(tmp, "tpl.xlsx"))
        df, cols, _ = resolve_delay_dataframe(tpl)
        assert len(df) == 2 and cols["Delay Event ID"] == "Delay Event ID"

        # Verify SpecifiedActivityList importer
        act_path = os.path.join(tmp, "activities.xlsx")
        pd.DataFrame(
            {
                "Milestone ID": ["MS-01", "ACT-10"],
                "Task Code": ["T3-CA-MS-01", "T3-ACT-1010"],
                "Milestone Description": ["Milestone 1", "Earthworks"],
            }
        ).to_excel(act_path, index=False)
        spec = load_ms_pattern_from_excel(act_path)
        assert isinstance(spec, SpecifiedActivityList) and len(spec.items) == 2

        import all_delays_EOT
        import export_EOT
        import delete_projects

        assert hasattr(all_delays_EOT, "process_combined_delay_events")
        assert hasattr(export_EOT, "build_export")
        assert hasattr(delete_projects, "delete_project_from_db")
        print("[OK] All core modules, specified-activity importers, and template parsers verified successfully.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────


def main():
    args = parse_args()
    if args.self_test:
        run_self_test()
        sys.exit(0)
    if args.generate_template is not None:
        out_path = args.generate_template if args.generate_template else None
        create_template(out_path)
        sys.exit(0)

    configure_logging(args.log_file)

    print("=" * 72)
    print("  PRIMAVERA P6 PURE SQLITE EOT ANALYSIS AUTOMATION (NO GUI)")
    print("=" * 72)

    excel_path = args.excel
    if not excel_path:
        try:
            excel_path = input("Enter the path of the delay events Excel file: ").strip().strip("'\"")
            while not excel_path or not os.path.exists(excel_path):
                if not excel_path:
                    excel_path = input("Path cannot be empty. Please enter the Excel file path: ").strip().strip("'\"")
                else:
                    excel_path = (
                        input(f"File not found at '{excel_path}'. Please enter a valid Excel file path: ")
                        .strip()
                        .strip("'\"")
                    )
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            sys.exit(0)
    else:
        excel_path = excel_path.strip("'\"")

    # Resolve specified activities list or milestone recognition pattern
    ms_pattern = resolve_milestone_pattern_interactive(args)

    source_project = args.source_project
    if not source_project:
        try:
            source_project = input("Enter the baseline project ID (PROJ_SHORT_NAME): ").strip()
            while not source_project:
                source_project = input("Project ID cannot be empty. Please enter the baseline project ID: ").strip()
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            sys.exit(0)

    db_path = args.db_path
    if not db_path or not os.path.exists(db_path):
        try:
            if db_path and db_path != DEFAULT_DB_PATH:
                print(f"[WARNING] Database not found at '{db_path}'.")
            db_path = (
                input(f"Enter the Primavera SQLite database path (default: {DEFAULT_DB_PATH}): ").strip().strip("'\"")
            )
            if not db_path:
                db_path = DEFAULT_DB_PATH
            while not db_path or not os.path.exists(db_path):
                db_path = input("Database file not found. Please enter a valid SQLite database path: ").strip().strip("'\"")
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            sys.exit(0)
    else:
        db_path = db_path.strip("'\"")
    args.db_path = db_path

    if not os.path.exists(excel_path):
        err(f"Excel file not found: {excel_path}")
        sys.exit(1)

    if not os.path.exists(args.db_path):
        err(f"Primavera database file not found at: {args.db_path}")
        sys.exit(1)

    try:
        results, skipped_rows = process_delay_events(
            excel_path=excel_path,
            db_path=args.db_path,
            source_project=source_project,
            default_remarks=args.default_remarks,
        )
    except Exception as exc:
        err(f"Database process failed: {exc}")
        sys.exit(1)

    print("\n" + "=" * 72)
    print("  ALL EOT ANALYSIS PROCESSES COMPLETE (PURE SQL)")
    print("=" * 72)

    # Generate the single consolidated Excel output file in the Combined EOT format
    explicit_out = args.output if (args.output and os.path.abspath(args.output) != os.path.abspath(OUTPUT_FILE)) else None
    try:
        run_milestone_vs_delay(args.db_path, excel_path, source_project, ms_pattern, output_path=explicit_out)
    except Exception as e:
        err(f"Failed to generate consolidated Milestone vs Delay report: {e}")
        sys.exit(1)

    return skipped_rows


if __name__ == "__main__":
    try:
        skips = main()
        show_popup(success=True, skipped_rows=skips)
    except SystemExit as e:
        if e.code != 0:
            show_popup(success=False)
        raise e
    except Exception:
        show_popup(success=False)
