#!/usr/bin/env python3
"""
P6 Project -> .XER Exporter (SQLite, P6 Professional Standalone)
==================================================================

Produces .xer files matching P6's manual "File > Export > Primavera PM (XER)" output.
Field lists are hardcoded per table to exactly match the P6 23.12 XER schema.
Virtual XER columns (suffixed with "*") are populated with P6-compliant export defaults
(such as PROJECT.export_flag='Y', PROJECT.loaded_scope_level='7', RCATTYPE.export_flag='Y',
PROJWBS.plan_open_state='Y'/'N', RSRC.load_tasks_flag='N', and TASKRSRC.has_rsrchours='N')
so Primavera P6 displays, imports, and schedules every exported project accurately.
"""

import argparse
import datetime
import decimal
import os
import re
import sqlite3
import sys

EXPORT_VERSION = "23.12"
EXPORT_DB_NAME = "dbxDatabaseNoName"
EXPORT_DB_TYPE = "Project Management"
EXPORT_CURRENCY = "USD"
EXPORT_USER = "admin"

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB_PATH = r"C:\Users\91943\Downloads\T3.db"
DEFAULT_LOG_PATH = os.path.join(_SCRIPT_DIR, "eot_process_logs.log")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


# ---------------------------------------------------------------------------
# Exact P6 23.12 field lists per table (in P6 export column order).
# A column name suffixed with "*" means it does NOT exist in the SQLite DB
# and is populated from VIRTUAL_FIELD_DEFAULTS (or NULL if None).
# ---------------------------------------------------------------------------
P6_FIELDS = {
    "CURRTYPE": [
        "curr_id", "decimal_digit_cnt", "curr_symbol", "decimal_symbol",
        "digit_group_symbol", "pos_curr_fmt_type", "neg_curr_fmt_type",
        "curr_type", "curr_short_name", "group_digit_cnt", "base_exch_rate",
    ],
    "FINTMPL": ["fintmpl_id", "fintmpl_name", "default_flag"],
    "OBS": [
        "obs_id", "parent_obs_id", "guid", "seq_num", "obs_name", "obs_descr",
    ],
    "RCATTYPE": [
        "rsrc_catg_type_id", "seq_num", "rsrc_catg_short_len",
        "rsrc_catg_type", "export_flag*",
    ],
    "RCATVAL": [
        "rsrc_catg_id", "rsrc_catg_type_id", "seq_num",
        "rsrc_catg_short_name", "rsrc_catg_name", "parent_rsrc_catg_id",
    ],
    "PROJECT": [
        "proj_id", "fy_start_month_num", "rsrc_self_add_flag",
        "allow_complete_flag", "rsrc_multi_assign_flag", "checkout_flag",
        "project_flag", "step_complete_flag", "cost_qty_recalc_flag",
        "batch_sum_flag", "name_sep_char", "def_complete_pct_type",
        "proj_short_name", "acct_id", "orig_proj_id", "source_proj_id",
        "base_type_id", "clndr_id", "sum_base_proj_id", "task_code_base",
        "task_code_step", "priority_num", "wbs_max_sum_level",
        "strgy_priority_num", "last_checksum", "critical_drtn_hr_cnt",
        "def_cost_per_qty", "last_recalc_date", "plan_start_date",
        "plan_end_date", "scd_end_date", "add_date", "last_tasksum_date",
        "fcst_start_date", "def_duration_type", "task_code_prefix", "guid",
        "def_qty_type", "add_by_name", "web_local_root_path", "proj_url",
        "def_rate_type", "add_act_remain_flag", "act_this_per_link_flag",
        "def_task_type", "act_pct_link_flag", "critical_path_type",
        "task_code_prefix_flag", "def_rollup_dates_flag",
        "use_project_baseline_flag", "rem_target_link_flag",
        "reset_planned_flag", "allow_neg_act_flag", "sum_assign_level",
        "last_fin_dates_id", "fintmpl_id", "last_baseline_update_date",
        "cr_external_key", "apply_actuals_date", "location_id",
        "last_schedule_date",
        "loaded_scope_level*", "export_flag*", "new_fin_dates_id*",
        "baselines_to_export*", "baseline_names_to_export*",
        "next_data_date*", "close_period_flag*", "sum_refresh_date*",
        "trsrcsum_loaded*", "sumtask_loaded*",
    ],
    "CALENDAR": [
        "clndr_id", "default_flag", "clndr_name", "proj_id", "base_clndr_id",
        "last_chng_date", "clndr_type", "day_hr_cnt", "week_hr_cnt",
        "month_hr_cnt", "year_hr_cnt", "rsrc_private", "clndr_data",
    ],
    "SCHEDOPTIONS": [
        "schedoptions_id", "proj_id", "sched_outer_depend_type",
        "sched_open_critical_flag", "sched_lag_early_start_flag",
        "sched_retained_logic", "sched_setplantoforecast", "sched_float_type",
        "sched_calendar_on_relationship_lag", "sched_use_expect_end_flag",
        "sched_progress_override", "level_float_thrs_cnt",
        "level_outer_assign_flag", "level_outer_assign_priority",
        "level_over_alloc_pct", "level_within_float_flag",
        "level_keep_sched_date_flag", "level_all_rsrc_flag",
        "sched_use_project_end_date_for_float",
        "enable_multiple_longest_path_calc", "limit_multiple_longest_path_calc",
        "max_multiple_longest_path", "use_total_float_multiple_longest_paths",
        "key_activity_for_multiple_longest_paths", "LevelPriorityList",
    ],
    "PROJWBS": [
        "wbs_id", "proj_id", "obs_id", "seq_num", "est_wt", "proj_node_flag",
        "sum_data_flag", "status_code", "wbs_short_name", "wbs_name",
        "phase_id", "parent_wbs_id", "ev_user_pct", "ev_etc_user_value",
        "orig_cost", "indep_remain_total_cost", "ann_dscnt_rate_pct",
        "dscnt_period_type", "indep_remain_work_qty", "anticip_start_date",
        "anticip_end_date", "ev_compute_type", "ev_etc_compute_type",
        "guid", "tmpl_guid", "plan_open_state*",
    ],
    "RSRC": [
        "rsrc_id", "parent_rsrc_id", "clndr_id", "role_id", "shift_id",
        "user_id", "pobs_id", "guid", "rsrc_seq_num", "email_addr",
        "employee_code", "office_phone", "other_phone", "rsrc_name",
        "rsrc_short_name", "rsrc_title_name", "def_qty_per_hr",
        "cost_qty_type", "ot_factor", "active_flag", "auto_compute_act_flag",
        "def_cost_qty_link_flag", "ot_flag", "curr_id", "unit_id",
        "rsrc_type", "location_id", "rsrc_notes",
        "load_tasks_flag*", "level_flag*", "last_checksum*",
    ],
    "RSRCRATE": [
        "rsrc_rate_id", "rsrc_id", "max_qty_per_hr", "cost_per_qty",
        "start_date", "shift_period_id", "cost_per_qty2", "cost_per_qty3",
        "cost_per_qty4", "cost_per_qty5",
    ],
    "RSRCRCAT": ["rsrc_id", "rsrc_catg_type_id", "rsrc_catg_id"],
    "TASK": [
        "task_id", "proj_id", "wbs_id", "clndr_id", "phys_complete_pct",
        "rev_fdbk_flag", "est_wt", "lock_plan_flag", "auto_compute_act_flag",
        "complete_pct_type", "task_type", "duration_type", "status_code",
        "task_code", "task_name", "rsrc_id", "total_float_hr_cnt",
        "free_float_hr_cnt", "remain_drtn_hr_cnt", "act_work_qty",
        "remain_work_qty", "target_work_qty", "target_drtn_hr_cnt",
        "target_equip_qty", "act_equip_qty", "remain_equip_qty",
        "cstr_date", "act_start_date", "act_end_date", "late_start_date",
        "late_end_date", "expect_end_date", "early_start_date",
        "early_end_date", "restart_date", "reend_date",
        "target_start_date", "target_end_date", "rem_late_start_date",
        "rem_late_end_date", "cstr_type", "priority_type",
        "suspend_date", "resume_date", "float_path", "float_path_order",
        "guid", "tmpl_guid", "cstr_date2", "cstr_type2",
        "driving_path_flag", "act_this_per_work_qty", "act_this_per_equip_qty",
        "external_early_start_date", "external_late_end_date",
        "create_date", "update_date", "create_user", "update_user",
        "location_id", "crt_path_num*",
    ],
    "TASKPRED": [
        "task_pred_id", "task_id", "pred_task_id", "proj_id", "pred_proj_id",
        "pred_type", "lag_hr_cnt", "comments",
        "float_path*", "aref*", "arls*",
    ],
    "TASKRSRC": [
        "taskrsrc_id", "task_id", "proj_id", "cost_qty_link_flag", "role_id",
        "acct_id", "rsrc_id", "pobs_id", "skill_level", "remain_qty",
        "target_qty", "remain_qty_per_hr", "target_lag_drtn_hr_cnt",
        "target_qty_per_hr", "act_ot_qty", "act_reg_qty",
        "relag_drtn_hr_cnt", "ot_factor", "cost_per_qty", "target_cost",
        "act_reg_cost", "act_ot_cost", "remain_cost", "act_start_date",
        "act_end_date", "restart_date", "reend_date", "target_start_date",
        "target_end_date", "rem_late_start_date", "rem_late_end_date",
        "rollup_dates_flag", "target_crv", "remain_crv", "actual_crv",
        "ts_pend_act_end_flag", "guid", "rate_type", "act_this_per_cost",
        "act_this_per_qty", "curv_id", "rsrc_type", "cost_per_qty_source_type",
        "create_user", "create_date", "cbs_id*", "has_rsrchours*",
        "taskrsrc_sum_id*",
    ],
}

# P6 XER virtual column defaults required by P6's XER Import Wizard
VIRTUAL_FIELD_DEFAULTS = {
    ("RCATTYPE", "export_flag"): "Y",
    ("PROJECT", "loaded_scope_level"): "7",
    ("PROJECT", "export_flag"): "Y",
    ("PROJECT", "close_period_flag"): "N",
    ("PROJECT", "trsrcsum_loaded"): "N",
    ("PROJECT", "sumtask_loaded"): "N",
    ("PROJWBS", "plan_open_state"): "N",
    ("RSRC", "load_tasks_flag"): "N",
    ("RSRC", "level_flag"): "N",
    ("TASKRSRC", "has_rsrchours"): "N",
}

SCHEDOPTIONS_DEFAULTS = {
    "sched_outer_depend_type": "SD_None",
    "sched_open_critical_flag": "N",
    "sched_lag_early_start_flag": "Y",
    "sched_retained_logic": "Y",
    "sched_setplantoforecast": "N",
    "sched_float_type": "FT_TotalFloat",
    "sched_calendar_on_relationship_lag": "rcal_Predecessor",
    "sched_use_expect_end_flag": "Y",
    "sched_progress_override": "N",
    "level_float_thrs_cnt": "0",
    "level_outer_assign_flag": "N",
    "level_outer_assign_priority": "10",
    "level_over_alloc_pct": "0",
    "level_within_float_flag": "Y",
    "level_keep_sched_date_flag": "Y",
    "level_all_rsrc_flag": "Y",
    "sched_use_project_end_date_for_float": "Y",
    "enable_multiple_longest_path_calc": "N",
    "limit_multiple_longest_path_calc": "Y",
    "max_multiple_longest_path": "5",
    "use_total_float_multiple_longest_paths": "Y",
}

TABLE_ORDER = [
    "CURRTYPE", "FINTMPL", "OBS", "RCATTYPE", "RCATVAL",
    "PROJECT", "CALENDAR", "SCHEDOPTIONS", "PROJWBS",
    "RSRC", "RSRCRATE", "RSRCRCAT", "TASK", "TASKPRED", "TASKRSRC",
]

DELAY_PROJ_REGEX = re.compile(r"^(DE|VE|DC)\d+$", re.IGNORECASE)

conn = None


def is_delay_project_name(name: str) -> bool:
    """Checks whether a project short name matches any delay/EOT simulation naming pattern."""
    if not name:
        return False
    return (
        "_DE" in name
        or "_VE" in name
        or "_DC" in name
        or "_EOT" in name
        or name.startswith("delay_")
        or bool(DELAY_PROJ_REGEX.match(name))
    )


def safe_exit(exit_code=0, error_msg=None):
    if "conn" in globals() and conn:
        try:
            conn.close()
        except Exception:
            pass
    if os.environ.get("EOT_AUTO_NO_POPUP") == "1" or os.environ.get("EOT_BATCH_MODE") == "1":
        sys.exit(exit_code)
    else:
        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            if error_msg:
                messagebox.showerror("Process Status", f"Error occurred: {error_msg}")
            else:
                messagebox.showinfo("Process Status", "The process completed successfully.")
            root.destroy()
            input("\nPress Enter to exit...")
        except Exception:
            pass
        sys.exit(exit_code)


def get_connection(db_path):
    global conn
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM PROJECT LIMIT 1")
        cursor.close()
        return conn
    except sqlite3.Error as e:
        print(f"[ERROR] Could not open this as a P6 database: {e}")
        safe_exit(1, f"Could not open P6 database: {e}")


def get_existing_tables(cursor):
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    return {row["name"].upper() for row in cursor.fetchall()}


def get_db_columns(cursor, table):
    cursor.execute(f"PRAGMA table_info({table})")
    return {row["name"].lower() for row in cursor.fetchall()}


def get_val(row, col_name):
    if row is None:
        return None
    for k in row.keys():
        if k.lower() == col_name.lower():
            return row[k]
    return None


_DATETIME_SECONDS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}):\d{2}(\.\d+)?$")


def format_value(value):
    """
    Formats a SQLite field value for P6 XER output.
    Crucially converts CRLF (\\r\\n), CR (\\r), and LF (\\n) to P6's single-byte XER
    in-field newline marker (\\x7f) and replaces tabs with spaces so embedded multi-line
    blobs like CALENDAR.clndr_data never split %R records across lines during P6 import.
    """
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, datetime.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, bool):
        return "Y" if value else "N"
    if isinstance(value, decimal.Decimal):
        return str(int(value)) if value == value.to_integral_value() else f"{value:.4f}"
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else f"{value:.4f}"
    if isinstance(value, bytes):
        s = value.decode("utf-8", errors="ignore")
        return s.replace("\r\n", "\x7f").replace("\r", "\x7f").replace("\n", "\x7f").replace("\t", " ")
    if isinstance(value, str):
        m = _DATETIME_SECONDS_RE.match(value)
        if m:
            return m.group(1)
        return value.replace("\r\n", "\x7f").replace("\r", "\x7f").replace("\n", "\x7f").replace("\t", " ")
    return str(value)


def build_select(table, p6_field_list, db_cols):
    """
    Build SELECT clause using P6_FIELDS list.
    Fields ending with '*' are virtual (not in SQLite DB) -> populated with
    VIRTUAL_FIELD_DEFAULTS if defined, or NULL otherwise.
    Output column alias is always the clean lowercase name (no '*').
    """
    parts = []
    clean_names = []
    table_upper = table.upper()
    for f in p6_field_list:
        is_virtual = f.endswith("*")
        clean = f.rstrip("*")
        alias = f'"{clean}"'
        if is_virtual or clean.lower() not in db_cols:
            default_val = VIRTUAL_FIELD_DEFAULTS.get((table_upper, clean))
            if default_val is not None:
                parts.append(f"'{default_val}' AS {alias}")
            else:
                parts.append(f"NULL AS {alias}")
        else:
            real = clean
            parts.append(f'"{real}" AS {alias}')
        clean_names.append(clean)
    return ", ".join(parts), clean_names


def fetch_table(cursor, table, p6_field_list, where_sql="", params=()):
    db_cols = get_db_columns(cursor, table)
    select_clause, clean_names = build_select(table, p6_field_list, db_cols)
    sql = f"SELECT {select_clause} FROM {table}"
    if where_sql:
        sql += " " + where_sql
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    result = []
    for row in rows:
        result.append([format_value(row[c]) for c in clean_names])
    return clean_names, result


def _set_row_col(fields, row_list, col_name, new_val):
    """Helper to set a column value in a formatted row list if the column exists."""
    if col_name in fields:
        row_list[fields.index(col_name)] = format_value(new_val)


def _get_row_col(fields, row_list, col_name):
    """Helper to get a column value from a formatted row list."""
    if col_name in fields:
        return row_list[fields.index(col_name)]
    return ""


def _resolve_baseline_proj_id(cursor, proj_id, existing_tables, explicit_source_proj_id=None):
    """
    Resolves the baseline source project ID for a simulated delay/EOT project copy.
    Uses explicit_source_proj_id if provided; otherwise inspects PROJECT.orig_proj_id /
    source_proj_id or matches TASK_CODE overlap against non-delay projects in the DB.
    """
    if explicit_source_proj_id is not None and str(explicit_source_proj_id) != str(proj_id):
        return explicit_source_proj_id

    if "PROJECT" not in existing_tables:
        return None

    try:
        cursor.execute("SELECT * FROM PROJECT WHERE proj_id = ?", (proj_id,))
        prow = cursor.fetchone()
        if not prow:
            return None
        pname = get_val(prow, "proj_short_name") or ""
        for fk_col in ("source_proj_id", "orig_proj_id"):
            fk_val = get_val(prow, fk_col)
            if fk_val and str(fk_val) != str(proj_id):
                cursor.execute("SELECT proj_id FROM PROJECT WHERE proj_id = ?", (fk_val,))
                if cursor.fetchone():
                    return fk_val

        delay_wbs_pids = _get_delay_wbs_proj_ids(cursor, existing_tables)
        if not is_delay_project_name(pname) and proj_id not in delay_wbs_pids:
            return None

        cursor.execute("SELECT task_code FROM TASK WHERE proj_id = ?", (proj_id,))
        my_codes = {get_val(r, "task_code") for r in cursor.fetchall() if get_val(r, "task_code")}
        if not my_codes:
            return None

        cursor.execute("SELECT proj_id, proj_short_name FROM PROJECT WHERE project_flag = 'Y' AND proj_id != ?", (proj_id,))
        best_pid = None
        best_overlap = 0
        for cand in cursor.fetchall():
            cpid = get_val(cand, "proj_id")
            cname = get_val(cand, "proj_short_name") or ""
            if is_delay_project_name(cname) or cpid in delay_wbs_pids:
                continue
            cursor.execute("SELECT task_code FROM TASK WHERE proj_id = ?", (cpid,))
            ccodes = {get_val(r, "task_code") for r in cursor.fetchall() if get_val(r, "task_code")}
            overlap = len(my_codes & ccodes)
            if overlap > best_overlap:
                best_overlap = overlap
                best_pid = cpid
        return best_pid
    except Exception:
        return None


def get_project_obs_ids(cursor, proj_id, existing_tables):
    """Collects all OBS IDs referenced by PROJWBS in this project."""
    if "PROJWBS" not in existing_tables:
        return []
    db_cols = get_db_columns(cursor, "PROJWBS")
    if "obs_id" not in db_cols:
        return []
    cursor.execute(
        "SELECT DISTINCT obs_id FROM PROJWBS WHERE proj_id = ? AND obs_id IS NOT NULL",
        (proj_id,),
    )
    return [get_val(r, "obs_id") for r in cursor.fetchall() if get_val(r, "obs_id") is not None]


def get_project_root_obs_id(cursor, proj_id, existing_tables):
    if "PROJWBS" not in existing_tables:
        return None
    db_cols = get_db_columns(cursor, "PROJWBS")
    if "obs_id" not in db_cols or "proj_node_flag" not in db_cols:
        return None
    cursor.execute(
        "SELECT obs_id FROM PROJWBS WHERE proj_id = ? AND proj_node_flag = 'Y'",
        (proj_id,),
    )
    row = cursor.fetchone()
    return get_val(row, "obs_id") if row else None


def get_direct_rsrc_ids(cursor, proj_id, existing_tables=None, source_proj_id=None):
    """rsrc_ids directly assigned to tasks (via TASK or TASKRSRC) in this project."""
    ids = set()
    cursor.execute(
        "SELECT DISTINCT rsrc_id FROM TASK WHERE proj_id = ? AND rsrc_id IS NOT NULL",
        (proj_id,),
    )
    ids.update(get_val(row, "rsrc_id") for row in cursor.fetchall() if get_val(row, "rsrc_id"))

    if existing_tables and "TASKRSRC" in existing_tables:
        db_cols = get_db_columns(cursor, "TASKRSRC")
        if "rsrc_id" in db_cols:
            cursor.execute(
                "SELECT DISTINCT rsrc_id FROM TASKRSRC WHERE proj_id = ? AND rsrc_id IS NOT NULL",
                (proj_id,),
            )
            tr_ids = {get_val(row, "rsrc_id") for row in cursor.fetchall() if get_val(row, "rsrc_id")}
            if not tr_ids and source_proj_id is not None and str(source_proj_id) != str(proj_id):
                cursor.execute(
                    "SELECT DISTINCT rsrc_id FROM TASKRSRC WHERE proj_id = ? AND rsrc_id IS NOT NULL",
                    (source_proj_id,),
                )
                tr_ids = {get_val(row, "rsrc_id") for row in cursor.fetchall() if get_val(row, "rsrc_id")}
            ids.update(tr_ids)
    return ids


def get_all_rsrc_ids(cursor, proj_id, existing_tables, source_proj_id=None):
    """Direct task rsrc_ids PLUS their parent_rsrc_id ancestors (for RSRC table)."""
    ids = get_direct_rsrc_ids(cursor, proj_id, existing_tables, source_proj_id)
    if not ids or "RSRC" not in existing_tables:
        return ids
    db_cols = get_db_columns(cursor, "RSRC")
    if "parent_rsrc_id" not in db_cols:
        return ids
    expanded, frontier = set(ids), set(ids)
    while frontier:
        ph = ",".join("?" * len(frontier))
        cursor.execute(
            f"SELECT DISTINCT parent_rsrc_id FROM RSRC "
            f"WHERE rsrc_id IN ({ph}) AND parent_rsrc_id IS NOT NULL",
            tuple(frontier),
        )
        new_ids = {get_val(r, "parent_rsrc_id") for r in cursor.fetchall() if get_val(r, "parent_rsrc_id")} - expanded
        expanded |= new_ids
        frontier = new_ids
    return expanded


def get_calendar_ids(cursor, proj_id, existing_tables, source_proj_id=None):
    ids = set()
    if "PROJECT" in existing_tables:
        cursor.execute("SELECT clndr_id FROM PROJECT WHERE proj_id = ?", (proj_id,))
        r = cursor.fetchone()
        if r and get_val(r, "clndr_id"):
            ids.add(get_val(r, "clndr_id"))
        elif source_proj_id is not None and str(source_proj_id) != str(proj_id):
            cursor.execute("SELECT clndr_id FROM PROJECT WHERE proj_id = ?", (source_proj_id,))
            sr = cursor.fetchone()
            if sr and get_val(sr, "clndr_id"):
                ids.add(get_val(sr, "clndr_id"))
    if "TASK" in existing_tables:
        cursor.execute(
            "SELECT DISTINCT clndr_id FROM TASK WHERE proj_id = ? AND clndr_id IS NOT NULL",
            (proj_id,),
        )
        ids.update(get_val(row, "clndr_id") for row in cursor.fetchall() if get_val(row, "clndr_id"))
    if "RSRC" in existing_tables:
        db_cols = get_db_columns(cursor, "RSRC")
        if "clndr_id" in db_cols:
            all_rids = get_all_rsrc_ids(cursor, proj_id, existing_tables, source_proj_id)
            if all_rids:
                ph = ",".join("?" * len(all_rids))
                cursor.execute(
                    f"SELECT DISTINCT clndr_id FROM RSRC WHERE rsrc_id IN ({ph}) AND clndr_id IS NOT NULL",
                    tuple(all_rids),
                )
                ids.update(get_val(row, "clndr_id") for row in cursor.fetchall() if get_val(row, "clndr_id"))
    if ids and "CALENDAR" in existing_tables:
        db_cols = get_db_columns(cursor, "CALENDAR")
        if "base_clndr_id" in db_cols:
            expanded, frontier = set(ids), set(ids)
            while frontier:
                ph = ",".join("?" * len(frontier))
                cursor.execute(
                    f"SELECT DISTINCT base_clndr_id FROM CALENDAR "
                    f"WHERE clndr_id IN ({ph}) AND base_clndr_id IS NOT NULL",
                    tuple(frontier),
                )
                new_ids = {
                    get_val(r, "base_clndr_id") for r in cursor.fetchall() if get_val(r, "base_clndr_id")
                } - expanded
                expanded |= new_ids
                frontier = new_ids
            return expanded
    return ids


def get_rcatval_ids(cursor, proj_id, direct_rsrc_ids, existing_tables):
    """rsrc_catg_ids used by direct task resources, plus parent_rsrc_catg_id ancestors."""
    if "RSRCRCAT" not in existing_tables or not direct_rsrc_ids:
        return set()
    ph = ",".join("?" * len(direct_rsrc_ids))
    cursor.execute(
        f"SELECT DISTINCT rsrc_catg_id FROM RSRCRCAT WHERE rsrc_id IN ({ph})",
        tuple(direct_rsrc_ids),
    )
    ids = {get_val(r, "rsrc_catg_id") for r in cursor.fetchall() if get_val(r, "rsrc_catg_id")}
    if ids and "RCATVAL" in existing_tables:
        db_cols = get_db_columns(cursor, "RCATVAL")
        if "parent_rsrc_catg_id" in db_cols:
            expanded, frontier = set(ids), set(ids)
            while frontier:
                ph = ",".join("?" * len(frontier))
                cursor.execute(
                    f"SELECT DISTINCT parent_rsrc_catg_id FROM RCATVAL "
                    f"WHERE rsrc_catg_id IN ({ph}) AND parent_rsrc_catg_id IS NOT NULL",
                    tuple(frontier),
                )
                new_ids = {
                    get_val(r, "parent_rsrc_catg_id")
                    for r in cursor.fetchall()
                    if get_val(r, "parent_rsrc_catg_id")
                } - expanded
                expanded |= new_ids
                frontier = new_ids
            return expanded
    return ids


def _sort_calendars_topologically(fields, rows, proj_id_str):
    """
    Ensures project calendars (CA_Proj) have proj_id set to the exported proj_id,
    global/resource calendars (CA_Base, CA_Rsrc) have proj_id cleared to '',
    clears self-referential base_clndr_id, and orders base calendars before derived calendars.
    """
    if not rows:
        return rows
    sanitized = []
    for r in rows:
        row_copy = list(r)
        cid = _get_row_col(fields, row_copy, "clndr_id")
        bid = _get_row_col(fields, row_copy, "base_clndr_id")
        ctype = _get_row_col(fields, row_copy, "clndr_type")

        if bid and bid == cid:
            _set_row_col(fields, row_copy, "base_clndr_id", "")
            bid = ""
        if ctype == "CA_Proj":
            _set_row_col(fields, row_copy, "proj_id", proj_id_str)
        else:
            _set_row_col(fields, row_copy, "proj_id", "")
        sanitized.append(row_copy)

    # Topological sort: base calendars (no base_clndr_id or base_clndr_id not in set) first
    by_id = {_get_row_col(fields, r, "clndr_id"): r for r in sanitized}
    ordered = []
    emitted = set()
    remaining = list(sanitized)
    while remaining:
        progressed = False
        next_remaining = []
        for r in remaining:
            cid = _get_row_col(fields, r, "clndr_id")
            bid = _get_row_col(fields, r, "base_clndr_id")
            if not bid or bid not in by_id or bid in emitted:
                ordered.append(r)
                emitted.add(cid)
                progressed = True
            else:
                next_remaining.append(r)
        if not progressed:
            ordered.extend(next_remaining)
            break
        remaining = next_remaining
    return ordered


def _sort_and_sanitize_projwbs(fields, rows, proj_id_str, default_obs_id_str):
    """
    Orders PROJWBS rows so the project root WBS node (proj_node_flag == 'Y') is always first,
    followed by child WBS nodes in topological parent-before-child order, and populates
    required P6 WBS defaults.
    """
    if not rows:
        return rows

    wbs_id_set = {_get_row_col(fields, r, "wbs_id") for r in rows}
    root_wbs_id = None
    for r in rows:
        if _get_row_col(fields, r, "proj_node_flag") == "Y":
            root_wbs_id = _get_row_col(fields, r, "wbs_id")
            break

    sanitized = []
    for r in rows:
        row_copy = list(r)
        is_root = _get_row_col(fields, row_copy, "proj_node_flag") == "Y"
        _set_row_col(fields, row_copy, "proj_id", proj_id_str)
        # Clear phase_id since PHASE dictionary table is not exported
        _set_row_col(fields, row_copy, "phase_id", "")
        if not _get_row_col(fields, row_copy, "obs_id") and default_obs_id_str:
            _set_row_col(fields, row_copy, "obs_id", default_obs_id_str)
        if not _get_row_col(fields, row_copy, "status_code"):
            _set_row_col(fields, row_copy, "status_code", "WS_Open")
        if not _get_row_col(fields, row_copy, "est_wt"):
            _set_row_col(fields, row_copy, "est_wt", "1")
        if not _get_row_col(fields, row_copy, "ev_user_pct"):
            _set_row_col(fields, row_copy, "ev_user_pct", "0")
        if not _get_row_col(fields, row_copy, "ev_etc_user_value"):
            _set_row_col(fields, row_copy, "ev_etc_user_value", "1")
        if not _get_row_col(fields, row_copy, "orig_cost"):
            _set_row_col(fields, row_copy, "orig_cost", "0")
        if not _get_row_col(fields, row_copy, "indep_remain_total_cost"):
            _set_row_col(fields, row_copy, "indep_remain_total_cost", "0")
        if not _get_row_col(fields, row_copy, "indep_remain_work_qty"):
            _set_row_col(fields, row_copy, "indep_remain_work_qty", "0")
        if not _get_row_col(fields, row_copy, "ev_compute_type"):
            _set_row_col(fields, row_copy, "ev_compute_type", "EC_Cmp_pct")
        if not _get_row_col(fields, row_copy, "ev_etc_compute_type"):
            _set_row_col(fields, row_copy, "ev_etc_compute_type", "EE_Rem_hr")

        if is_root:
            _set_row_col(fields, row_copy, "plan_open_state", "Y")
        else:
            _set_row_col(fields, row_copy, "plan_open_state", "N")
            parent_id = _get_row_col(fields, row_copy, "parent_wbs_id")
            if (not parent_id or parent_id not in wbs_id_set) and root_wbs_id:
                _set_row_col(fields, row_copy, "parent_wbs_id", root_wbs_id)

        sanitized.append(row_copy)

    def sort_key(r):
        seq_str = _get_row_col(fields, r, "seq_num")
        wid_str = _get_row_col(fields, r, "wbs_id")
        try:
            seq_val = int(float(seq_str))
        except Exception:
            seq_val = 999999
        try:
            wid_val = int(float(wid_str))
        except Exception:
            wid_val = 999999
        return (seq_val, wid_val)

    roots = [r for r in sanitized if _get_row_col(fields, r, "proj_node_flag") == "Y"]
    non_roots = [r for r in sanitized if _get_row_col(fields, r, "proj_node_flag") != "Y"]
    roots.sort(key=sort_key)
    non_roots.sort(key=sort_key)

    ordered = list(roots)
    emitted = {_get_row_col(fields, r, "wbs_id") for r in ordered}
    remaining = list(non_roots)
    while remaining:
        progressed = False
        next_remaining = []
        for r in remaining:
            wid = _get_row_col(fields, r, "wbs_id")
            pid = _get_row_col(fields, r, "parent_wbs_id")
            if not pid or pid not in wbs_id_set or pid in emitted:
                ordered.append(r)
                emitted.add(wid)
                progressed = True
            else:
                next_remaining.append(r)
        if not progressed:
            ordered.extend(next_remaining)
            break
        remaining = next_remaining

    return ordered



def normalize_lag_calendar_option(val):
    """
    Normalizes P6 'Calendar for Scheduling Relationship Lag' values to canonical
    P6 XER / PROJPROP codes:
      - 'rcal_Predecessor' (Predecessor Activity Calendar)
      - 'rcal_Successor' (Successor Activity Calendar)
      - 'rcal_ProjDefault' (Project Default Calendar)
      - 'rcal_24Hour' (24 Hour Calendar)
    """
    if not val:
        return "rcal_Predecessor"
    s = str(val).strip()
    sl = s.lower().replace("_", "").replace("-", "").replace(" ", "")
    if sl in ("rcalpredecessor", "predecessor", "predecessorcalendar", "predecessoractivitycalendar", "0"):
        return "rcal_Predecessor"
    if sl in ("rcalsuccessor", "successor", "successorcalendar", "successoractivitycalendar", "1"):
        return "rcal_Successor"
    if sl in ("rcalprojdefault", "rcalprojectdefault", "projectdefault", "projectdefaultcalendar", "projdefault", "2"):
        return "rcal_ProjDefault"
    if sl in ("rcal24hour", "rcal24hours", "24hour", "24hours", "24hourcalendar", "24hrcalendar", "twentyfourhour", "3"):
        return "rcal_24Hour"
    return s


def parse_p6_structured_text(raw):
    """
    Parses Primavera P6 StructuredText records stored in PROJPROP.prop_value
    (e.g. prop_name='schedoptions': `(0||142(sched_outer_depend_type|SD_None|...|sched_calendar_on_relationship_lag|rcal_24Hour|...)())`)
    into a dictionary of attribute names and values.
    """
    if raw is None:
        return {}
    if isinstance(raw, bytes):
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("cp1252", errors="ignore")
    else:
        text = str(raw)
    text = text.strip()
    if not text:
        return {}

    attrs = {}
    n = len(text)

    def skip_ws(p):
        while p < n and text[p] in " \t\r\n\x7f":
            p += 1
        return p

    def parse_record(p):
        p = skip_ws(p)
        if p >= n or text[p] != "(":
            return p
        p += 1
        while p < n and text[p:p+2] != "||" and text[p] not in "()":
            p += 1
        if text[p:p+2] == "||":
            p += 2
        rec_name_start = p
        while p < n and text[p] not in "()":
            p += 1
        rec_name = text[rec_name_start:p].strip()
        if rec_name and "schedoptions_id" not in attrs and rec_name.isdigit():
            attrs["schedoptions_id"] = rec_name

        p = skip_ws(p)
        if p < n and text[p] == "(":
            p += 1
            while p < n:
                p = skip_ws(p)
                if p >= n or text[p] == ")":
                    break
                k_start = p
                while p < n and text[p] not in "|)":
                    p += 1
                key = text[k_start:p].strip()
                if p >= n or text[p] == ")":
                    break
                p += 1  # skip |
                v_start = p
                paren_depth = 0
                while p < n:
                    ch = text[p]
                    if ch == "(":
                        paren_depth += 1
                    elif ch == ")":
                        if paren_depth > 0:
                            paren_depth -= 1
                        else:
                            break
                    elif ch == "|" and paren_depth == 0:
                        break
                    p += 1
                val = text[v_start:p].strip()
                if key:
                    attrs[key] = val
                if p < n and text[p] == "|":
                    p += 1
            if p < n and text[p] == ")":
                p += 1

        p = skip_ws(p)
        if p < n and text[p] == "(":
            p += 1
            while p < n:
                p = skip_ws(p)
                if p >= n or text[p] == ")":
                    break
                if text[p] == "(":
                    p = parse_record(p)
                else:
                    p += 1
            if p < n and text[p] == ")":
                p += 1

        p = skip_ws(p)
        if p < n and text[p] == ")":
            p += 1
        return p

    idx = text.find("(")
    if idx != -1:
        parse_record(idx)

    sched_cols = [f.rstrip("*") for f in P6_FIELDS["SCHEDOPTIONS"]]
    lower_existing = {k.lower() for k in attrs}
    for col in sched_cols:
        if col.lower() == "levelprioritylist":
            continue
        if col.lower() not in lower_existing:
            m = re.search(rf"(?i)\b{re.escape(col)}\s*[|=:]\s*([^\s|)\"\x27,}}]*)", text)
            if m:
                attrs[col] = m.group(1).strip()

    return attrs


def _extract_projprop_schedoptions(cursor, target_proj_id, existing_tables):
    """
    Extracts schedule options from the SQLite PROJPROP table (where P6 stores
    project Schedule Options under prop_name='schedoptions').
    """
    if "PROJPROP" not in existing_tables:
        return {}
    sched_cols = [f.rstrip("*") for f in P6_FIELDS["SCHEDOPTIONS"]]
    canonical_by_lower = {c.lower(): c for c in sched_cols}
    try:
        if target_proj_id is not None:
            cursor.execute("SELECT * FROM PROJPROP WHERE proj_id = ?", (target_proj_id,))
        else:
            cursor.execute("SELECT * FROM PROJPROP")
        rows = cursor.fetchall()
    except Exception:
        return {}

    if not rows:
        return {}

    # Prioritize rows with prop_name == 'schedoptions' (or similar), then any row containing sched_ keys
    prioritized = []
    secondary = []
    for r in rows:
        pname = str(get_val(r, "prop_name") or "").strip().lower()
        pval = get_val(r, "prop_value")
        if not pval:
            continue
        if pname in ("schedoptions", "sched_options", "scheduling", "scheduleoptions"):
            prioritized.append(pval)
        else:
            pval_str = pval.decode("utf-8", errors="ignore") if isinstance(pval, bytes) else str(pval)
            if "sched_" in pval_str.lower() or "rcal_" in pval_str.lower():
                secondary.append(pval)

    extracted = {}
    for pval in prioritized + secondary:
        parsed = parse_p6_structured_text(pval)
        for k, v in parsed.items():
            canon = canonical_by_lower.get(k.lower())
            if canon and v is not None and str(v) != "" and canon not in extracted:
                extracted[canon] = format_value(v)
    return extracted


def _extract_table_schedoptions(cursor, target_proj_id, existing_tables):
    """
    Extracts schedule options from a SCHEDOPTIONS SQLite table if one exists.
    """
    if "SCHEDOPTIONS" not in existing_tables:
        return {}
    try:
        if target_proj_id is not None:
            fields, rows = fetch_table(
                cursor, "SCHEDOPTIONS", P6_FIELDS["SCHEDOPTIONS"], "WHERE proj_id = ?", (target_proj_id,)
            )
        else:
            fields, rows = fetch_table(cursor, "SCHEDOPTIONS", P6_FIELDS["SCHEDOPTIONS"], "LIMIT 1")
        if rows:
            return {col: val for col, val in zip(fields, rows[0]) if val is not None and str(val) != ""}
    except Exception:
        pass
    return {}


def _build_schedoptions_rows(cursor, proj_id, baseline_proj_id, existing_tables, proj_id_str):
    """
    Synthesizes the %T SCHEDOPTIONS table for XER export by reading P6's PROJPROP
    table (prop_name='schedoptions') as well as any SCHEDOPTIONS table if present.
    Always returns a valid SCHEDOPTIONS row so Primavera P6 never resets project
    Schedule Options (such as Calendar for Scheduling Relationship Lag) to default.
    """
    fields = [f.rstrip("*") for f in P6_FIELDS["SCHEDOPTIONS"]]
    merged = {col: "" for col in fields}
    for col_name, def_val in SCHEDOPTIONS_DEFAULTS.items():
        merged[col_name] = def_val

    proj_table_opts = _extract_table_schedoptions(cursor, proj_id, existing_tables)
    proj_prop_opts = _extract_projprop_schedoptions(cursor, proj_id, existing_tables)

    base_table_opts = (
        _extract_table_schedoptions(cursor, baseline_proj_id, existing_tables)
        if baseline_proj_id is not None
        else {}
    )
    base_prop_opts = (
        _extract_projprop_schedoptions(cursor, baseline_proj_id, existing_tables)
        if baseline_proj_id is not None
        else {}
    )

    if not proj_table_opts and not proj_prop_opts and not base_table_opts and not base_prop_opts:
        fallback_table = _extract_table_schedoptions(cursor, None, existing_tables)
        fallback_prop = _extract_projprop_schedoptions(cursor, None, existing_tables)
        merged.update(fallback_table)
        merged.update(fallback_prop)
    else:
        # Apply target project options first, then overlay baseline project options
        # (if exporting a delay project copy) so Schedule Options always match the baseline project,
        # with PROJPROP taking highest priority on real P6 SQLite databases.
        merged.update(base_table_opts)
        merged.update(proj_table_opts)
        merged.update(base_prop_opts)
        merged.update(proj_prop_opts)

    merged["proj_id"] = proj_id_str
    if not merged.get("schedoptions_id"):
        merged["schedoptions_id"] = proj_id_str or "1"

    merged["sched_calendar_on_relationship_lag"] = normalize_lag_calendar_option(
        merged.get("sched_calendar_on_relationship_lag")
    )

    # Validate or remap key_activity_for_multiple_longest_paths to the exported project's task_id
    key_act = merged.get("key_activity_for_multiple_longest_paths") or ""
    if key_act and "TASK" in existing_tables:
        cursor.execute("SELECT task_id FROM TASK WHERE proj_id = ?", (proj_id,))
        proj_task_ids = {format_value(get_val(r, "task_id")) for r in cursor.fetchall()}
        if key_act not in proj_task_ids:
            remapped_key_act = ""
            if baseline_proj_id is not None:
                try:
                    cursor.execute(
                        "SELECT task_code FROM TASK WHERE proj_id = ? AND task_id = ?",
                        (baseline_proj_id, key_act),
                    )
                    brow = cursor.fetchone()
                    bcode = get_val(brow, "task_code") if brow else None
                    if bcode:
                        cursor.execute(
                            "SELECT task_id FROM TASK WHERE proj_id = ? AND task_code = ?",
                            (proj_id, bcode),
                        )
                        nrow = cursor.fetchone()
                        if nrow and get_val(nrow, "task_id") is not None:
                            remapped_key_act = format_value(get_val(nrow, "task_id"))
                except Exception:
                    remapped_key_act = ""
            merged["key_activity_for_multiple_longest_paths"] = remapped_key_act
    elif key_act:
        merged["key_activity_for_multiple_longest_paths"] = ""

    row_list = [format_value(merged.get(col, "")) for col in fields]
    return fields, [row_list]


def build_export(cursor, proj_id, existing_tables, source_proj_id=None):
    tables_out = []
    proj_id_str = format_value(proj_id)
    baseline_proj_id = _resolve_baseline_proj_id(cursor, proj_id, existing_tables, source_proj_id)
    direct_rsrc_ids = get_direct_rsrc_ids(cursor, proj_id, existing_tables, baseline_proj_id)

    def add(name, fields, rows):
        if rows:
            tables_out.append((name, fields, rows))

    # CURRTYPE
    if "CURRTYPE" in existing_tables:
        fields, rows = fetch_table(cursor, "CURRTYPE", P6_FIELDS["CURRTYPE"], "ORDER BY curr_id")
        add("CURRTYPE", fields, rows)

    # FINTMPL
    exported_fintmpl_id_str = ""
    if "FINTMPL" in existing_tables:
        cursor.execute("SELECT fintmpl_id FROM PROJECT WHERE proj_id = ?", (proj_id,))
        r = cursor.fetchone()
        fid = get_val(r, "fintmpl_id") if r else None
        fields, rows = [], []
        if fid is not None:
            fields, rows = fetch_table(cursor, "FINTMPL", P6_FIELDS["FINTMPL"], "WHERE fintmpl_id = ?", (fid,))
        if not rows:
            fields, rows = fetch_table(cursor, "FINTMPL", P6_FIELDS["FINTMPL"], "LIMIT 1")
        if rows:
            exported_fintmpl_id_str = _get_row_col(fields, rows[0], "fintmpl_id")
            add("FINTMPL", fields, rows)

    # OBS - walk ancestor chain for all OBS IDs referenced by PROJWBS in this project
    default_obs_id_str = ""
    if "OBS" in existing_tables:
        db_cols = get_db_columns(cursor, "OBS")
        p6_list = P6_FIELDS["OBS"]
        select_clause, clean_names = build_select("OBS", p6_list, db_cols)
        seed_obs_ids = get_project_obs_ids(cursor, proj_id, existing_tables)
        root_obs_id = get_project_root_obs_id(cursor, proj_id, existing_tables)
        if root_obs_id is not None and root_obs_id not in seed_obs_ids:
            seed_obs_ids.insert(0, root_obs_id)

        obs_by_id = {}
        for start_id in seed_obs_ids:
            current = start_id
            while current is not None and current not in obs_by_id:
                cursor.execute(f"SELECT {select_clause} FROM OBS WHERE obs_id = ?", (current,))
                row = cursor.fetchone()
                if not row:
                    break
                obs_by_id[current] = [format_value(row[c]) for c in clean_names]
                current = row["parent_obs_id"] if "parent_obs_id" in row.keys() else None

        obs_rows = []
        if obs_by_id:
            # Order parent-before-child
            emitted_obs = set()
            remaining_obs = list(obs_by_id.values())
            id_set = {format_value(k) for k in obs_by_id.keys()}
            while remaining_obs:
                progressed = False
                next_rem = []
                for r in remaining_obs:
                    oid = _get_row_col(clean_names, r, "obs_id")
                    poid = _get_row_col(clean_names, r, "parent_obs_id")
                    if not poid or poid not in id_set or poid in emitted_obs:
                        obs_rows.append(r)
                        emitted_obs.add(oid)
                        progressed = True
                    else:
                        next_rem.append(r)
                if not progressed:
                    obs_rows.extend(next_rem)
                    break
                remaining_obs = next_rem
        else:
            _, obs_rows = fetch_table(cursor, "OBS", P6_FIELDS["OBS"], "ORDER BY seq_num LIMIT 1")

        if obs_rows:
            default_obs_id_str = _get_row_col(clean_names, obs_rows[-1], "obs_id")
        add("OBS", clean_names, obs_rows)

    # RCATTYPE - filtered by types used by direct task resources
    if "RCATTYPE" in existing_tables and "RSRCRCAT" in existing_tables:
        if direct_rsrc_ids:
            ph = ",".join("?" * len(direct_rsrc_ids))
            cursor.execute(
                f"SELECT DISTINCT rsrc_catg_type_id FROM RSRCRCAT WHERE rsrc_id IN ({ph})",
                tuple(direct_rsrc_ids),
            )
            type_ids = {get_val(r, "rsrc_catg_type_id") for r in cursor.fetchall() if get_val(r, "rsrc_catg_type_id")}
            if type_ids:
                ph2 = ",".join("?" * len(type_ids))
                fields, rows = fetch_table(
                    cursor,
                    "RCATTYPE",
                    P6_FIELDS["RCATTYPE"],
                    f"WHERE rsrc_catg_type_id IN ({ph2}) ORDER BY seq_num",
                    tuple(type_ids),
                )
                add("RCATTYPE", fields, rows)

    # RCATVAL - with parent ancestor walk
    if "RCATVAL" in existing_tables:
        catg_ids = get_rcatval_ids(cursor, proj_id, direct_rsrc_ids, existing_tables)
        if catg_ids:
            ph = ",".join("?" * len(catg_ids))
            fields, rows = fetch_table(
                cursor,
                "RCATVAL",
                P6_FIELDS["RCATVAL"],
                f"WHERE rsrc_catg_id IN ({ph}) ORDER BY seq_num",
                tuple(catg_ids),
            )
            add("RCATVAL", fields, rows)

    # PROJECT
    if "PROJECT" in existing_tables:
        fields, rows = fetch_table(cursor, "PROJECT", P6_FIELDS["PROJECT"], "WHERE proj_id = ?", (proj_id,))
        sanitized_proj_rows = []
        for r in rows:
            row_copy = list(r)

            _set_row_col(fields, row_copy, "proj_id", proj_id_str)
            _set_row_col(fields, row_copy, "project_flag", "Y")
            _set_row_col(fields, row_copy, "export_flag", "Y")
            _set_row_col(fields, row_copy, "loaded_scope_level", "7")
            _set_row_col(fields, row_copy, "checkout_flag", "N")
            _set_row_col(fields, row_copy, "close_period_flag", "N")
            _set_row_col(fields, row_copy, "trsrcsum_loaded", "N")
            _set_row_col(fields, row_copy, "sumtask_loaded", "N")
            # Clear external/baseline project foreign keys not included in single-project XER export
            _set_row_col(fields, row_copy, "orig_proj_id", "")
            _set_row_col(fields, row_copy, "source_proj_id", "")
            _set_row_col(fields, row_copy, "base_type_id", "")
            _set_row_col(fields, row_copy, "sum_base_proj_id", "")
            _set_row_col(fields, row_copy, "acct_id", "")
            _set_row_col(fields, row_copy, "location_id", "")
            _set_row_col(fields, row_copy, "last_fin_dates_id", "")
            _set_row_col(fields, row_copy, "fintmpl_id", exported_fintmpl_id_str)
            if not _get_row_col(fields, row_copy, "fy_start_month_num"):
                _set_row_col(fields, row_copy, "fy_start_month_num", "1")
            if not _get_row_col(fields, row_copy, "rsrc_self_add_flag"):
                _set_row_col(fields, row_copy, "rsrc_self_add_flag", "Y")
            if not _get_row_col(fields, row_copy, "allow_complete_flag"):
                _set_row_col(fields, row_copy, "allow_complete_flag", "Y")
            if not _get_row_col(fields, row_copy, "rsrc_multi_assign_flag"):
                _set_row_col(fields, row_copy, "rsrc_multi_assign_flag", "Y")
            if not _get_row_col(fields, row_copy, "step_complete_flag"):
                _set_row_col(fields, row_copy, "step_complete_flag", "N")
            if not _get_row_col(fields, row_copy, "cost_qty_recalc_flag"):
                _set_row_col(fields, row_copy, "cost_qty_recalc_flag", "Y")
            if not _get_row_col(fields, row_copy, "batch_sum_flag"):
                _set_row_col(fields, row_copy, "batch_sum_flag", "Y")
            if not _get_row_col(fields, row_copy, "name_sep_char"):
                _set_row_col(fields, row_copy, "name_sep_char", ".")
            if not _get_row_col(fields, row_copy, "def_complete_pct_type"):
                _set_row_col(fields, row_copy, "def_complete_pct_type", "CP_Drtn")
            if not _get_row_col(fields, row_copy, "def_duration_type"):
                _set_row_col(fields, row_copy, "def_duration_type", "DT_FixedDrtn")
            if not _get_row_col(fields, row_copy, "def_qty_type"):
                _set_row_col(fields, row_copy, "def_qty_type", "QT_Hour")
            if not _get_row_col(fields, row_copy, "def_rate_type"):
                _set_row_col(fields, row_copy, "def_rate_type", "COST_PER_QTY")
            if not _get_row_col(fields, row_copy, "def_task_type"):
                _set_row_col(fields, row_copy, "def_task_type", "TT_Task")
            if not _get_row_col(fields, row_copy, "critical_path_type"):
                _set_row_col(fields, row_copy, "critical_path_type", "CT_TotFloat")
            if not _get_row_col(fields, row_copy, "task_code_base"):
                _set_row_col(fields, row_copy, "task_code_base", "1000")
            if not _get_row_col(fields, row_copy, "task_code_step"):
                _set_row_col(fields, row_copy, "task_code_step", "10")
            if not _get_row_col(fields, row_copy, "task_code_prefix"):
                _set_row_col(fields, row_copy, "task_code_prefix", "A")
            if not _get_row_col(fields, row_copy, "task_code_prefix_flag"):
                _set_row_col(fields, row_copy, "task_code_prefix_flag", "Y")
            if not _get_row_col(fields, row_copy, "priority_num"):
                _set_row_col(fields, row_copy, "priority_num", "10")
            if not _get_row_col(fields, row_copy, "wbs_max_sum_level"):
                _set_row_col(fields, row_copy, "wbs_max_sum_level", "0")
            if not _get_row_col(fields, row_copy, "strgy_priority_num"):
                _set_row_col(fields, row_copy, "strgy_priority_num", "500")
            if not _get_row_col(fields, row_copy, "critical_drtn_hr_cnt"):
                _set_row_col(fields, row_copy, "critical_drtn_hr_cnt", "0")
            if not _get_row_col(fields, row_copy, "def_cost_per_qty"):
                _set_row_col(fields, row_copy, "def_cost_per_qty", "0")
            if not _get_row_col(fields, row_copy, "add_act_remain_flag"):
                _set_row_col(fields, row_copy, "add_act_remain_flag", "N")
            if not _get_row_col(fields, row_copy, "act_this_per_link_flag"):
                _set_row_col(fields, row_copy, "act_this_per_link_flag", "Y")
            if not _get_row_col(fields, row_copy, "act_pct_link_flag"):
                _set_row_col(fields, row_copy, "act_pct_link_flag", "Y")
            if not _get_row_col(fields, row_copy, "def_rollup_dates_flag"):
                _set_row_col(fields, row_copy, "def_rollup_dates_flag", "Y")
            if not _get_row_col(fields, row_copy, "use_project_baseline_flag"):
                _set_row_col(fields, row_copy, "use_project_baseline_flag", "Y")
            if not _get_row_col(fields, row_copy, "rem_target_link_flag"):
                _set_row_col(fields, row_copy, "rem_target_link_flag", "Y")
            if not _get_row_col(fields, row_copy, "reset_planned_flag"):
                _set_row_col(fields, row_copy, "reset_planned_flag", "N")
            if not _get_row_col(fields, row_copy, "allow_neg_act_flag"):
                _set_row_col(fields, row_copy, "allow_neg_act_flag", "N")
            if not _get_row_col(fields, row_copy, "sum_assign_level"):
                _set_row_col(fields, row_copy, "sum_assign_level", "SL_Taskrsrc")
            sanitized_proj_rows.append(row_copy)
        add("PROJECT", fields, sanitized_proj_rows)

    # CALENDAR
    if "CALENDAR" in existing_tables:
        clndr_ids = get_calendar_ids(cursor, proj_id, existing_tables, baseline_proj_id)
        if clndr_ids:
            ph = ",".join("?" * len(clndr_ids))
            fields, rows = fetch_table(
                cursor,
                "CALENDAR",
                P6_FIELDS["CALENDAR"],
                f"WHERE clndr_id IN ({ph}) ORDER BY clndr_id",
                tuple(clndr_ids),
            )
            rows = _sort_calendars_topologically(fields, rows, proj_id_str)
            add("CALENDAR", fields, rows)

    # SCHEDOPTIONS - synthesized from PROJPROP (prop_name='schedoptions') and/or SCHEDOPTIONS table
    sched_fields, sched_rows = _build_schedoptions_rows(
        cursor, proj_id, baseline_proj_id, existing_tables, proj_id_str
    )
    add("SCHEDOPTIONS", sched_fields, sched_rows)

    # PROJWBS - ordered topologically with root WBS node first
    if "PROJWBS" in existing_tables:
        fields, rows = fetch_table(
            cursor,
            "PROJWBS",
            P6_FIELDS["PROJWBS"],
            "WHERE proj_id = ? ORDER BY seq_num, wbs_id",
            (proj_id,),
        )
        rows = _sort_and_sanitize_projwbs(fields, rows, proj_id_str, default_obs_id_str)
        add("PROJWBS", fields, rows)

    # RSRC - full ancestor set
    all_rsrc_ids = get_all_rsrc_ids(cursor, proj_id, existing_tables, baseline_proj_id)
    if "RSRC" in existing_tables and all_rsrc_ids:
        db_cols = get_db_columns(cursor, "RSRC")
        order_col = "rsrc_seq_num" if "rsrc_seq_num" in db_cols else "rsrc_id"
        ph = ",".join("?" * len(all_rsrc_ids))
        fields, rows = fetch_table(
            cursor,
            "RSRC",
            P6_FIELDS["RSRC"],
            f"WHERE rsrc_id IN ({ph}) ORDER BY {order_col}",
            tuple(all_rsrc_ids),
        )
        sanitized_rsrc = []
        for r in rows:
            row_copy = list(r)
            # Clear optional foreign keys to unexported dictionary tables
            _set_row_col(fields, row_copy, "role_id", "")
            _set_row_col(fields, row_copy, "shift_id", "")
            _set_row_col(fields, row_copy, "user_id", "")
            _set_row_col(fields, row_copy, "pobs_id", "")
            _set_row_col(fields, row_copy, "unit_id", "")
            _set_row_col(fields, row_copy, "location_id", "")
            sanitized_rsrc.append(row_copy)
        add("RSRC", fields, sanitized_rsrc)

    # RSRCRATE - DIRECT rsrc_ids only (not ancestors)
    if "RSRCRATE" in existing_tables and direct_rsrc_ids:
        db_cols = get_db_columns(cursor, "RSRCRATE")
        if "rsrc_id" in db_cols:
            ph = ",".join("?" * len(direct_rsrc_ids))
            order_col = "rsrc_rate_id" if "rsrc_rate_id" in db_cols else "rsrc_id"
            fields, rows = fetch_table(
                cursor,
                "RSRCRATE",
                P6_FIELDS["RSRCRATE"],
                f"WHERE rsrc_id IN ({ph}) ORDER BY {order_col}",
                tuple(direct_rsrc_ids),
            )
            sanitized_rates = []
            for r in rows:
                row_copy = list(r)
                _set_row_col(fields, row_copy, "shift_period_id", "")
                sanitized_rates.append(row_copy)
            add("RSRCRATE", fields, sanitized_rates)

    # RSRCRCAT - DIRECT rsrc_ids only
    if "RSRCRCAT" in existing_tables and direct_rsrc_ids:
        db_cols = get_db_columns(cursor, "RSRCRCAT")
        if "rsrc_id" in db_cols:
            ph = ",".join("?" * len(direct_rsrc_ids))
            fields, rows = fetch_table(
                cursor,
                "RSRCRCAT",
                P6_FIELDS["RSRCRCAT"],
                f"WHERE rsrc_id IN ({ph})",
                tuple(direct_rsrc_ids),
            )
            add("RSRCRCAT", fields, rows)

    # TASK
    exported_task_ids = set()
    task_code_to_exported_id = {}
    if "TASK" in existing_tables:
        fields, rows = fetch_table(
            cursor,
            "TASK",
            P6_FIELDS["TASK"],
            "WHERE proj_id = ? ORDER BY task_id",
            (proj_id,),
        )
        sanitized_tasks = []
        for r in rows:
            row_copy = list(r)
            tid_str = _get_row_col(fields, row_copy, "task_id")
            tcode_str = _get_row_col(fields, row_copy, "task_code")
            if tid_str:
                exported_task_ids.add(tid_str)
            if tcode_str and tid_str:
                task_code_to_exported_id[tcode_str] = tid_str

            _set_row_col(fields, row_copy, "proj_id", proj_id_str)
            _set_row_col(fields, row_copy, "location_id", "")
            if not _get_row_col(fields, row_copy, "phys_complete_pct"):
                _set_row_col(fields, row_copy, "phys_complete_pct", "0")
            if not _get_row_col(fields, row_copy, "rev_fdbk_flag"):
                _set_row_col(fields, row_copy, "rev_fdbk_flag", "N")
            if not _get_row_col(fields, row_copy, "est_wt"):
                _set_row_col(fields, row_copy, "est_wt", "1")
            if not _get_row_col(fields, row_copy, "lock_plan_flag"):
                _set_row_col(fields, row_copy, "lock_plan_flag", "N")
            if not _get_row_col(fields, row_copy, "auto_compute_act_flag"):
                _set_row_col(fields, row_copy, "auto_compute_act_flag", "N")
            if not _get_row_col(fields, row_copy, "complete_pct_type"):
                _set_row_col(fields, row_copy, "complete_pct_type", "CP_Drtn")
            if not _get_row_col(fields, row_copy, "task_type"):
                _set_row_col(fields, row_copy, "task_type", "TT_Task")
            if not _get_row_col(fields, row_copy, "duration_type"):
                _set_row_col(fields, row_copy, "duration_type", "DT_FixedDrtn")
            if not _get_row_col(fields, row_copy, "status_code"):
                _set_row_col(fields, row_copy, "status_code", "TK_NotStart")
            if not _get_row_col(fields, row_copy, "priority_type"):
                _set_row_col(fields, row_copy, "priority_type", "PT_Normal")
            if not _get_row_col(fields, row_copy, "driving_path_flag"):
                _set_row_col(fields, row_copy, "driving_path_flag", "N")
            for qty_col in (
                "act_work_qty",
                "remain_work_qty",
                "target_work_qty",
                "target_equip_qty",
                "act_equip_qty",
                "remain_equip_qty",
                "act_this_per_work_qty",
                "act_this_per_equip_qty",
            ):
                if not _get_row_col(fields, row_copy, qty_col):
                    _set_row_col(fields, row_copy, qty_col, "0")
            sanitized_tasks.append(row_copy)
        add("TASK", fields, sanitized_tasks)

    # TASKPRED
    if "TASKPRED" in existing_tables:
        fields, rows = fetch_table(
            cursor,
            "TASKPRED",
            P6_FIELDS["TASKPRED"],
            "WHERE proj_id = ? ORDER BY task_pred_id",
            (proj_id,),
        )
        sanitized_preds = []
        for r in rows:
            row_copy = list(r)
            tid_str = _get_row_col(fields, row_copy, "task_id")
            ptid_str = _get_row_col(fields, row_copy, "pred_task_id")
            if exported_task_ids and (tid_str not in exported_task_ids or ptid_str not in exported_task_ids):
                continue
            _set_row_col(fields, row_copy, "proj_id", proj_id_str)
            _set_row_col(fields, row_copy, "pred_proj_id", proj_id_str)
            if not _get_row_col(fields, row_copy, "lag_hr_cnt"):
                _set_row_col(fields, row_copy, "lag_hr_cnt", "0")
            if not _get_row_col(fields, row_copy, "pred_type"):
                _set_row_col(fields, row_copy, "pred_type", "PR_FS")
            sanitized_preds.append(row_copy)
        add("TASKPRED", fields, sanitized_preds)

    # TASKRSRC - export activity resource assignments (with fallback mapping from baseline_proj_id)
    if "TASKRSRC" in existing_tables:
        fields, rows = fetch_table(
            cursor,
            "TASKRSRC",
            P6_FIELDS["TASKRSRC"],
            "WHERE proj_id = ? ORDER BY taskrsrc_id",
            (proj_id,),
        )
        if not rows and baseline_proj_id is not None and task_code_to_exported_id:
            cursor.execute("SELECT task_id, task_code FROM TASK WHERE proj_id = ?", (baseline_proj_id,))
            base_tid_to_code = {
                format_value(get_val(r, "task_id")): format_value(get_val(r, "task_code"))
                for r in cursor.fetchall()
                if get_val(r, "task_id") is not None and get_val(r, "task_code")
            }
            fields, base_tr_rows = fetch_table(
                cursor,
                "TASKRSRC",
                P6_FIELDS["TASKRSRC"],
                "WHERE proj_id = ? ORDER BY taskrsrc_id",
                (baseline_proj_id,),
            )
            remapped_rows = []
            for r in base_tr_rows:
                row_copy = list(r)
                btid = _get_row_col(fields, row_copy, "task_id")
                bcode = base_tid_to_code.get(btid, "")
                mapped_tid = task_code_to_exported_id.get(bcode, "")
                if not mapped_tid:
                    continue
                _set_row_col(fields, row_copy, "task_id", mapped_tid)
                _set_row_col(fields, row_copy, "proj_id", proj_id_str)
                remapped_rows.append(row_copy)
            rows = remapped_rows

        if rows:
            sanitized_taskrsrc = []
            for r in rows:
                row_copy = list(r)
                tid_str = _get_row_col(fields, row_copy, "task_id")
                if exported_task_ids and tid_str not in exported_task_ids:
                    continue
                _set_row_col(fields, row_copy, "proj_id", proj_id_str)
                _set_row_col(fields, row_copy, "role_id", "")
                _set_row_col(fields, row_copy, "acct_id", "")
                _set_row_col(fields, row_copy, "pobs_id", "")
                _set_row_col(fields, row_copy, "curv_id", "")
                if not _get_row_col(fields, row_copy, "has_rsrchours"):
                    _set_row_col(fields, row_copy, "has_rsrchours", "N")
                sanitized_taskrsrc.append(row_copy)
            add("TASKRSRC", fields, sanitized_taskrsrc)

    return tables_out


def write_xer(path, db_short_name, tables_out):
    today = datetime.date.today().strftime("%Y-%m-%d")
    header = "\t".join(
        [
            "ERMHDR",
            EXPORT_VERSION,
            today,
            "Project",
            EXPORT_USER,
            db_short_name,
            EXPORT_DB_NAME,
            EXPORT_DB_TYPE,
            EXPORT_CURRENCY,
        ]
    )
    lines = [header]
    for name, fields, rows in tables_out:
        lines.append("\t".join(["%T", name]))
        lines.append("\t".join(["%F"] + fields))
        for row in rows:
            lines.append("\t".join(["%R"] + row))
    lines.append("%E")
    lines.append("")

    text = "\r\n".join(lines)
    parent_dir = os.path.dirname(path)
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir)
    with open(path, "wb") as f:
        f.write(text.encode("cp1252", errors="replace"))


def get_db_short_name(cursor):
    for table, col in [("USERS", "user_name"), ("ADMIN", "db_short_name")]:
        try:
            cursor.execute(f"SELECT {col} FROM {table} LIMIT 1")
            row = cursor.fetchone()
            if row and row[0]:
                return str(row[0]).upper()
        except Exception:
            pass
    return "ADMIN"


def parse_log_sessions(log_path):
    """Extracts chronological EOT run sessions from the process log file (most recent first)."""
    sessions, current_session = [], None
    try:
        with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                if "Reading delay events:" in line:
                    parts = line.split("Reading delay events:")
                    current_session = {
                        "timestamp": line[:19],
                        "excel": os.path.basename(parts[1].strip()) if len(parts) > 1 else "Unknown",
                        "projects": [],
                    }
                    sessions.append(current_session)
                elif current_session is not None:
                    name = None
                    if "Project Copy Code:" in line:
                        parts = line.split("Project Copy Code:")
                        name = parts[1].strip() if len(parts) > 1 else ""
                    elif "->" in line and "Replicating baseline project:" in line:
                        parts = line.split("->")
                        name = parts[1].strip().split("...")[0].strip() if len(parts) > 1 else ""
                    elif "->" in line and "Replicating project:" in line:
                        parts = line.split("->")
                        name = parts[1].strip().split("...")[0].strip() if len(parts) > 1 else ""
                    if name and name not in current_session["projects"]:
                        current_session["projects"].append(name)
    except Exception as e:
        print(f"[WARNING] Could not parse log file: {e}")

    sessions = [s for s in sessions if s["projects"]]
    sessions.reverse()
    return sessions


def _get_delay_wbs_proj_ids(cursor, existing_tables):
    """Returns project IDs that contain a 'Delay Events' WBS node created by the EOT toolkit."""
    if "PROJWBS" not in existing_tables:
        return set()
    try:
        cursor.execute("SELECT DISTINCT PROJ_ID FROM PROJWBS WHERE WBS_NAME = 'Delay Events'")
        return {get_val(r, "PROJ_ID") for r in cursor.fetchall() if get_val(r, "PROJ_ID") is not None}
    except Exception:
        return set()


def main():
    global conn

    parser = argparse.ArgumentParser(description="P6 Project -> .XER Exporter")
    parser.add_argument("--db-path", "--db", help="Path to Primavera SQLite DB")
    parser.add_argument("--project", "-p", help="Baseline project short name")
    parser.add_argument("--mode", "-m", choices=["1", "2", "3"], help="Export mode (1, 2, or 3)")
    parser.add_argument("--log-path", "--log", help="Path to EOT process log file")
    parser.add_argument("--serials", "-s", help="Comma-separated serial numbers of projects to export")
    parser.add_argument("--out-dir", "-o", help="Output directory for XER files")
    parser.add_argument("--yes", "-y", action="store_true", help="Automatically confirm export prompt")
    args, _ = parser.parse_known_args()

    # DB Path
    db_path = args.db_path
    if not db_path:
        try:
            db_path = (
                input(f"Enter the Primavera SQLite database path (default: {DEFAULT_DB_PATH}): ").strip().strip("'\"")
            )
            if not db_path:
                db_path = DEFAULT_DB_PATH
            while not os.path.exists(db_path):
                db_path = input(f"[ERROR] File not found at '{db_path}'. Enter a valid path: ").strip().strip("'\"")
                if not db_path:
                    db_path = DEFAULT_DB_PATH
        except KeyboardInterrupt:
            print("\nCancelled.")
            sys.exit(0)
    else:
        db_path = db_path.strip("'\"")
        if not os.path.exists(db_path):
            print(f"[ERROR] Database file not found at '{db_path}'.")
            safe_exit(1, f"Database file not found at '{db_path}'.")

    print("=" * 72)
    print("  PRIMAVERA P6 EOT PROJECT EXPORT UTILITY")
    print("=" * 72)
    print(f"Connecting to database: {db_path}...")
    conn = get_connection(db_path)
    cur = conn.cursor()
    existing_tables = get_existing_tables(cur)
    db_short_name = get_db_short_name(cur)

    # Baseline Project ID
    source_project = args.project
    if not source_project:
        source_project = input("\nEnter the baseline project ID (PROJ_SHORT_NAME): ").strip()
        while not source_project:
            source_project = input("Project ID cannot be empty: ").strip()
    else:
        source_project = source_project.strip()

    cur.execute("SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'", (source_project,))
    base_row = cur.fetchone()
    resolved_source_proj_id = get_val(base_row, "PROJ_ID") if base_row else None

    # Choice Mode
    choice = args.mode
    if not choice:
        print("\nPlease select an export mode:")
        print("  1. Export delay events scheduled in last run only [DEFAULT]")
        print("  2. Load all delay events scheduled in a project (select by serial number)")
        print("  3. Export all projects in database")
        choice = input("\nEnter choice (1, 2, or 3): ").strip() or "1"

    projects_to_export = []
    delay_wbs_proj_ids = _get_delay_wbs_proj_ids(cur, existing_tables)

    if choice == "1":
        log_path = args.log_path
        if not log_path:
            try:
                print("Enter the log file path (press Enter to browse):")
                log_path = input("Log File Path: ").strip().strip("'\"")
                if not log_path:
                    if os.path.exists(DEFAULT_LOG_PATH) and (
                        os.environ.get("EOT_AUTO_NO_POPUP") == "1" or os.environ.get("EOT_BATCH_MODE") == "1"
                    ):
                        log_path = DEFAULT_LOG_PATH
                    elif os.environ.get("EOT_AUTO_NO_POPUP") != "1":
                        try:
                            import tkinter as tk
                            from tkinter import filedialog

                            root = tk.Tk()
                            root.withdraw()
                            root.attributes("-topmost", True)
                            log_path = filedialog.askopenfilename(
                                title="Select EOT Process Log File",
                                filetypes=[("Log files", "*.log"), ("All files", "*.*")],
                            )
                            root.destroy()
                            log_path = os.path.abspath(log_path) if log_path else ""
                        except Exception:
                            log_path = DEFAULT_LOG_PATH if os.path.exists(DEFAULT_LOG_PATH) else ""
                    else:
                        log_path = ""
            except KeyboardInterrupt:
                print("\nCancelled.")
                sys.exit(0)
        else:
            log_path = log_path.strip("'\"")

        if not log_path or not os.path.exists(log_path):
            print("\n[ERROR] Option 1 requires a valid, existing log file.")
            safe_exit(1, "Option 1 requires a valid, existing log file.")

        sessions = parse_log_sessions(log_path)
        if not sessions:
            print("\nNo EOT Run Sessions found in log history.")
            safe_exit(0)

        selected = sessions[0]
        print(f"\nAuto-targeted last EOT run from {selected['timestamp']} with {len(selected['projects'])} project(s).")
        for name in selected["projects"]:
            cur.execute("SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'", (name,))
            row = cur.fetchone()
            if row:
                projects_to_export.append((get_val(row, "PROJ_ID"), name))

    elif choice == "2":
        if not resolved_source_proj_id:
            print(f"[ERROR] Baseline project '{source_project}' not found.")
            safe_exit(1, f"Baseline project '{source_project}' not found.")
        source_proj_id = resolved_source_proj_id

        cur.execute("SELECT TASK_CODE FROM TASK WHERE PROJ_ID = ?", (source_proj_id,))
        source_task_codes = {get_val(r, "TASK_CODE") for r in cur.fetchall() if get_val(r, "TASK_CODE")}

        cur.execute(
            "SELECT PROJ_ID, PROJ_SHORT_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y' AND PROJ_SHORT_NAME != ?",
            (source_project,),
        )
        candidate_projects = []
        for r in cur.fetchall():
            pid, name = get_val(r, "PROJ_ID"), get_val(r, "PROJ_SHORT_NAME")
            is_match = source_project in name if name else False
            if not is_match and name and (DELAY_PROJ_REGEX.match(name) or pid in delay_wbs_proj_ids):
                cur.execute("SELECT TASK_CODE FROM TASK WHERE PROJ_ID = ? AND TASK_CODE != ?", (pid, name))
                if {get_val(rr, "TASK_CODE") for rr in cur.fetchall() if get_val(rr, "TASK_CODE")} & source_task_codes:
                    is_match = True
            if is_match:
                candidate_projects.append((pid, name))

        if not candidate_projects:
            print(f"\nNo delay/EOT projects found associated with '{source_project}'.")
            safe_exit(0)

        print(f"\nFound {len(candidate_projects)} project(s):")
        for idx, (pid, name) in enumerate(candidate_projects, 1):
            print(f"  {idx}. {name}")

        selection = args.serials
        if not selection:
            selection = input(
                "\nEnter the serial numbers of the projects to export (comma-separated, e.g. 1,3,4): "
            ).strip()
        selected_indices = [
            int(v.strip()) - 1
            for v in selection.split(",")
            if v.strip().isdigit() and 0 <= int(v.strip()) - 1 < len(candidate_projects)
        ]
        if not selected_indices:
            print("\nNo valid selections made.")
            safe_exit(0)
        projects_to_export = [candidate_projects[i] for i in selected_indices]

    elif choice == "3":
        cur.execute("SELECT PROJ_ID, PROJ_SHORT_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y'")
        for r in cur.fetchall():
            name, pid = get_val(r, "PROJ_SHORT_NAME"), get_val(r, "PROJ_ID")
            if name != source_project and (is_delay_project_name(name) or pid in delay_wbs_proj_ids):
                projects_to_export.append((pid, name))
    else:
        print("\nInvalid choice.")
        safe_exit(1, "Invalid choice.")

    if not projects_to_export:
        print("\nNo matching projects found to export.")
        safe_exit(0)

    print(f"\nFound {len(projects_to_export)} project(s) to export:")
    for idx, (_, name) in enumerate(projects_to_export, 1):
        print(f"  {idx}. {name}")

    if not args.yes:
        confirm = input("\nDo you want to export these projects to XER format? (y/n): ").strip().lower()
        if confirm not in ("y", "yes"):
            print("\nExport cancelled.")
            safe_exit(0)

    # Output directory handling
    out_dir = args.out_dir
    if not out_dir:
        try:
            print("\nEnter output directory for XER files (press Enter for default):")
            out_dir = input("Output Directory: ").strip().strip("'\"")

            if not out_dir:
                if os.environ.get("EOT_AUTO_NO_POPUP") == "1" or os.environ.get("EOT_BATCH_MODE") == "1":
                    out_dir = os.path.join(_SCRIPT_DIR, "exports")
                else:
                    try:
                        import tkinter as tk
                        from tkinter import filedialog

                        root = tk.Tk()
                        root.withdraw()
                        root.attributes("-topmost", True)
                        out_dir = filedialog.askdirectory(title="Select Output Directory for XER Files")
                        root.destroy()
                    except Exception:
                        out_dir = os.path.join(_SCRIPT_DIR, "exports")
        except Exception as e:
            print(f"[WARNING] Could not determine output directory: {e}")
            out_dir = None

    if not out_dir:
        out_dir = os.path.join(_SCRIPT_DIR, "exports")
        print(f"No folder selected. Using default: {out_dir}")

    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    print(f"\nExporting to: {out_dir}")
    export_failed = False
    for proj_id, name in projects_to_export:
        print(f"  Exporting: {name} (ID: {proj_id})...")
        try:
            tables_out = build_export(cur, proj_id, existing_tables, source_proj_id=resolved_source_proj_id)
            safe_fname = re.sub(r'[<>:"/\\|?*]', "_", name)
            out_path = os.path.join(out_dir, f"{safe_fname}_export.xer")
            write_xer(out_path, db_short_name, tables_out)
            print(f"    [OK] {os.path.basename(out_path)}")
        except Exception as e:
            print(f"    [ERROR] Failed to export {name}: {e}")
            import traceback

            traceback.print_exc()
            export_failed = True

    if export_failed:
        print("\nExport completed with some errors.")
        safe_exit(1, "Some projects failed to export.")
    else:
        print("\nExport completed successfully!")
        safe_exit(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nCancelled.")
