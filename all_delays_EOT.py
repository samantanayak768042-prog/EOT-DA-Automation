"""
===============================================================================
  PRIMAVERA P6 EOT (Extension of Time) AUTOMATION – PURE SQLITE VERSION
Copy project->rename it by "delay_<project_name>" ->create wbs in it->add all delay events there->establish relationship
===============================================================================
This script performs Combined Extension of Time (EOT) analysis for Primavera P6
by manipulating the SQLite database directly. It replicates the baseline project
once, inserts all delay events into a single project copy, maps predecessor
relationships, and calculates the cumulative early/late start and finish dates
using the shared Critical Path Method (CPM) engine in db_eot_automation.py.

Outputs a single consolidated Excel report displaying the impacted dates for all
specified activities imported from an Excel file (or matching the standard
milestone pattern).

NO GUI INVOLVEMENT: Runs completely in background without opening Primavera P6.

PREREQUISITES
──────────────
    pip install openpyxl pandas
===============================================================================
"""

import os
import re
import sqlite3
import sys
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# Reuse core SQLite, Calendar, CPM, and Reporting primitives from db_eot_automation
from db_eot_automation import (
    DEFAULT_DB_PATH,
    DELAY_COLUMN_ALIASES,
    LOG_FILE,
    MS_PATTERN,
    OUTPUT_FILE,
    SpecifiedActivityList,
    add_delay_event_to_project,
    add_workdays,
    add_workdays_cal,
    allocate_keys,
    build_milestone_workbook,
    cleanup_project_tables,
    configure_logging,
    duplicate_project_in_db,
    err,
    extract_delay_ids_from_excel,
    format_date_ms,
    generate_p6_guid,
    get_next_workday_start_cal,
    get_project_finish_date_from_db,
    get_unique_output_path,
    insert_row,
    is_workday,
    load_and_resolve_calendars,
    load_ms_pattern_from_excel,
    log,
    map_constraint_type,
    parse_args,
    parse_p6_datetime,
    prompt_save_path,
    resolve_delay_dataframe,
    resolve_milestone_pattern_interactive,
    retrieve_project_milestones,
    run_cpm_scheduling,
    show_popup,
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# COMBINED MILESTONE VS DELAY REPORT GENERATOR
# ─────────────────────────────────────────────────────────────────────────────


def run_milestone_vs_delay(db_path, excel_path, source_project, ms_pattern=None, output_path=None):
    print("=" * 72)
    print("  PRIMAVERA P6 MILESTONE VS DELAY REPORT GENERATOR (COMBINED)")
    print("=" * 72)

    db_path = db_path.strip("'\"")
    excel_path = excel_path.strip("'\"")

    if not os.path.exists(db_path):
        print(f"[ERROR] Primavera SQLite database not found at: {db_path}")
        return

    # 1. Connect to SQLite and fetch project IDs
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

        # 2. Retrieve specified activities or milestones from baseline project
        milestones = retrieve_project_milestones(cur, baseline_proj_id, ms_pattern)
        print(f"Found {len(milestones)} item(s) to report on from baseline project.")

        # 3. Extract dates for the single combined delay project copy
        delay_dates = {}
        copy_name = f"delay_{source_project}"
        cur.execute(
            "SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
            (copy_name,),
        )
        proj_row = cur.fetchone()
        if not proj_row:
            print(
                f"[WARNING] Combined project copy '{copy_name}' not found in database. Cannot generate Milestone vs Delay report."
            )
            return
        d_proj_id = proj_row[0]

        cur.execute("SELECT TASK_CODE, EARLY_END_DATE FROM TASK WHERE PROJ_ID = ?", (d_proj_id,))
        task_dates = dict(cur.fetchall())
        delay_dates[copy_name] = task_dates

    except Exception as e:
        print(f"[ERROR] Database query failed: {e}")
        return
    finally:
        conn.close()

    # 4. Build the styled Excel report using openpyxl
    print("\nGenerating consolidated report...")
    wb = build_milestone_workbook(
        db_path,
        source_project,
        milestones,
        delay_dates,
        "P6 Milestone Finish Dates vs Delay Event Impacts (Combined)",
    )

    if not output_path:
        excel_dir = os.path.dirname(os.path.abspath(excel_path))
        output_path = get_unique_output_path(excel_dir)

    try:
        wb.save(output_path)
        print("\nSuccess! Consolidated Milestone vs Delay report generated at:")
        print(f"   -> {output_path}")
        print("=" * 72)
    except PermissionError:
        print(f"\n[ERROR] Permission denied when writing to {output_path}.")
        print("Please close the Excel file in Excel and run the script again.")
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# COMBINED EOT CORE PROCESSING
# ─────────────────────────────────────────────────────────────────────────────


def process_combined_delay_events(
    excel_path,
    db_path,
    source_project,
    default_remarks="",
    output_excel=None,
    logger=None,
    error_logger=None,
):
    """
    Replicates the baseline project once (as 'delay_<source_project>'), inserts
    all valid delay events from the Excel file into that single copy, runs CPM
    scheduling, and returns (results, skipped_rows).

    Note: The consolidated Excel output is generated by run_milestone_vs_delay().
    If output_excel is explicitly passed by an external caller, df_out is also
    written there for backward compatibility.
    """
    _log = logger or log
    _err = error_logger or err

    df_delays, resolved_cols, best_idx = resolve_delay_dataframe(
        excel_path, logger=_log, error_logger=_err
    )

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

    copy_name = f"delay_{source_project}"
    results = []
    skipped_rows = []

    conn = sqlite3.connect(db_path)
    try:
        print("\n" + "-" * 72)
        _log(f"Replicating baseline project: {source_project} -> {copy_name}")
        _log(f"Project Copy Code: {copy_name}")
        print("-" * 72)
        new_proj_id, delay_wbs_id = duplicate_project_in_db(conn, source_project, copy_name)
        conn.commit()

        _log(f"\nInserting delay events into project {copy_name}...")
        for idx, row in df_delays.iterrows():
            delay_id = str(row[resolved_cols["Delay Event ID"]]).strip()
            delay_name = str(row[resolved_cols["Description of Event"]]).strip()
            excel_row_num = idx + (best_idx + 3 if best_idx != -1 else 2)
            try:
                add_delay_event_to_project(
                    conn, new_proj_id, delay_wbs_id, row, resolved_cols, source_proj_id, resolved_calendars
                )
                conn.commit()
            except Exception as exc:
                conn.rollback()
                _err(f"Discrepancy / error at Excel row {excel_row_num} (Delay ID: {delay_id}): {exc}")
                skipped_rows.append(
                    {
                        "excel_row": excel_row_num,
                        "delay_id": delay_id,
                        "delay_name": delay_name,
                        "error": str(exc),
                    }
                )

        # Execute mathematical CPM scheduling on the copy project with all delays
        _log(f"\n[CPM] Running critical path calculation on {copy_name}...")
        run_cpm_scheduling(conn, new_proj_id, resolved_calendars)
        conn.commit()

        finish_date = get_project_finish_date_from_db(db_path, new_proj_id)
        _log(f"  Calculated combined project finish date: {finish_date}")

        skipped_ids_rows = {(s["excel_row"], s["delay_id"]) for s in skipped_rows}
        for idx, row in df_delays.iterrows():
            delay_id = str(row[resolved_cols["Delay Event ID"]]).strip()
            excel_row_num = idx + (best_idx + 3 if best_idx != -1 else 2)
            if (excel_row_num, delay_id) in skipped_ids_rows:
                continue
            lag_col = resolved_cols.get("Lag")
            lag_val = (
                float(row[lag_col])
                if (lag_col and lag_col in row and pd.notna(row[lag_col]) and str(row[lag_col]).strip() != "")
                else 0.0
            )
            results.append(
                {
                    "Delay Event ID": delay_id,
                    "Delay Event Name": str(row[resolved_cols["Description of Event"]]).strip(),
                    "Impacting Activity": str(row[resolved_cols["Activity Impacted"]]).strip(),
                    "Relationship Type": str(row[resolved_cols["Relationship"]]).strip(),
                    "Lag (days)": lag_val,
                    "New Finish Date": finish_date,
                    "Observation": default_remarks if default_remarks else "",
                }
            )

        _log("\nAll database changes successfully committed.")
    except Exception as exc:
        conn.rollback()
        _err(f"Database process failed: {exc}")
        raise
    finally:
        conn.close()

    if results and output_excel:
        df_out = pd.DataFrame(results)
        df_out.to_excel(output_excel, index=False)
        _log(f"\nResults written to -> {output_excel}")
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
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────


def main():
    args = parse_args()
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
        results, skipped_rows = process_combined_delay_events(
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
    except Exception as e:
        show_popup(success=False)
        raise e
