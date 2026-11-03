#!/usr/bin/env python3
"""
Primavera P6 EOT Project Deletion Utility (SQLite)
==================================================
Safely purges simulated EOT/delay project copies and their metadata from
TASKRSRC, TASKPRED, TASK, PROJWBS, SCHEDOPTIONS, SUMTASK, CALENDAR (CA_Proj),
PROJSHAR, PROJPROP, and PROJECT tables.
Supports both interactive stdin prompts (used by EOT_GUI.py) and CLI arguments.
"""

import argparse
import os
import re
import sqlite3
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB_PATH = r"C:\Users\91943\Downloads\T3.db"
DEFAULT_LOG_PATH = os.path.join(_SCRIPT_DIR, "eot_process_logs.log")

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
    global conn
    if conn:
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
                messagebox.showinfo("Process Status", "The process is completed successfully.")
            root.destroy()

            input("\nPress Enter to exit...")
        except Exception:
            pass
        sys.exit(exit_code)


def parse_log_sessions(log_path):
    """Parses EOT Run Sessions from the log file (most recent first)."""
    sessions = []
    current_session = None
    try:
        with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                if "Reading delay events:" in line:
                    timestamp = line[:19]
                    parts = line.split("Reading delay events:")
                    excel_path = parts[1].strip() if len(parts) > 1 else "Unknown"

                    current_session = {
                        "timestamp": timestamp,
                        "excel": os.path.basename(excel_path),
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


def delete_project_from_db(cur, proj_id):
    """Deletes a project and all its associated records across P6 tables."""
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


def main():
    global conn

    parser = argparse.ArgumentParser(description="Primavera P6 EOT Project Deletion Utility")
    parser.add_argument("--db-path", "--db", help="Path to Primavera SQLite DB")
    parser.add_argument("--project", "-p", help="Baseline project short name")
    parser.add_argument("--mode", "-m", choices=["1", "2", "3"], help="Deletion mode (1, 2, or 3)")
    parser.add_argument("--log-path", "--log", help="Path to EOT process log file")
    parser.add_argument("--serials", "-s", help="Comma-separated serial numbers of projects to delete")
    parser.add_argument("--yes", "-y", action="store_true", help="Automatically confirm deletion prompt")
    args, _ = parser.parse_known_args()

    db_path = args.db_path
    if not db_path:
        try:
            db_path = (
                input(f"Enter the Primavera SQLite database path (default: {DEFAULT_DB_PATH}): ").strip().strip("'\"")
            )
            if not db_path:
                db_path = DEFAULT_DB_PATH
            while not os.path.exists(db_path):
                db_path = (
                    input(f"[ERROR] Database file not found at '{db_path}'. Please enter a valid path: ")
                    .strip()
                    .strip("'\"")
                )
                if not db_path:
                    db_path = DEFAULT_DB_PATH
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            sys.exit(0)
    else:
        db_path = db_path.strip("'\"")
        if not os.path.exists(db_path):
            print(f"[ERROR] Database file not found at '{db_path}'.")
            safe_exit(1, f"Database file not found at '{db_path}'.")

    print("=" * 72)
    print("  PRIMAVERA P6 EOT PROJECT DELETION UTILITY")
    print("=" * 72)

    print(f"Connecting to database: {db_path}...")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Prompt for baseline project name/code first
    source_project = args.project
    if not source_project:
        source_project = input("\nEnter the baseline project ID (PROJ_SHORT_NAME): ").strip()
        while not source_project:
            source_project = input("Project ID cannot be empty. Please enter the baseline project ID: ").strip()
    else:
        source_project = source_project.strip()

    choice = args.mode
    if not choice:
        print("\nPlease select a deletion mode:")
        print("  1. Delete delay events scheduled in last run only [DEFAULT]")
        print("  2. Load all delay events scheduled in a project (select by serial number)")
        print("  3. Scan & delete ALL delay projects (containing '_DE', '_VE', or '_EOT' in name)")

        choice = input("\nEnter choice (1, 2, or 3): ").strip()
        if not choice:
            choice = "1"

    existing_projects = []

    if choice == "1":
        log_path = args.log_path
        if not log_path:
            try:
                print("Enter the log file path (press Enter to browse / open file picker, or paste path):")
                log_path = input("Log File Path: ").strip().strip("'\"")
                if not log_path:
                    if os.path.exists(DEFAULT_LOG_PATH) and (
                        os.environ.get("EOT_AUTO_NO_POPUP") == "1" or os.environ.get("EOT_BATCH_MODE") == "1"
                    ):
                        log_path = DEFAULT_LOG_PATH
                    else:
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
            except KeyboardInterrupt:
                print("\nOperation cancelled by user.")
                sys.exit(0)
        else:
            log_path = log_path.strip("'\"")

        if not log_path or not os.path.exists(log_path):
            print("\n[ERROR] Option 1 (Delete last EOT run only) requires a valid, existing log file.")
            safe_exit(1, "Option 1 (Delete last EOT run only) requires a valid, existing log file.")

        sessions = parse_log_sessions(log_path)

        if not sessions:
            print("\nNo EOT Run Sessions found in the log history.")
            safe_exit(0)

        selected_session = sessions[0]
        target_names = selected_session["projects"]

        print(
            f"\nAuto-targeting the last EOT run from {selected_session['timestamp']} with {len(target_names)} project(s)."
        )

        for name in target_names:
            cur.execute("SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'", (name,))
            row = cur.fetchone()
            if row:
                existing_projects.append((row[0], name))

        if not existing_projects:
            print("\nNone of the projects from that log session currently exist in the database.")
            safe_exit(0)

    elif choice == "2":
        cur.execute("SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'", (source_project,))
        row = cur.fetchone()
        if not row:
            print(f"[ERROR] Baseline project '{source_project}' not found in database.")
            safe_exit(1, f"Baseline project '{source_project}' not found in database.")
        source_proj_id = row[0]

        cur.execute("SELECT TASK_CODE FROM TASK WHERE PROJ_ID = ?", (source_proj_id,))
        source_task_codes = {r[0] for r in cur.fetchall() if r[0]}

        try:
            cur.execute("SELECT DISTINCT PROJ_ID FROM PROJWBS WHERE WBS_NAME = 'Delay Events'")
            delay_wbs_proj_ids = {r[0] for r in cur.fetchall() if r[0] is not None}
        except Exception:
            delay_wbs_proj_ids = set()

        cur.execute(
            "SELECT PROJ_ID, PROJ_SHORT_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y' AND PROJ_SHORT_NAME != ?",
            (source_project,),
        )
        all_other_projs = cur.fetchall()

        candidate_projects = []
        for pid, name in all_other_projs:
            is_match = False
            if name and source_project in name:
                is_match = True
            elif name and (DELAY_PROJ_REGEX.match(name) or pid in delay_wbs_proj_ids):
                cur.execute("SELECT TASK_CODE FROM TASK WHERE PROJ_ID = ? AND TASK_CODE != ?", (pid, name))
                proj_tasks = {r[0] for r in cur.fetchall() if r[0]}
                if proj_tasks.intersection(source_task_codes):
                    is_match = True

            if is_match:
                candidate_projects.append((pid, name))

        if not candidate_projects:
            print(f"\nNo delay/EOT projects found associated with '{source_project}'.")
            safe_exit(0)

        print(f"\nFound {len(candidate_projects)} delay project(s) associated with '{source_project}':")
        for idx, (pid, name) in enumerate(candidate_projects, 1):
            print(f"  {idx}. {name}")

        selection_str = args.serials
        if not selection_str:
            selection_str = input(
                "\nEnter the serial numbers of the projects to delete (comma-separated, e.g. 1,3,4): "
            ).strip()
        if not selection_str:
            print("\nNo selection made. Operation cancelled.")
            safe_exit(0)

        selected_indices = []
        for part in selection_str.split(","):
            part = part.strip()
            if part.isdigit():
                idx = int(part) - 1
                if 0 <= idx < len(candidate_projects):
                    selected_indices.append(idx)
                else:
                    print(f"[WARNING] Serial number {part} is out of range. Ignoring.")
            else:
                print(f"[WARNING] Invalid input '{part}'. Ignoring.")

        selected_indices = sorted(list(set(selected_indices)))
        if not selected_indices:
            print("\nNo valid selections made. Operation cancelled.")
            safe_exit(0)

        existing_projects = [candidate_projects[i] for i in selected_indices]

    elif choice == "3":
        try:
            cur.execute("SELECT DISTINCT PROJ_ID FROM PROJWBS WHERE WBS_NAME = 'Delay Events'")
            delay_wbs_proj_ids = {r[0] for r in cur.fetchall() if r[0] is not None}
        except Exception:
            delay_wbs_proj_ids = set()

        cur.execute("SELECT PROJ_ID, PROJ_SHORT_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y'")
        all_projs = cur.fetchall()

        for pid, name in all_projs:
            if name != source_project and (is_delay_project_name(name) or pid in delay_wbs_proj_ids):
                existing_projects.append((pid, name))

        if not existing_projects:
            print("\nNo simulated EOT projects found in the database.")
            safe_exit(0)
    else:
        print("\nInvalid choice. Exiting.")
        safe_exit(1, "Invalid choice.")

    print(f"\nFound {len(existing_projects)} project(s) to delete:")
    for idx, (pid, name) in enumerate(existing_projects, 1):
        print(f"  {idx}. {name}")

    if not args.yes:
        confirm = input("\nWARNING: Are you sure you want to PERMANENTLY delete these projects? (y/n): ").strip().lower()
        if confirm not in ("y", "yes"):
            print("\nDeletion cancelled.")
            safe_exit(0)

    print("\nDeleting projects from database...")
    try:
        for pid, name in existing_projects:
            print(f"  Deleting: {name}...")
            delete_project_from_db(cur, pid)

        conn.commit()
        print("\nAll selected projects and their associated tasks/WBS have been successfully deleted!")
    except Exception as e:
        conn.rollback()
        print(f"\n[ERROR] Error occurred while deleting projects: {e}")
        safe_exit(1, str(e))
    finally:
        safe_exit(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(0)
