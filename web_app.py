"""
web_app.py - Enterprise Web Application for Primavera P6 EOT Automation Toolkit
Built with FastAPI, WebSocket Real-time Log Streaming, and Native Process Isolation.
"""

import asyncio
from collections import deque
from contextlib import asynccontextmanager
from datetime import datetime
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
from typing import Dict, List, Optional, Set
import uuid
import zipfile

from fastapi import (
    FastAPI,
    File,
    Form,
    HTTPException,
    Query,
    Request,
    UploadFile,
    WebSocket,
    WebSocketDisconnect,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

# Base Paths
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(_SCRIPT_DIR, "uploads")
OUTPUTS_DIR = os.path.join(_SCRIPT_DIR, "outputs")
STATIC_DIR = os.path.join(_SCRIPT_DIR, "static")
TEMPLATES_DIR = os.path.join(_SCRIPT_DIR, "templates")
ASSETS_DIR = os.path.join(_SCRIPT_DIR, "assets")

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUTS_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)

DEFAULT_DB_PATH = ""
DEFAULT_LOG_PATH = os.path.join(_SCRIPT_DIR, "eot_process_logs.log")
TEMPLATE_EXCEL_PATH = os.path.join(_SCRIPT_DIR, "Delay_Events_template.xlsx")

# Script Paths
DB_EOT_SCRIPT = os.path.join(_SCRIPT_DIR, "db_eot_automation.py")
ALL_DELAYS_SCRIPT = os.path.join(_SCRIPT_DIR, "all_delays_EOT.py")
EXPORT_EOT_SCRIPT = os.path.join(_SCRIPT_DIR, "export_EOT.py")
DELETE_PROJECTS_SCRIPT = os.path.join(_SCRIPT_DIR, "delete_projects.py")

DELAY_PROJ_REGEX = re.compile(r"^(DE|VE|DC)\d+$", re.IGNORECASE)

# ─────────────────────────────────────────────────────────────────────────────
# FASTAPI APP INITIALIZATION
# ─────────────────────────────────────────────────────────────────────────────

# Global Event Loop Reference for WebSocket Threadsafe notifications
_MAIN_LOOP: Optional[asyncio.AbstractEventLoop] = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    global _MAIN_LOOP
    _MAIN_LOOP = asyncio.get_running_loop()
    yield


app = FastAPI(
    title="Primavera P6 EOT Automation Toolkit",
    description="Web Application for Critical Path Schedule Delay Simulations and XER Packaging",
    version="6.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")

templates = Jinja2Templates(directory=TEMPLATES_DIR)


# ─────────────────────────────────────────────────────────────────────────────
# TASK & PROCESS MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────


class Task:
    def __init__(self, task_id: str, task_type: str, cmd: List[str], work_dir: str):
        self.task_id = task_id
        self.task_type = task_type
        self.cmd = cmd
        self.work_dir = work_dir
        self.status = "queued"  # queued, running, completed, failed, cancelled
        self.exit_code: Optional[int] = None
        self.logs: deque = deque(maxlen=8000)
        self.discrepancies: Optional[str] = None
        self.output_files: List[str] = []
        self.created_at = datetime.now().isoformat()
        self.completed_at: Optional[str] = None
        self.process: Optional[subprocess.Popen] = None
        self.clients: Set[WebSocket] = set()
        self.lock = threading.Lock()
        self.auto_responses: Dict[str, str] = {}

    def add_log(self, text: str):
        with self.lock:
            self.logs.append(text)
        self.broadcast_log(text)

    def broadcast_log(self, text: str):
        if not _MAIN_LOOP or not self.clients:
            return
        msg = json.dumps({"type": "log", "data": text})
        for ws in list(self.clients):
            try:
                asyncio.run_coroutine_threadsafe(ws.send_text(msg), _MAIN_LOOP)
            except Exception:
                pass

    def broadcast_event(self, event_type: str, payload: dict):
        if not _MAIN_LOOP or not self.clients:
            return
        msg = json.dumps({"type": event_type, "payload": payload})
        for ws in list(self.clients):
            try:
                asyncio.run_coroutine_threadsafe(ws.send_text(msg), _MAIN_LOOP)
            except Exception:
                pass


TASKS: Dict[str, Task] = {}


def extract_discrepancies_from_text(log_text: str) -> Optional[str]:
    pattern = r"SUMMARY OF SKIPPED ROWS WITH DISCREPANCIES\s*={5,}\s*(.*?)\s*={5,}"
    match = re.search(pattern, log_text, re.DOTALL)
    if match:
        content = match.group(1).strip()
        if "Total rows skipped: 0" in content or "No rows were skipped" in content:
            return None
        return content
    return None


def run_task_worker(task: Task):
    task.status = "running"
    task.broadcast_event("status_change", {"status": "running"})

    env = os.environ.copy()
    env["EOT_AUTO_NO_POPUP"] = "1"
    env["PYTHONUNBUFFERED"] = "1"

    try:
        task.process = subprocess.Popen(
            task.cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            cwd=task.work_dir,
            bufsize=1,
        )

        buffer = ""
        full_log_accum = []

        while True:
            char = task.process.stdout.read(1)
            if not char:
                if task.process.poll() is not None:
                    break
                time.sleep(0.01)
                continue

            task.add_log(char)
            buffer += char
            full_log_accum.append(char)

            # Check auto responses
            for pattern, response in list(task.auto_responses.items()):
                if pattern in buffer:
                    try:
                        task.process.stdin.write(response + "\n")
                        task.process.stdin.flush()
                        task.add_log(f"\n> {response}\n")
                    except Exception:
                        pass
                    del task.auto_responses[pattern]
                    buffer = ""
                    break

            if len(buffer) > 500:
                buffer = buffer[-500:]

        rc = task.process.wait()
        task.exit_code = rc
        task.completed_at = datetime.now().isoformat()

        full_log_text = "".join(full_log_accum)
        discrepancies = extract_discrepancies_from_text(full_log_text)
        if discrepancies:
            task.discrepancies = discrepancies
            task.broadcast_event("discrepancies", {"discrepancies": discrepancies})

        # Scan for output files
        found_outputs = []
        for root, _, files in os.walk(task.work_dir):
            for file in files:
                if file.endswith((".xlsx", ".xer", ".log")):
                    found_outputs.append(os.path.join(root, file))

        # Check explicitly referenced output paths in log
        out_match = re.search(r"->\s*([^\r\n]+\.xlsx)", full_log_text)
        if out_match:
            candidate = out_match.group(1).strip()
            if os.path.exists(candidate) and candidate not in found_outputs:
                found_outputs.append(candidate)

        task.output_files = found_outputs

        if rc == 0:
            task.status = "completed"
            task.add_log(f"\n\n[Process completed successfully with exit code {rc}]")
        elif task.status == "cancelled":
            task.add_log("\n\n[Process was cancelled by user]")
        else:
            task.status = "failed"
            task.add_log(f"\n\n[Process terminated with error code {rc}]")

    except Exception as e:
        task.status = "failed"
        task.exit_code = -1
        task.add_log(f"\n\n[Subprocess Execution Error: {e}]")
    finally:
        if task.process:
            try:
                task.process.stdout.close()
            except Exception:
                pass
            try:
                task.process.stdin.close()
            except Exception:
                pass
        task.broadcast_event(
            "status_change",
            {
                "status": task.status,
                "exit_code": task.exit_code,
                "discrepancies": bool(task.discrepancies),
                "outputs_count": len(task.output_files),
            },
        )


# ─────────────────────────────────────────────────────────────────────────────
# DATABASE UTILITIES & HELPERS
# ─────────────────────────────────────────────────────────────────────────────


def is_delay_project_name(name: str) -> bool:
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


def parse_log_sessions(log_path: str) -> List[dict]:
    sessions = []
    current_session = None
    if not log_path or not os.path.exists(log_path):
        return sessions

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
    except Exception:
        pass
    return sessions


def inspect_database_details(db_path: str) -> dict:
    if not db_path or not os.path.exists(db_path):
        return {"connected": False, "error": "Database file not found"}

    try:
        conn = sqlite3.connect(f"file:{os.path.abspath(db_path)}?mode=ro", uri=True)
        cur = conn.cursor()
        cur.execute(
            "SELECT PROJ_ID, PROJ_SHORT_NAME, PROJ_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y' ORDER BY PROJ_SHORT_NAME"
        )
        rows = cur.fetchall()
        conn.close()

        projects = [{"id": r[0], "code": r[1], "name": r[2] or r[1]} for r in rows]
        return {
            "connected": True,
            "db_path": os.path.abspath(db_path),
            "db_name": os.path.basename(db_path),
            "project_count": len(projects),
            "projects": projects,
        }
    except Exception as e:
        return {"connected": False, "error": str(e)}


def scan_candidate_projects(
    db_path: str, source_project: str, mode: str, action: str = "export", log_path: Optional[str] = None
) -> List[dict]:
    if not db_path or not os.path.exists(db_path):
        raise ValueError(f"Database not found at '{db_path}'")

    conn = sqlite3.connect(f"file:{os.path.abspath(db_path)}?mode=ro", uri=True)
    cur = conn.cursor()

    candidates = []

    try:
        # Check for Delay Events WBS
        try:
            cur.execute("SELECT DISTINCT PROJ_ID FROM PROJWBS WHERE WBS_NAME = 'Delay Events'")
            delay_wbs_proj_ids = {r[0] for r in cur.fetchall() if r[0] is not None}
        except Exception:
            delay_wbs_proj_ids = set()

        if mode == "1":
            # Mode 1: Last EOT run from log
            if not log_path or not os.path.exists(log_path):
                if os.path.exists(DEFAULT_LOG_PATH):
                    log_path = DEFAULT_LOG_PATH
                else:
                    conn.close()
                    raise ValueError("Mode 1 requires a valid existing process log file.")

            sessions = parse_log_sessions(log_path)
            if not sessions:
                conn.close()
                return []

            target_names = sessions[0]["projects"]
            for name in target_names:
                cur.execute(
                    "SELECT PROJ_ID, PROJ_SHORT_NAME, PROJ_NAME FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
                    (name,),
                )
                row = cur.fetchone()
                if row:
                    candidates.append({"id": row[0], "code": row[1], "name": row[2] or row[1]})

        elif mode == "2":
            # Mode 2: Associated with baseline project
            cur.execute(
                "SELECT PROJ_ID FROM PROJECT WHERE PROJ_SHORT_NAME = ? AND PROJECT_FLAG = 'Y'",
                (source_project,),
            )
            base_row = cur.fetchone()
            if not base_row:
                conn.close()
                raise ValueError(f"Baseline project '{source_project}' not found in database.")
            source_proj_id = base_row[0]

            cur.execute("SELECT TASK_CODE FROM TASK WHERE PROJ_ID = ?", (source_proj_id,))
            source_task_codes = {r[0] for r in cur.fetchall() if r[0]}

            cur.execute(
                "SELECT PROJ_ID, PROJ_SHORT_NAME, PROJ_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y' AND PROJ_SHORT_NAME != ?",
                (source_project,),
            )
            all_other = cur.fetchall()

            for pid, name, full_name in all_other:
                is_match = False
                if name and source_project in name:
                    is_match = True
                elif name and (is_delay_project_name(name) or pid in delay_wbs_proj_ids):
                    cur.execute("SELECT TASK_CODE FROM TASK WHERE PROJ_ID = ? AND TASK_CODE != ?", (pid, name))
                    proj_tasks = {r[0] for r in cur.fetchall() if r[0]}
                    if proj_tasks.intersection(source_task_codes):
                        is_match = True

                if is_match:
                    candidates.append({"id": pid, "code": name, "name": full_name or name})

        elif mode == "3":
            # Mode 3:
            # For Export: all projects in database
            # For Deletion: all delay projects matching pattern
            cur.execute(
                "SELECT PROJ_ID, PROJ_SHORT_NAME, PROJ_NAME FROM PROJECT WHERE PROJECT_FLAG = 'Y' ORDER BY PROJ_SHORT_NAME"
            )
            all_projs = cur.fetchall()
            for pid, name, full_name in all_projs:
                if action == "export":
                    candidates.append({"id": pid, "code": name, "name": full_name or name})
                else:  # delete
                    if name != source_project and (is_delay_project_name(name) or pid in delay_wbs_proj_ids):
                        candidates.append({"id": pid, "code": name, "name": full_name or name})

    finally:
        conn.close()

    # Assign 1-based serials
    result = []
    for idx, c in enumerate(candidates, 1):
        result.append({"serial": idx, "id": c["id"], "code": c["code"], "name": c["name"]})
    return result


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES - PAGES & STATIC VIEWS
# ─────────────────────────────────────────────────────────────────────────────


@app.get("/", response_class=HTMLResponse)
async def serve_index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES - DASHBOARD & SYSTEM STATS
# ─────────────────────────────────────────────────────────────────────────────


@app.get("/api/dashboard/stats")
async def get_dashboard_stats(db_path: Optional[str] = None):
    # Determine DB path
    target_db = db_path if db_path and os.path.exists(db_path) else DEFAULT_DB_PATH
    db_info = inspect_database_details(target_db) if target_db else {"connected": False}

    # Parse last activity from log
    last_activity_time = "No activity recorded"
    last_activity_msg = "Run an analysis or export to generate logs"

    if os.path.exists(DEFAULT_LOG_PATH):
        try:
            with open(DEFAULT_LOG_PATH, "r", encoding="utf-8", errors="ignore") as f:
                lines = f.readlines()
                for line in reversed(lines):
                    line_stripped = line.strip()
                    if not line_stripped:
                        continue
                    if len(line_stripped) >= 19 and re.match(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", line_stripped[:19]):
                        last_activity_time = line_stripped[:19]
                        parts = (
                            line_stripped.split(" - ")
                            if " - " in line_stripped
                            else line_stripped.split("] ")
                        )
                        last_activity_msg = parts[-1].strip() if len(parts) > 1 else line_stripped[20:].strip()
                        break
        except Exception:
            pass

    return {
        "db": db_info,
        "last_activity": {
            "timestamp": last_activity_time,
            "message": last_activity_msg,
        },
    }


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES - FILE MANAGEMENT & UPLOADS
# ─────────────────────────────────────────────────────────────────────────────


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = f"{timestamp}_{file.filename}"
        dest_path = os.path.join(UPLOADS_DIR, safe_name)

        with open(dest_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        return {
            "success": True,
            "filename": file.filename,
            "saved_path": dest_path,
            "size": os.path.getsize(dest_path),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"File upload failed: {e}")


@app.get("/api/download/template")
async def download_template():
    candidates = [
        TEMPLATE_EXCEL_PATH,
        os.path.join(_SCRIPT_DIR, "delay_events_template.xlsx"),
    ]
    template_path = next((p for p in candidates if os.path.exists(p)), None)

    if not template_path:
        try:
            from db_eot_automation import create_template

            template_path = create_template(TEMPLATE_EXCEL_PATH)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Could not generate template: {e}")

    return FileResponse(
        template_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Delay_Events_template.xlsx",
    )


@app.get("/api/download/log")
async def download_log():
    if not os.path.exists(DEFAULT_LOG_PATH):
        raise HTTPException(status_code=404, detail="Log file does not exist yet.")
    return FileResponse(DEFAULT_LOG_PATH, media_type="text/plain", filename="eot_process_logs.log")


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES - DB INSPECTION & PROJECT SCANNING
# ─────────────────────────────────────────────────────────────────────────────


@app.post("/api/db/inspect")
async def inspect_db_endpoint(payload: dict):
    db_path = payload.get("db_path", "").strip().strip("'\"")
    if not db_path:
        raise HTTPException(status_code=400, detail="Database path cannot be empty.")
    result = inspect_database_details(db_path)
    return result


@app.post("/api/projects/scan")
async def scan_projects_endpoint(payload: dict):
    db_path = payload.get("db_path", "").strip().strip("'\"")
    source_project = payload.get("source_project", "").strip()
    mode = str(payload.get("mode", "1")).strip()
    action = payload.get("action", "export").strip().lower()
    log_path = payload.get("log_path", "").strip().strip("'\"") or None

    try:
        candidates = scan_candidate_projects(
            db_path=db_path,
            source_project=source_project,
            mode=mode,
            action=action,
            log_path=log_path,
        )
        return {"success": True, "count": len(candidates), "projects": candidates}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES - TASK LAUNCHERS
# ─────────────────────────────────────────────────────────────────────────────


@app.post("/api/run/individual")
async def run_individual_eot(payload: dict):
    excel_path = payload.get("excel_path", "").strip().strip("'\"")
    db_path = payload.get("db_path", "").strip().strip("'\"")
    source_project = payload.get("source_project", "").strip()
    milestone_mode = payload.get("milestone_mode", "standard").strip()
    milestone_config = payload.get("milestone_config", "").strip().strip("'\"")
    remarks = payload.get("remarks", "").strip()

    if not excel_path or not os.path.exists(excel_path):
        raise HTTPException(status_code=400, detail="Valid Delay Events Excel file is required.")
    if not db_path or not os.path.exists(db_path):
        raise HTTPException(status_code=400, detail="Valid Primavera SQLite DB file is required.")
    if not source_project:
        raise HTTPException(status_code=400, detail="Baseline Project ID is required.")
    if milestone_mode == "custom" and (not milestone_config or not os.path.exists(milestone_config)):
        raise HTTPException(status_code=400, detail="Valid custom activities/milestones Excel file is required.")

    task_id = f"ind_{uuid.uuid4().hex[:8]}"
    task_work_dir = os.path.join(OUTPUTS_DIR, task_id)
    os.makedirs(task_work_dir, exist_ok=True)

    explicit_output = os.path.join(task_work_dir, f"EOTR_{source_project}_{task_id}.xlsx")

    cmd = [
        sys.executable,
        "-u",
        DB_EOT_SCRIPT,
        "--excel",
        excel_path,
        "--source-project",
        source_project,
        "--db-path",
        db_path,
        "--output",
        explicit_output,
        "--milestone-mode",
        milestone_mode,
    ]
    if milestone_mode == "custom" and milestone_config:
        cmd.extend(["--milestone-config", milestone_config])
    if remarks:
        cmd.extend(["--default-remarks", remarks])

    task = Task(task_id=task_id, task_type="individual", cmd=cmd, work_dir=task_work_dir)
    TASKS[task_id] = task

    threading.Thread(target=run_task_worker, args=(task,), daemon=True).start()
    return {"success": True, "task_id": task_id}


@app.post("/api/run/combined")
async def run_combined_eot(payload: dict):
    excel_path = payload.get("excel_path", "").strip().strip("'\"")
    db_path = payload.get("db_path", "").strip().strip("'\"")
    source_project = payload.get("source_project", "").strip()
    milestone_mode = payload.get("milestone_mode", "standard").strip()
    milestone_config = payload.get("milestone_config", "").strip().strip("'\"")
    remarks = payload.get("remarks", "").strip()

    if not excel_path or not os.path.exists(excel_path):
        raise HTTPException(status_code=400, detail="Valid Delay Events Excel file is required.")
    if not db_path or not os.path.exists(db_path):
        raise HTTPException(status_code=400, detail="Valid Primavera SQLite DB file is required.")
    if not source_project:
        raise HTTPException(status_code=400, detail="Baseline Project ID is required.")
    if milestone_mode == "custom" and (not milestone_config or not os.path.exists(milestone_config)):
        raise HTTPException(status_code=400, detail="Valid custom activities/milestones Excel file is required.")

    task_id = f"comb_{uuid.uuid4().hex[:8]}"
    task_work_dir = os.path.join(OUTPUTS_DIR, task_id)
    os.makedirs(task_work_dir, exist_ok=True)

    explicit_output = os.path.join(task_work_dir, f"EOTR_Combined_{source_project}_{task_id}.xlsx")

    cmd = [
        sys.executable,
        "-u",
        ALL_DELAYS_SCRIPT,
        "--excel",
        excel_path,
        "--source-project",
        source_project,
        "--db-path",
        db_path,
        "--output",
        explicit_output,
        "--milestone-mode",
        milestone_mode,
    ]
    if milestone_mode == "custom" and milestone_config:
        cmd.extend(["--milestone-config", milestone_config])
    if remarks:
        cmd.extend(["--default-remarks", remarks])

    task = Task(task_id=task_id, task_type="combined", cmd=cmd, work_dir=task_work_dir)
    TASKS[task_id] = task

    threading.Thread(target=run_task_worker, args=(task,), daemon=True).start()
    return {"success": True, "task_id": task_id}


@app.post("/api/run/export")
async def run_export_utility(payload: dict):
    db_path = payload.get("db_path", "").strip().strip("'\"")
    source_project = payload.get("source_project", "").strip()
    mode = str(payload.get("mode", "1")).strip()
    log_path = payload.get("log_path", "").strip().strip("'\"") or None
    serials = payload.get("serials", "").strip()

    if not db_path or not os.path.exists(db_path):
        raise HTTPException(status_code=400, detail="Valid Primavera SQLite DB file is required.")
    if not source_project:
        raise HTTPException(status_code=400, detail="Baseline Project ID is required.")

    task_id = f"exp_{uuid.uuid4().hex[:8]}"
    task_work_dir = os.path.join(OUTPUTS_DIR, task_id)
    os.makedirs(task_work_dir, exist_ok=True)

    out_xer_dir = os.path.join(task_work_dir, "xers")
    os.makedirs(out_xer_dir, exist_ok=True)

    cmd = [
        sys.executable,
        "-u",
        EXPORT_EOT_SCRIPT,
        "--db-path",
        db_path,
        "--project",
        source_project,
        "--mode",
        mode,
        "--out-dir",
        out_xer_dir,
        "--yes",
    ]
    if mode == "1" and log_path:
        cmd.extend(["--log-path", log_path])
    if mode == "2" and serials:
        cmd.extend(["--serials", serials])

    task = Task(task_id=task_id, task_type="export", cmd=cmd, work_dir=task_work_dir)
    TASKS[task_id] = task

    threading.Thread(target=run_task_worker, args=(task,), daemon=True).start()
    return {"success": True, "task_id": task_id}


@app.post("/api/run/delete")
async def run_deletion_utility(payload: dict):
    db_path = payload.get("db_path", "").strip().strip("'\"")
    source_project = payload.get("source_project", "").strip()
    mode = str(payload.get("mode", "1")).strip()
    log_path = payload.get("log_path", "").strip().strip("'\"") or None
    serials = payload.get("serials", "").strip()

    if not db_path or not os.path.exists(db_path):
        raise HTTPException(status_code=400, detail="Valid Primavera SQLite DB file is required.")
    if not source_project:
        raise HTTPException(status_code=400, detail="Baseline Project ID is required.")

    task_id = f"del_{uuid.uuid4().hex[:8]}"
    task_work_dir = os.path.join(OUTPUTS_DIR, task_id)
    os.makedirs(task_work_dir, exist_ok=True)

    cmd = [
        sys.executable,
        "-u",
        DELETE_PROJECTS_SCRIPT,
        "--db-path",
        db_path,
        "--project",
        source_project,
        "--mode",
        mode,
        "--yes",
    ]
    if mode == "1" and log_path:
        cmd.extend(["--log-path", log_path])
    if mode == "2" and serials:
        cmd.extend(["--serials", serials])

    task = Task(task_id=task_id, task_type="delete", cmd=cmd, work_dir=task_work_dir)
    TASKS[task_id] = task

    threading.Thread(target=run_task_worker, args=(task,), daemon=True).start()
    return {"success": True, "task_id": task_id}


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES - TASK INTERACTIONS (STATUS, LOGS, CANCEL, STDIN, DOWNLOADS)
# ─────────────────────────────────────────────────────────────────────────────


@app.get("/api/tasks/{task_id}/status")
async def get_task_status(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found.")
    return {
        "task_id": task.task_id,
        "type": task.task_type,
        "status": task.status,
        "exit_code": task.exit_code,
        "created_at": task.created_at,
        "completed_at": task.completed_at,
        "has_discrepancies": bool(task.discrepancies),
        "discrepancies": task.discrepancies,
        "output_files": [os.path.basename(f) for f in task.output_files],
    }


@app.get("/api/tasks/{task_id}/logs")
async def get_task_logs(task_id: str, offset: int = 0):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found.")

    with task.lock:
        all_logs = "".join(task.logs)

    new_logs = all_logs[offset:] if offset < len(all_logs) else ""
    return {
        "offset": len(all_logs),
        "logs": new_logs,
        "status": task.status,
        "exit_code": task.exit_code,
    }


@app.post("/api/tasks/{task_id}/stdin")
async def send_task_stdin(task_id: str, payload: dict):
    task = TASKS.get(task_id)
    if not task or not task.process or task.process.poll() is not None:
        raise HTTPException(status_code=400, detail="Task process is not currently running.")

    text = payload.get("input", "")
    try:
        task.process.stdin.write(text + "\n")
        task.process.stdin.flush()
        task.add_log(f"\n> {text}\n")
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to send stdin: {e}")


@app.post("/api/tasks/{task_id}/kill")
async def kill_task(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found.")

    if task.process and task.process.poll() is None:
        task.status = "cancelled"
        try:
            task.process.kill()
        except Exception:
            pass
        task.add_log("\n\n[Process terminated by user request]")
        return {"success": True, "message": "Task killed successfully."}
    return {"success": True, "message": "Task was not running."}


@app.get("/api/tasks/{task_id}/download-output")
async def download_task_output(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found.")

    excel_files = [f for f in task.output_files if f.endswith(".xlsx")]
    if not excel_files:
        raise HTTPException(status_code=404, detail="No Excel output report found for this task.")

    file_path = excel_files[0]
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(file_path),
    )


@app.get("/api/tasks/{task_id}/download-xers")
async def download_task_xers(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found.")

    xer_files = [f for f in task.output_files if f.endswith(".xer")]
    if not xer_files:
        # Check task directory
        for root, _, files in os.walk(task.work_dir):
            for file in files:
                if file.endswith(".xer"):
                    xer_files.append(os.path.join(root, file))

    if not xer_files:
        raise HTTPException(status_code=404, detail="No exported .XER files found for this task.")

    if len(xer_files) == 1:
        return FileResponse(
            xer_files[0],
            media_type="application/octet-stream",
            filename=os.path.basename(xer_files[0]),
        )

    # If multiple, zip them
    zip_path = os.path.join(task.work_dir, f"exported_projects_{task_id}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in xer_files:
            zf.write(f, os.path.basename(f))

    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=os.path.basename(zip_path),
    )


# ─────────────────────────────────────────────────────────────────────────────
# WEBSOCKET STREAMING
# ─────────────────────────────────────────────────────────────────────────────


@app.websocket("/ws/logs/{task_id}")
async def websocket_logs_endpoint(websocket: WebSocket, task_id: str):
    await websocket.accept()
    task = TASKS.get(task_id)
    if not task:
        await websocket.send_text(json.dumps({"type": "error", "message": "Task not found."}))
        await websocket.close()
        return

    task.clients.add(websocket)

    # Send existing backlog
    with task.lock:
        backlog = "".join(task.logs)
    if backlog:
        await websocket.send_text(json.dumps({"type": "backlog", "data": backlog}))

    # Send current status
    await websocket.send_text(
        json.dumps(
            {
                "type": "status_change",
                "payload": {
                    "status": task.status,
                    "exit_code": task.exit_code,
                    "discrepancies": bool(task.discrepancies),
                    "outputs_count": len(task.output_files),
                },
            }
        )
    )

    try:
        while True:
            # Keep connection alive & listen for client messages (like stdin or cancel)
            data = await websocket.receive_text()
            try:
                msg = json.loads(data)
                if msg.get("type") == "stdin":
                    inp = msg.get("data", "")
                    if task.process and task.process.poll() is None:
                        task.process.stdin.write(inp + "\n")
                        task.process.stdin.flush()
                        task.add_log(f"\n> {inp}\n")
                elif msg.get("type") == "kill":
                    if task.process and task.process.poll() is None:
                        task.status = "cancelled"
                        task.process.kill()
                        task.add_log("\n\n[Process terminated by user]")
            except Exception:
                pass
    except WebSocketDisconnect:
        pass
    finally:
        task.clients.discard(websocket)


# ─────────────────────────────────────────────────────────────────────────────
# SERVER ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", 8000))
    host = "0.0.0.0"  # Must bind to 0.0.0.0 for Railway / any cloud host
    print(f"Starting P6 EOT Toolkit Web App on http://{host}:{port}...")
    uvicorn.run("web_app:app", host=host, port=port, reload=False)

