/**
 * app.js - Primavera P6 EOT Automation Toolkit Web UI
 * Handles navigation, API calls, file uploads, task management,
 * real-time WebSocket log streaming, and dynamic UI updates.
 */

"use strict";

/* ──────────────────────────────────────────────
   STATE
────────────────────────────────────────────── */
const State = {
  dbPath: "",
  dbConnected: false,
  uploadedExcelPath: "",
  uploadedExcelName: "",
  uploadedMilestonePath: "",
  uploadedMilestoneName: "",
  logPath: "",
  activeSockets: {},   // task_id -> WebSocket
  activeTaskId: null,  // task currently shown in the log viewer
  tasks: {},           // task_id -> {type, status, exit_code, ...}
};

/* ──────────────────────────────────────────────
   NAVIGATION
────────────────────────────────────────────── */
function navigate(sectionId) {
  document.querySelectorAll(".page-section").forEach(s => s.classList.remove("active"));
  document.querySelectorAll(".nav-link").forEach(l => l.classList.remove("active"));

  const section = document.getElementById("section-" + sectionId);
  if (section) section.classList.add("active");

  const link = document.querySelector(`[data-nav="${sectionId}"]`);
  if (link) link.classList.add("active");

  document.getElementById("topbar-title").textContent =
    link?.getAttribute("data-title") || "EOT Automation";

  // Lazy loads
  if (sectionId === "dashboard") loadDashboard();
  if (sectionId === "tasks") renderTasksList();
  if (sectionId === "help") renderHelpSteps();
}

/* ──────────────────────────────────────────────
   TOASTS
────────────────────────────────────────────── */
function toast(msg, type = "info", duration = 4000) {
  const ct = document.getElementById("toast-container");
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  const icon = { success:"✓", error:"✗", info:"ℹ", warning:"⚠" }[type] || "•";
  el.innerHTML = `<span>${icon}</span><span>${escHtml(msg)}</span>`;
  ct.appendChild(el);
  setTimeout(() => {
    el.classList.add("removing");
    setTimeout(() => el.remove(), 300);
  }, duration);
}

/* ──────────────────────────────────────────────
   HELPERS
────────────────────────────────────────────── */
function escHtml(s) {
  if (!s) return "";
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

function formatDate(isoStr) {
  if (!isoStr) return "—";
  try {
    const d = new Date(isoStr);
    return d.toLocaleDateString() + " " + d.toLocaleTimeString();
  } catch { return isoStr; }
}

function badgeHtml(status) {
  const map = {
    queued:    "badge-queued",
    running:   "badge-running pulse",
    completed: "badge-completed",
    failed:    "badge-failed",
    cancelled: "badge-cancelled",
  };
  const cls = map[status] || "badge-queued";
  const icons = { running:"⏳", completed:"✓", failed:"✗", cancelled:"◼", queued:"…" };
  return `<span class="badge ${cls}">${icons[status] || ""}${escHtml(status)}</span>`;
}

async function api(method, path, body) {
  const opts = {
    method,
    headers: { "Content-Type": "application/json" },
  };
  if (body !== undefined) opts.body = JSON.stringify(body);
  const res = await fetch(path, opts);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.detail || `HTTP ${res.status}`);
  return data;
}

function setLoading(btn, loading) {
  if (loading) {
    btn.dataset.origText = btn.innerHTML;
    btn.innerHTML = `<span class="spinner"></span> Working…`;
    btn.disabled = true;
  } else {
    btn.innerHTML = btn.dataset.origText || "Go";
    btn.disabled = false;
  }
}

/* ──────────────────────────────────────────────
   DASHBOARD
────────────────────────────────────────────── */
async function loadDashboard() {
  try {
    const dbParam = State.dbPath ? `?db_path=${encodeURIComponent(State.dbPath)}` : "";
    const data = await api("GET", `/api/dashboard/stats${dbParam}`);

    // DB tile
    const tile = document.getElementById("dash-db-tile");
    if (data.db && data.db.connected) {
      tile.innerHTML = `
        <div class="stat-icon">🗄️</div>
        <div class="stat-label">Database Projects</div>
        <div class="stat-value">${data.db.project_count}</div>
        <div class="stat-sub">${escHtml(data.db.db_name)}</div>`;
    } else {
      tile.innerHTML = `
        <div class="stat-icon">🗄️</div>
        <div class="stat-label">Database Projects</div>
        <div class="stat-value text-muted">—</div>
        <div class="stat-sub text-muted">No DB connected</div>`;
    }

    // Activity
    const el = document.getElementById("dash-last-activity");
    if (data.last_activity) {
      el.innerHTML = `
        <div class="stat-label">Last Activity</div>
        <div class="stat-value" style="font-size:14px;margin-top:4px">${escHtml(data.last_activity.timestamp)}</div>
        <div class="stat-sub">${escHtml(data.last_activity.message)}</div>`;
    }

    // Task count
    const total = Object.keys(State.tasks).length;
    const running = Object.values(State.tasks).filter(t => t.status === "running").length;
    document.getElementById("dash-tasks-tile").innerHTML = `
      <div class="stat-icon">⚙️</div>
      <div class="stat-label">Tasks This Session</div>
      <div class="stat-value">${total}</div>
      <div class="stat-sub">${running} running</div>`;

  } catch (e) {
    console.error("Dashboard error:", e);
  }
}

/* ──────────────────────────────────────────────
   DB CONFIGURATION
────────────────────────────────────────────── */
async function connectDatabase() {
  const btn = document.getElementById("btn-connect-db");
  const path = document.getElementById("inp-db-path").value.trim();
  if (!path) { toast("Please enter the database file path.", "warning"); return; }

  setLoading(btn, true);
  try {
    const data = await api("POST", "/api/db/inspect", { db_path: path });
    if (data.connected) {
      State.dbPath = path;
      State.dbConnected = true;
      updateDbStatusBar(data);
      populateProjectDropdowns(data.projects);
      toast(`Connected — ${data.project_count} projects found in ${data.db_name}`, "success");
    } else {
      toast(`Connection failed: ${data.error}`, "error");
    }
  } catch (e) {
    toast("Error: " + e.message, "error");
  } finally {
    setLoading(btn, false);
  }
}

function updateDbStatusBar(dbInfo) {
  const bars = document.querySelectorAll(".db-status-bar");
  bars.forEach(bar => {
    if (dbInfo && dbInfo.connected) {
      bar.innerHTML = `
        <div class="dot connected"></div>
        <div class="info">
          <div class="db-name">${escHtml(dbInfo.db_name)}</div>
          <div class="db-meta">${dbInfo.project_count} projects • ${escHtml(dbInfo.db_path)}</div>
        </div>`;
    } else {
      bar.innerHTML = `
        <div class="dot disconnected"></div>
        <div class="info">
          <div class="db-name">No database connected</div>
          <div class="db-meta">Configure the database in Settings</div>
        </div>`;
    }
  });
}

function populateProjectDropdowns(projects) {
  const selects = document.querySelectorAll(".project-select");
  selects.forEach(sel => {
    const cur = sel.value;
    sel.innerHTML = `<option value="">— Select baseline project —</option>`;
    projects.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p.code;
      opt.textContent = `${p.code}  —  ${p.name}`;
      sel.appendChild(opt);
    });
    if (cur) sel.value = cur;
  });
}

/* ──────────────────────────────────────────────
   FILE UPLOAD
────────────────────────────────────────────── */
function setupDropzone(zoneId, inputId, stateKey, nameKey, acceptLabel) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  if (!zone || !input) return;

  zone.addEventListener("click", () => input.click());
  zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("dragover");
    if (e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0], zone, stateKey, nameKey);
  });
  input.addEventListener("change", () => {
    if (input.files[0]) uploadFile(input.files[0], zone, stateKey, nameKey);
  });
}

async function uploadFile(file, zone, stateKey, nameKey) {
  const info = zone.querySelector(".uploaded-info") || document.createElement("div");
  info.className = "uploaded-info";
  info.textContent = `⏳ Uploading ${file.name}…`;
  if (!zone.contains(info)) zone.appendChild(info);

  try {
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "Upload failed");

    State[stateKey] = data.saved_path;
    State[nameKey] = data.filename;

    info.textContent = `✓ ${data.filename} (${(data.size / 1024).toFixed(1)} KB)`;
    toast(`Uploaded: ${data.filename}`, "success");
    syncUploadedPaths();
  } catch (e) {
    info.textContent = `✗ Error: ${e.message}`;
    toast("Upload failed: " + e.message, "error");
  }
}

function syncUploadedPaths() {
  // Sync excel path into all relevant inputs
  document.querySelectorAll(".excel-path-display").forEach(el => {
    el.textContent = State.uploadedExcelName || "None";
    el.classList.toggle("text-muted", !State.uploadedExcelName);
  });
  document.querySelectorAll(".milestone-path-display").forEach(el => {
    el.textContent = State.uploadedMilestoneName || "None";
    el.classList.toggle("text-muted", !State.uploadedMilestoneName);
  });
}

/* ──────────────────────────────────────────────
   MILESTONE MODE TOGGLE
────────────────────────────────────────────── */
function setupMilestoneToggle(radioGroupId, panelId) {
  const group = document.getElementById(radioGroupId);
  const panel = document.getElementById(panelId);
  if (!group || !panel) return;

  group.querySelectorAll("input[type=radio]").forEach(radio => {
    radio.addEventListener("change", () => {
      panel.classList.toggle("hidden", radio.value !== "custom");
      group.querySelectorAll(".radio-card").forEach(card => {
        card.classList.toggle("selected", card.querySelector("input[type=radio]").checked);
      });
    });
  });
}

/* ──────────────────────────────────────────────
   PROJECT SCANNING (for Export / Delete)
────────────────────────────────────────────── */
async function scanProjects(section) {
  const btn = document.getElementById(`btn-scan-${section}`);
  const modeEl = document.querySelector(`input[name="mode-${section}"]:checked`);
  const projectEl = document.getElementById(`sel-project-${section}`);
  const tableContainer = document.getElementById(`scan-results-${section}`);

  if (!State.dbConnected) { toast("Please connect to a database first.", "warning"); return; }
  if (!modeEl) { toast("Select a mode first.", "warning"); return; }

  const mode = modeEl.value;
  const source_project = projectEl?.value || "";

  if ((mode === "2") && !source_project) {
    toast("Select a baseline project for Mode 2.", "warning"); return;
  }

  setLoading(btn, true);
  try {
    const data = await api("POST", "/api/projects/scan", {
      db_path: State.dbPath,
      source_project,
      mode,
      action: section,
      log_path: State.logPath || null,
    });

    if (data.projects && data.projects.length > 0) {
      renderProjectsTable(tableContainer, data.projects);
      toast(`Found ${data.count} candidate project(s).`, "success");
    } else {
      tableContainer.innerHTML = `<div class="tbl-placeholder">No candidate projects found for selected mode.</div>`;
      toast("No projects found.", "info");
    }
  } catch (e) {
    toast("Scan error: " + e.message, "error");
  } finally {
    setLoading(btn, false);
  }
}

function renderProjectsTable(container, projects) {
  let html = `<div class="tbl-wrapper"><table class="projects-table">
    <thead><tr>
      <th>#</th><th>Code</th><th>Name</th>
    </tr></thead><tbody>`;
  projects.forEach(p => {
    html += `<tr>
      <td>${p.serial}</td>
      <td class="code">${escHtml(p.code)}</td>
      <td>${escHtml(p.name)}</td>
    </tr>`;
  });
  html += `</tbody></table></div>`;
  container.innerHTML = html;
}

/* ──────────────────────────────────────────────
   TASK LAUNCHING
────────────────────────────────────────────── */
function getCommonPayload(prefix) {
  return {
    excel_path: State.uploadedExcelPath,
    db_path: State.dbPath,
    source_project: document.getElementById(`sel-project-${prefix}`)?.value || "",
    milestone_mode: document.querySelector(`input[name="ms-mode-${prefix}"]:checked`)?.value || "standard",
    milestone_config: State.uploadedMilestonePath || "",
    remarks: document.getElementById(`inp-remarks-${prefix}`)?.value?.trim() || "",
  };
}

function validateCommonPayload(p, needsExcel = true) {
  if (needsExcel && !p.excel_path) {
    toast("Please upload a Delay Events Excel file first.", "warning"); return false;
  }
  if (!p.db_path) {
    toast("Please connect to a Primavera database first.", "warning"); return false;
  }
  if (!p.source_project) {
    toast("Please select a baseline project.", "warning"); return false;
  }
  return true;
}

async function launchTask(endpoint, payload, btn) {
  setLoading(btn, true);
  try {
    const data = await api("POST", endpoint, payload);
    if (data.task_id) {
      State.tasks[data.task_id] = { type: endpoint.split("/").pop(), status: "queued" };
      subscribeTask(data.task_id);
      openTaskLog(data.task_id);
      navigate("logs");
      toast(`Task started: ${data.task_id}`, "success");
    }
  } catch (e) {
    toast("Failed to start task: " + e.message, "error");
  } finally {
    setLoading(btn, false);
  }
}

function launchIndividual() {
  const btn = document.getElementById("btn-run-individual");
  const p = getCommonPayload("individual");
  if (!validateCommonPayload(p)) return;
  launchTask("/api/run/individual", p, btn);
}

function launchCombined() {
  const btn = document.getElementById("btn-run-combined");
  const p = getCommonPayload("combined");
  if (!validateCommonPayload(p)) return;
  launchTask("/api/run/combined", p, btn);
}

function launchExport() {
  const btn = document.getElementById("btn-run-export");
  const modeEl = document.querySelector('input[name="mode-export"]:checked');
  const serials = document.getElementById("inp-serials-export")?.value?.trim() || "";
  if (!State.dbConnected) { toast("Connect to a database first.", "warning"); return; }
  const source_project = document.getElementById("sel-project-export")?.value || "";
  if (!source_project) { toast("Select a baseline project.", "warning"); return; }
  const mode = modeEl?.value || "1";
  if (mode === "2" && !serials) { toast("Enter project serial numbers for Mode 2.", "warning"); return; }

  launchTask("/api/run/export", {
    db_path: State.dbPath,
    source_project,
    mode,
    log_path: State.logPath || null,
    serials,
  }, btn);
}

function launchDelete() {
  const btn = document.getElementById("btn-run-delete");
  const modeEl = document.querySelector('input[name="mode-delete"]:checked');
  const serials = document.getElementById("inp-serials-delete")?.value?.trim() || "";
  if (!State.dbConnected) { toast("Connect to a database first.", "warning"); return; }
  const source_project = document.getElementById("sel-project-delete")?.value || "";
  if (!source_project) { toast("Select a baseline project.", "warning"); return; }
  const mode = modeEl?.value || "1";

  // Confirmation for delete
  if (!confirm(`⚠️ This will permanently DELETE Delay Projects from the database.\n\nBaseline: ${source_project}\nMode: ${mode}\n\nAre you absolutely sure?`)) return;

  launchTask("/api/run/delete", {
    db_path: State.dbPath,
    source_project,
    mode,
    log_path: State.logPath || null,
    serials,
  }, btn);
}

/* ──────────────────────────────────────────────
   LOG VIEWER & WEBSOCKET
────────────────────────────────────────────── */
function openTaskLog(taskId) {
  State.activeTaskId = taskId;
  const terminal = document.getElementById("log-terminal");
  terminal.innerHTML = `<span class="log-placeholder">Connecting to task ${escHtml(taskId)}…</span>`;
  document.getElementById("log-task-id").textContent = taskId;
  document.getElementById("log-task-status").innerHTML = badgeHtml("queued");

  // Show/hide control buttons
  document.getElementById("btn-kill-task").dataset.taskId = taskId;
  document.getElementById("btn-dl-output").dataset.taskId = taskId;
  document.getElementById("btn-dl-xers").dataset.taskId = taskId;

  subscribeTask(taskId);
}

function subscribeTask(taskId) {
  if (State.activeSockets[taskId]) return; // already subscribed

  const protocol = location.protocol === "https:" ? "wss:" : "ws:";
  const wsUrl = `${protocol}//${location.host}/ws/logs/${taskId}`;
  const ws = new WebSocket(wsUrl);
  State.activeSockets[taskId] = ws;

  const terminal = document.getElementById("log-terminal");
  let isBacklogReceived = false;

  ws.onopen = () => {
    if (State.activeTaskId === taskId) {
      terminal.innerHTML = "";
      appendLog(terminal, `[WebSocket connected — streaming live output]\n`, "log-info");
    }
  };

  ws.onmessage = (e) => {
    try {
      const msg = JSON.parse(e.data);

      if (msg.type === "backlog") {
        if (State.activeTaskId === taskId) {
          terminal.innerHTML = "";
          appendColoredLog(terminal, msg.data);
          isBacklogReceived = true;
        }
      } else if (msg.type === "log") {
        if (State.activeTaskId === taskId) {
          appendColoredLog(terminal, msg.data);
        }
      } else if (msg.type === "status_change") {
        const pl = msg.payload || {};
        State.tasks[taskId] = { ...State.tasks[taskId], ...pl };
        if (State.activeTaskId === taskId) {
          document.getElementById("log-task-status").innerHTML = badgeHtml(pl.status);
          if (pl.status === "completed") {
            appendLog(terminal, "\n✓ Process completed successfully.\n", "log-success");
            showDownloadButtons(taskId, pl.outputs_count);
          } else if (pl.status === "failed") {
            appendLog(terminal, "\n✗ Process terminated with errors.\n", "log-error");
          }
        }
        // Refresh tasks list if visible
        const tasksSection = document.getElementById("section-tasks");
        if (tasksSection?.classList.contains("active")) renderTasksList();
      } else if (msg.type === "discrepancies") {
        const dp = msg.payload?.discrepancies;
        if (dp && State.activeTaskId === taskId) {
          const box = document.getElementById("discrepancies-box");
          if (box) {
            box.classList.remove("hidden");
            box.querySelector("pre").textContent = dp;
          }
        }
      }
    } catch {}
  };

  ws.onerror = () => {
    if (State.activeTaskId === taskId) {
      appendLog(terminal, "\n[WebSocket connection error]\n", "log-error");
    }
  };

  ws.onclose = () => {
    delete State.activeSockets[taskId];
    if (State.activeTaskId === taskId && State.tasks[taskId]?.status === "running") {
      appendLog(terminal, "\n[WebSocket disconnected]\n", "log-warn");
    }
  };
}

function appendLog(terminal, text, cls = "") {
  const span = document.createElement("span");
  if (cls) span.className = cls;
  span.textContent = text;
  terminal.appendChild(span);
  terminal.scrollTop = terminal.scrollHeight;
}

function appendColoredLog(terminal, text) {
  // Simple colorizer based on content patterns
  const lines = text.split("\n");
  lines.forEach((line, i) => {
    let cls = "";
    const l = line.toLowerCase();
    if (l.includes("error") || l.includes("failed") || l.includes("✗")) cls = "log-error";
    else if (l.includes("warning") || l.includes("warn") || l.includes("skipped")) cls = "log-warn";
    else if (l.includes("success") || l.includes("✓") || l.includes("completed")) cls = "log-success";
    else if (l.includes("info") || l.startsWith("[") || l.includes("→") || l.includes("->")) cls = "log-info";
    else if (l.includes("enter ") || l.includes("select ") || line.endsWith("?")) cls = "log-prompt";

    const span = document.createElement("span");
    if (cls) span.className = cls;
    span.textContent = line + (i < lines.length - 1 ? "\n" : "");
    terminal.appendChild(span);
  });
  terminal.scrollTop = terminal.scrollHeight;
}

function showDownloadButtons(taskId, outputsCount) {
  const dlBtn = document.getElementById("btn-dl-output");
  const xersBtn = document.getElementById("btn-dl-xers");
  if (dlBtn) dlBtn.classList.toggle("hidden", outputsCount === 0);
  if (xersBtn) xersBtn.classList.remove("hidden");
}

async function killActiveTask() {
  const taskId = document.getElementById("btn-kill-task")?.dataset.taskId;
  if (!taskId) return;
  try {
    await api("POST", `/api/tasks/${taskId}/kill`);
    toast("Task termination requested.", "warning");
  } catch (e) {
    toast("Kill failed: " + e.message, "error");
  }
}

function sendStdinInput() {
  const taskId = State.activeTaskId;
  if (!taskId) return;
  const inp = document.getElementById("stdin-input");
  const val = inp?.value?.trim();
  if (!val) return;

  const ws = State.activeSockets[taskId];
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: "stdin", data: val }));
  } else {
    api("POST", `/api/tasks/${taskId}/stdin`, { input: val })
      .then(() => toast("Input sent.", "info"))
      .catch(e => toast("Send failed: " + e.message, "error"));
  }
  if (inp) inp.value = "";
}

function downloadOutput(type) {
  const taskId = State.activeTaskId;
  if (!taskId) { toast("No active task selected.", "warning"); return; }
  const ep = type === "xer"
    ? `/api/tasks/${taskId}/download-xers`
    : `/api/tasks/${taskId}/download-output`;
  window.location.href = ep;
}

function clearLog() {
  const terminal = document.getElementById("log-terminal");
  if (terminal) {
    terminal.innerHTML = `<span class="log-placeholder">Log cleared.</span>`;
  }
  const box = document.getElementById("discrepancies-box");
  if (box) box.classList.add("hidden");
}

/* ──────────────────────────────────────────────
   TASKS LIST PAGE
────────────────────────────────────────────── */
function renderTasksList() {
  const container = document.getElementById("tasks-list-container");
  if (!container) return;

  const taskIds = Object.keys(State.tasks);
  if (taskIds.length === 0) {
    container.innerHTML = `<div class="tbl-placeholder">No tasks have been started yet.</div>`;
    return;
  }

  let html = "";
  [...taskIds].reverse().forEach(tid => {
    const t = State.tasks[tid];
    html += `
      <div class="task-item">
        <div class="task-info">
          <div class="task-id">${escHtml(tid)}</div>
          <div class="task-type">${escHtml(t.type || "task")}</div>
          <div class="task-meta">${badgeHtml(t.status || "queued")}</div>
        </div>
        <div class="task-actions">
          <button class="btn btn-ghost btn-sm" onclick="viewTask('${escHtml(tid)}')">View Logs</button>
        </div>
      </div>`;
  });
  container.innerHTML = html;
}

function viewTask(taskId) {
  openTaskLog(taskId);
  navigate("logs");
}

/* ──────────────────────────────────────────────
   HELP / GUIDE
────────────────────────────────────────────── */
function renderHelpSteps() {
  // Static content is already in HTML; we just ensure step images load
  const imgs = document.querySelectorAll(".step-img");
  imgs.forEach(img => {
    if (img.src) return; // already set
    img.src = img.dataset.src || "";
  });
}

/* ──────────────────────────────────────────────
   SETTINGS
────────────────────────────────────────────── */
function saveLogPath() {
  const val = document.getElementById("inp-log-path")?.value?.trim();
  if (val) {
    State.logPath = val;
    toast("Log path saved.", "success");
  }
}

function downloadLog() {
  window.location.href = "/api/download/log";
}

function downloadTemplate() {
  window.location.href = "/api/download/template";
}

/* ──────────────────────────────────────────────
   STDIN ENTER KEY
────────────────────────────────────────────── */
function setupStdinInput() {
  const inp = document.getElementById("stdin-input");
  if (inp) {
    inp.addEventListener("keydown", e => {
      if (e.key === "Enter") sendStdinInput();
    });
  }
}

/* ──────────────────────────────────────────────
   RADIO CARD INTERACTIVITY
────────────────────────────────────────────── */
function setupRadioCards() {
  document.querySelectorAll(".radio-card").forEach(card => {
    const radio = card.querySelector("input[type=radio]");
    if (!radio) return;
    card.addEventListener("click", () => {
      const group = card.closest(".radio-group");
      group?.querySelectorAll(".radio-card").forEach(c => c.classList.remove("selected"));
      radio.checked = true;
      card.classList.add("selected");
      radio.dispatchEvent(new Event("change", { bubbles: true }));
    });
    if (radio.checked) card.classList.add("selected");
  });
}

/* ──────────────────────────────────────────────
   MOBILE SIDEBAR TOGGLE
────────────────────────────────────────────── */
function toggleSidebar() {
  document.querySelector(".sidebar")?.classList.toggle("open");
}

/* ──────────────────────────────────────────────
   INIT
────────────────────────────────────────────── */
document.addEventListener("DOMContentLoaded", () => {
  // Navigation
  document.querySelectorAll(".nav-link").forEach(link => {
    link.addEventListener("click", () => {
      navigate(link.dataset.nav);
      document.querySelector(".sidebar")?.classList.remove("open"); // close on mobile
    });
  });

  // Upload dropzones
  setupDropzone("upload-zone-excel",     "file-input-excel",     "uploadedExcelPath",     "uploadedExcelName",    ".xlsx");
  setupDropzone("upload-zone-milestone", "file-input-milestone", "uploadedMilestonePath", "uploadedMilestoneName", ".xlsx");

  // Milestone mode toggles
  setupMilestoneToggle("ms-mode-group-individual", "ms-custom-panel-individual");
  setupMilestoneToggle("ms-mode-group-combined",   "ms-custom-panel-combined");

  // Radio cards
  setupRadioCards();

  // Stdin input
  setupStdinInput();

  // Default nav
  navigate("dashboard");
});

