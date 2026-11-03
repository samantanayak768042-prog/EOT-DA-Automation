"""
EOT_GUI.py - Unified Launcher Frontend for Primavera P6 EOT Automation Scripts
"""

import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Default paths
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB_PATH = ""
DEFAULT_LOG_PATH = os.path.join(_SCRIPT_DIR, "eot_process_logs.log")

# Script paths (1-to-1 mapping for each task in GUI panel)
DB_EOT_SCRIPT = os.path.join(_SCRIPT_DIR, "db_eot_automation.py")
ALL_DELAYS_SCRIPT = os.path.join(_SCRIPT_DIR, "all_delays_EOT.py")
EXPORT_EOT_SCRIPT = os.path.join(_SCRIPT_DIR, "export_EOT.py")
DELETE_PROJECTS_SCRIPT = os.path.join(_SCRIPT_DIR, "delete_projects.py")

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class ProcessRunner:
    """Runs a python script as a subprocess, streaming output in real-time."""

    def __init__(self, cmd, on_output_func, on_finished_func, auto_responses=None):
        self.cmd = cmd
        self.on_output = on_output_func
        self.on_finished = on_finished_func
        self.auto_responses = auto_responses or {}
        self.process = None
        self.thread = None
        self.is_running = False

    def start(self):
        self.is_running = True
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def _run(self):
        try:
            env = os.environ.copy()
            env["EOT_AUTO_NO_POPUP"] = "1"
            self.process = subprocess.Popen(
                self.cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )

            buffer = ""
            while self.is_running:
                char = self.process.stdout.read(1)
                if not char:
                    if self.process.poll() is not None:
                        break
                    continue

                self.on_output(char)
                buffer += char

                # Check for auto-responses
                for pattern, response in list(self.auto_responses.items()):
                    if pattern in buffer:
                        if callable(response):
                            response()
                        else:
                            self.send_input(response)
                        del self.auto_responses[pattern]
                        buffer = ""
                        break

                if len(buffer) > 500:
                    buffer = buffer[-500:]

            rc = self.process.wait()
            self.on_output(f"\n\n[Process completed with exit code {rc}]")
        except Exception as e:
            self.on_output(f"\n\n[Process execution error: {e}]")
        finally:
            self.is_running = False
            self.cleanup()
            self.on_finished()

    def send_input(self, text):
        if self.process and self.process.poll() is None:
            try:
                self.process.stdin.write(text + "\n")
                self.process.stdin.flush()
                self.on_output(f"\n> {text}\n")
            except Exception as e:
                self.on_output(f"\n[Error sending input: {e}]")

    def kill(self):
        self.is_running = False
        if self.process:
            try:
                self.process.kill()
            except Exception:
                pass
            self.cleanup()

    def cleanup(self):
        if self.process:
            try:
                self.process.stdout.close()
            except Exception:
                pass
            try:
                self.process.stdin.close()
            except Exception:
                pass


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("P6 EOT Automation Toolkit")
        self.geometry("1100x700")
        self.minsize(950, 600)
        # Open maximized by default after window load to prevent DPI scaling overrides
        self.after(50, self._maximize_window)

        self.runner = None
        self.current_page = "home"

        # Create Layout Grid
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Top Navigation Frame ---
        self.top_nav_frame = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.top_nav_frame.grid(row=0, column=0, sticky="ew")

        self.logo_label = ctk.CTkLabel(
            self.top_nav_frame,
            text="P6 EOT Toolkit",
            font=("Segoe UI", 16, "bold"),
        )
        self.logo_label.pack(side="left", padx=20, pady=10)

        # Navigation Buttons
        self.btn_home = ctk.CTkButton(
            self.top_nav_frame,
            text="Home",
            height=36,
            corner_radius=6,
            font=("Segoe UI", 12),
            command=lambda: self.select_frame("home"),
        )
        self.btn_home.pack(side="left", padx=5, pady=10)

        self.btn_ind = ctk.CTkButton(
            self.top_nav_frame,
            text="Individual EOT",
            height=36,
            corner_radius=6,
            font=("Segoe UI", 12),
            command=lambda: self.select_frame("individual"),
        )
        self.btn_ind.pack(side="left", padx=5, pady=10)

        self.btn_comb = ctk.CTkButton(
            self.top_nav_frame,
            text="Combined EOT",
            height=36,
            corner_radius=6,
            font=("Segoe UI", 12),
            command=lambda: self.select_frame("combined"),
        )
        self.btn_comb.pack(side="left", padx=5, pady=10)

        self.btn_export = ctk.CTkButton(
            self.top_nav_frame,
            text="Export EOT",
            height=36,
            corner_radius=6,
            font=("Segoe UI", 12),
            command=lambda: self.select_frame("export"),
        )
        self.btn_export.pack(side="left", padx=5, pady=10)

        self.btn_del = ctk.CTkButton(
            self.top_nav_frame,
            text="Delete Projects",
            height=36,
            corner_radius=6,
            font=("Segoe UI", 12),
            command=lambda: self.select_frame("delete"),
        )
        self.btn_del.pack(side="left", padx=5, pady=10)

        self.appearance_menu = ctk.CTkOptionMenu(
            self.top_nav_frame,
            values=["Light", "Dark", "System"],
            width=100,
            command=self.change_appearance_mode,
        )
        self.appearance_menu.set("Light")
        self.appearance_menu.pack(side="right", padx=20, pady=10)

        # --- Main Content Container ---
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        self.select_frame("home")

    def _maximize_window(self):
        try:
            self.state("zoomed")
        except Exception:
            try:
                self.attributes("-zoomed", True)
            except Exception:
                pass

    def change_appearance_mode(self, new_mode):
        ctk.set_appearance_mode(new_mode.lower())

    def get_active_db_path(self):
        for entry_name in ["ind_db_entry", "comb_db_entry", "exp_db_entry", "del_db_entry"]:
            if hasattr(self, entry_name):
                val = getattr(self, entry_name)
                if val:
                    path = val.get().strip().strip("'\"")
                    if path and os.path.exists(path):
                        return path
        return DEFAULT_DB_PATH

    def update_dashboard_stats(self):
        db_path = self.get_active_db_path()
        db_connected = False
        proj_count = 0
        db_name = "No Database Selected"

        if db_path and os.path.exists(db_path):
            try:
                conn = sqlite3.connect(db_path)
                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM PROJECT WHERE PROJECT_FLAG = 'Y'")
                row = cur.fetchone()
                if row:
                    proj_count = row[0]
                    db_connected = True
                    db_name = os.path.basename(db_path)
                conn.close()
            except Exception:
                pass

        last_activity = "No recent activity found"
        if os.path.exists(DEFAULT_LOG_PATH):
            try:
                with open(DEFAULT_LOG_PATH, "r", encoding="utf-8", errors="ignore") as f:
                    lines = f.readlines()
                    for line in reversed(lines):
                        line_stripped = line.strip()
                        if not line_stripped:
                            continue
                        if len(line_stripped) >= 19 and re.match(
                            r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", line_stripped[:19]
                        ):
                            parts = (
                                line_stripped.split(" - ")
                                if " - " in line_stripped
                                else line_stripped.split("] ")
                            )
                            msg = parts[-1].strip() if len(parts) > 1 else line_stripped[20:].strip()
                            last_activity = f"{line_stripped[:19]}\n{msg[:45]}..."
                            break
            except Exception:
                pass

        if hasattr(self, "lbl_stat_db"):
            if db_connected:
                self.lbl_stat_db.configure(text="● Connected", text_color="#28A745")
                self.lbl_db_name.configure(text=db_name)
            else:
                self.lbl_stat_db.configure(text="● Disconnected", text_color="#DC3545")
                self.lbl_db_name.configure(text="Offline / Missing DB")

        if hasattr(self, "lbl_stat_projs"):
            self.lbl_stat_projs.configure(text=str(proj_count))

        if hasattr(self, "lbl_stat_activity"):
            self.lbl_stat_activity.configure(text=last_activity)

    def select_frame(self, name):
        self.current_page = name

        unselected_fg = "transparent"
        unselected_text = ("#333333", "#CCCCCC")

        self.btn_home.configure(fg_color=unselected_fg, text_color=unselected_text, font=("Segoe UI", 12))
        self.btn_ind.configure(fg_color=unselected_fg, text_color=unselected_text, font=("Segoe UI", 12))
        self.btn_comb.configure(fg_color=unselected_fg, text_color=unselected_text, font=("Segoe UI", 12))
        self.btn_export.configure(fg_color=unselected_fg, text_color=unselected_text, font=("Segoe UI", 12))
        self.btn_del.configure(fg_color=unselected_fg, text_color=unselected_text, font=("Segoe UI", 12))

        selected_fg = ("#1f538d", "#1f538d")
        selected_text = "#FFFFFF"

        if name == "home":
            self.btn_home.configure(fg_color=selected_fg, text_color=selected_text, font=("Segoe UI", 12, "bold"))
        elif name == "individual":
            self.btn_ind.configure(fg_color=selected_fg, text_color=selected_text, font=("Segoe UI", 12, "bold"))
        elif name == "combined":
            self.btn_comb.configure(fg_color=selected_fg, text_color=selected_text, font=("Segoe UI", 12, "bold"))
        elif name == "export":
            self.btn_export.configure(fg_color=selected_fg, text_color=selected_text, font=("Segoe UI", 12, "bold"))
        elif name == "delete":
            self.btn_del.configure(fg_color=selected_fg, text_color=selected_text, font=("Segoe UI", 12, "bold"))

        if name not in self.frames:
            if name == "home":
                self.create_home_frame()
            elif name == "individual":
                self.create_ind_frame()
            elif name == "combined":
                self.create_comb_frame()
            elif name == "export":
                self.create_export_frame()
            elif name == "delete":
                self.create_del_frame()

        for f in self.frames.values():
            f.grid_remove()

        self.frames[name].grid(row=0, column=0, sticky="nsew")

    # ─────────────────────────────────────────────────────────────────────────────
    # VIEW BUILDERS
    # ─────────────────────────────────────────────────────────────────────────────

    def generate_guide_screenshots(self):
        assets_dir = os.path.join(_SCRIPT_DIR, "assets")
        os.makedirs(assets_dir, exist_ok=True)

    def show_db_guide_popup(self):
        self.generate_guide_screenshots()

        popup = ctk.CTkToplevel(self)
        popup.title("How to Find P6 SQLite Database Path")
        popup.geometry("650x750")
        popup.attributes("-topmost", True)

        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (650 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (750 // 2)
        popup.geometry(f"+{x}+{y}")

        scroll = ctk.CTkScrollableFrame(popup, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            scroll,
            text="Step-by-Step: Find Standalone SQLite DB Path",
            font=("Segoe UI", 16, "bold"),
        ).pack(pady=(0, 15))

        steps_data = [
            (
                "Step 1: Open P6 Login Window\nOn the login screen, click 'Edit database configuration'.",
                "step1.png",
                (350, 370),
            ),
            (
                "Step 2: Check Database List\nSelect your database alias (e.g. 'T3') and click 'Next'.",
                "step2.png",
                (450, 300),
            ),
            (
                "Step 3: Copy Database Location\nCopy the path from the 'Database Location' field and paste into the EOT Toolkit.",
                "step3.png",
                (450, 300),
            ),
        ]

        for idx, (desc, img_name, size) in enumerate(steps_data, 1):
            card = ctk.CTkFrame(
                scroll,
                fg_color=("#F7F9FC", "#151F2E"),
                corner_radius=8,
                border_width=1,
                border_color=("#E4E7EB", "#2D3748"),
            )
            card.pack(fill="x", pady=10)

            ctk.CTkLabel(
                card,
                text=desc,
                font=("Segoe UI", 12),
                justify="left",
                anchor="w",
            ).pack(fill="x", padx=15, pady=10)

            img_path = os.path.join(_SCRIPT_DIR, "assets", img_name)
            if os.path.exists(img_path):
                try:
                    from PIL import Image

                    pil_img = Image.open(img_path)
                    ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=size)
                    img_label = ctk.CTkLabel(card, image=ctk_img, text="")
                    img_label.image = ctk_img
                    img_label.pack(pady=(0, 15))
                except Exception as e:
                    ctk.CTkLabel(card, text=f"[Image Load Error: {e}]", text_color="red").pack(pady=(0, 15))
            else:
                ctk.CTkLabel(card, text="[Screenshot image not available]", text_color="gray").pack(pady=(0, 15))

    def download_excel_format(self):
        # Resolve bundled template (supporting both Delay_Events_template.xlsx and delay_events_template.xlsx)
        candidates = [
            os.path.join(_SCRIPT_DIR, "Delay_Events_template.xlsx"),
            os.path.join(_SCRIPT_DIR, "delay_events_template.xlsx"),
        ]
        template_src = next((p for p in candidates if os.path.exists(p)), candidates[0])

        # Generate on the fly if neither exists
        if not os.path.exists(template_src):
            try:
                from db_eot_automation import create_template

                create_template(template_src)
            except Exception as e:
                messagebox.showerror("Error", f"Could not generate template file: {e}")
                return

        dest_path = filedialog.asksaveasfilename(
            title="Save Input Excel Format Template",
            initialfile="Delay_Events_template.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if dest_path:
            try:
                shutil.copy(template_src, dest_path)
                messagebox.showinfo("Success", f"Excel format template saved successfully to:\n{dest_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {e}")

    def create_home_frame(self):
        home_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frames["home"] = home_frame

        home_frame.grid_rowconfigure(1, weight=1)
        home_frame.grid_columnconfigure(0, weight=1)

        # 1. Hero Banner Frame
        hero_frame = ctk.CTkFrame(
            home_frame,
            fg_color=("#1A2536", "#101D2E"),
            corner_radius=12,
            border_width=1,
            border_color=("#243B53", "#1B2A4A"),
        )
        hero_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        hero_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hero_frame,
            text="Primavera P6 Extension of Time (EOT) Toolkit",
            font=("Segoe UI", 26, "bold"),
            text_color="#FFFFFF",
            anchor="w",
        ).pack(fill="x", padx=20, pady=(20, 5))

        ctk.CTkLabel(
            hero_frame,
            text="An automated environment for critical path schedule delay simulations and project XER packaging.",
            font=("Segoe UI", 14),
            text_color="#9FB3C8",
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 20))

        # 2. Main Content Split Panel
        main_content = ctk.CTkFrame(home_frame, fg_color="transparent")
        main_content.grid(row=1, column=0, sticky="nsew")
        main_content.grid_rowconfigure(0, weight=1)
        main_content.grid_columnconfigure(0, weight=4)
        main_content.grid_columnconfigure(1, weight=6)

        # --- Left Pane: Instructions & Download Format ---
        left_pane = ctk.CTkFrame(
            main_content,
            fg_color=("#FFFFFF", "#1E1E1E"),
            corner_radius=10,
            border_width=1,
            border_color=("#E0E0E0", "#2D2D2D"),
        )
        left_pane.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_pane.grid_columnconfigure(0, weight=1)
        left_pane.grid_rowconfigure(0, weight=1)
        left_pane.grid_rowconfigure(1, weight=9)

        ctk.CTkLabel(
            left_pane,
            text="Instructions & Guidelines",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 10))

        inst_container = ctk.CTkFrame(left_pane, fg_color="transparent")
        inst_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=(5, 20))

        instructions = [
            {
                "title": "1. Locate your SQLite Database",
                "desc": "Standalone P6 databases (typically named 'P6ProStandalone.db') can be found in your local computer's 'Documents' folder.",
            },
            {
                "title": "2. Prepare Input Excel File",
                "desc": "Format your delay sheet to match the required columns, and optionally prepare an Excel file listing specific activities to report.",
            },
            {
                "title": "3. Run EOT Calculations",
                "desc": "Go to Individual EOT (for isolated events) or Combined EOT (for cumulative events), input details, and click Run to generate a single consolidated report.",
            },
            {
                "title": "4. Export & Clean up",
                "desc": "Export finished schedules as .XER files, then run the Deletion Utility to keep your SQLite database clean.",
            },
        ]

        for idx, item in enumerate(instructions):
            card = ctk.CTkFrame(
                inst_container,
                fg_color=("#F7F9FC", "#151F2E"),
                border_width=1,
                border_color=("#E4E7EB", "#2D3748"),
                corner_radius=8,
            )
            card.pack(fill="x", pady=6)

            title_lbl = ctk.CTkLabel(
                card,
                text=item["title"],
                font=("Segoe UI", 13, "bold"),
                text_color=("#1F538D", "#9FB3C8"),
                anchor="w",
                justify="left",
            )
            title_lbl.pack(fill="x", padx=15, pady=(10, 4))

            desc_lbl = ctk.CTkLabel(
                card,
                text=item["desc"],
                font=("Segoe UI", 12),
                text_color=("gray20", "gray80"),
                justify="left",
                anchor="w",
                wraplength=350,
            )

            has_link = idx in (0, 1)
            desc_lbl.pack(fill="x", padx=15, pady=(2, 10 if not has_link else 2))

            if idx == 0:
                guide_link = ctk.CTkLabel(
                    card,
                    text="How to find SQLite DB? Click here for guide with screenshots",
                    font=("Segoe UI", 11, "underline", "bold"),
                    text_color=("#2A75D3", "#6BA4E8"),
                    cursor="hand2",
                    anchor="w",
                )
                guide_link.pack(fill="x", padx=15, pady=(4, 10))
                guide_link.bind("<Button-1>", lambda e: self.show_db_guide_popup())

            elif idx == 1:
                download_link = ctk.CTkLabel(
                    card,
                    text="Click here to download input file excel format",
                    font=("Segoe UI", 11, "underline", "bold"),
                    text_color=("#2A75D3", "#6BA4E8"),
                    cursor="hand2",
                    anchor="w",
                )
                download_link.pack(fill="x", padx=15, pady=(4, 10))
                download_link.bind("<Button-1>", lambda e: self.download_excel_format())

        # --- Right Pane: Flowchart ---
        right_pane = ctk.CTkFrame(
            main_content,
            fg_color=("#FFFFFF", "#1E1E1E"),
            corner_radius=10,
            border_width=1,
            border_color=("#E0E0E0", "#2D2D2D"),
        )
        right_pane.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        ctk.CTkLabel(
            right_pane,
            text="Interactive Process Flowchart (Click Step to Navigate)",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(15, 10))

        flow_grid = ctk.CTkFrame(right_pane, fg_color="transparent")
        flow_grid.pack(fill="both", expand=True, padx=20, pady=(5, 15))
        flow_grid.grid_columnconfigure((0, 1), weight=1)

        flow_grid.grid_rowconfigure(0, weight=2)
        flow_grid.grid_rowconfigure(1, weight=1)
        flow_grid.grid_rowconfigure(2, weight=4)
        flow_grid.grid_rowconfigure(3, weight=1)
        flow_grid.grid_rowconfigure(4, weight=4)
        flow_grid.grid_rowconfigure(5, weight=1)
        flow_grid.grid_rowconfigure(6, weight=4)

        step1_card = ctk.CTkFrame(
            flow_grid,
            fg_color=("#F0F4F8", "#1A2536"),
            border_width=1,
            border_color=("#B0C4DE", "#243B53"),
            corner_radius=8,
        )
        step1_card.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=30, pady=2)
        ctk.CTkLabel(
            step1_card,
            text="STEP 1: Choose SQLite DB & Delay Excel File",
            font=("Segoe UI", 13, "bold"),
            text_color=("#1F538D", "#9FB3C8"),
        ).pack(expand=True, fill="both")

        ctk.CTkLabel(
            flow_grid,
            text="↓",
            font=("Segoe UI", 24, "bold"),
            text_color="gray",
        ).grid(row=1, column=0, columnspan=2, sticky="nsew", pady=2)

        card_ind = ctk.CTkButton(
            flow_grid,
            text="STEP 2A: Individual EOT\n\nRuns simulations for each Excel\ndelay event separately.",
            font=("Segoe UI", 13, "bold"),
            corner_radius=8,
            border_width=1,
            fg_color=("#E8F4FD", "#1E2D4A"),
            border_color=("#B0D4F1", "#2B476F"),
            text_color=("#102A43", "#FFFFFF"),
            hover_color=("#D4E9FC", "#243B5A"),
            command=lambda: self.select_frame("individual"),
        )
        card_ind.grid(row=2, column=0, sticky="nsew", padx=10, pady=2)

        card_comb = ctk.CTkButton(
            flow_grid,
            text="STEP 2B: Combined EOT\n\nSimulates cumulative delay impact\nof all events grouped together.",
            font=("Segoe UI", 13, "bold"),
            corner_radius=8,
            border_width=1,
            fg_color=("#E8F9EE", "#1E3A2B"),
            border_color=("#B8EBC6", "#2D5A3F"),
            text_color=("#0F3D1F", "#FFFFFF"),
            hover_color=("#D2F7DF", "#234732"),
            command=lambda: self.select_frame("combined"),
        )
        card_comb.grid(row=2, column=1, sticky="nsew", padx=10, pady=2)

        ctk.CTkLabel(
            flow_grid,
            text="↘",
            font=("Segoe UI", 24, "bold"),
            text_color="gray",
        ).grid(row=3, column=0, sticky="nsew", padx=(0, 45), pady=2)

        ctk.CTkLabel(
            flow_grid,
            text="↙",
            font=("Segoe UI", 24, "bold"),
            text_color="gray",
        ).grid(row=3, column=1, sticky="nsew", padx=(45, 0), pady=2)

        card_exp = ctk.CTkButton(
            flow_grid,
            text="STEP 3: Project Export Utility\n\nPackages baseline or replica projects into .XER files.",
            font=("Segoe UI", 13, "bold"),
            corner_radius=8,
            border_width=1,
            fg_color=("#FFF8E7", "#3A311E"),
            border_color=("#FADCA5", "#5C4B2E"),
            text_color=("#5C3E00", "#FFFFFF"),
            hover_color=("#FEEDC3", "#453A23"),
            command=lambda: self.select_frame("export"),
        )
        card_exp.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=30, pady=2)

        ctk.CTkLabel(
            flow_grid,
            text="↓",
            font=("Segoe UI", 24, "bold"),
            text_color="gray",
        ).grid(row=5, column=0, columnspan=2, sticky="nsew", pady=2)

        card_del = ctk.CTkButton(
            flow_grid,
            text="STEP 4: Project Deletion Utility\n\nCleanly wipes simulated replica projects from the SQLite database.",
            font=("Segoe UI", 13, "bold"),
            corner_radius=8,
            border_width=1,
            fg_color=("#FFEBEB", "#3D2121"),
            border_color=("#F5C2C2", "#5C2E2E"),
            text_color=("#610B0B", "#FFFFFF"),
            hover_color=("#FFD6D6", "#4A2727"),
            command=lambda: self.select_frame("delete"),
        )
        card_del.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=30, pady=2)

    def create_ind_frame(self):
        frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frames["individual"] = frame

        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Individual EOT Analysis", font=("Segoe UI", 20, "bold"), anchor="w").grid(
            row=0, column=0, sticky="ew", pady=(0, 15)
        )

        form_scroll = ctk.CTkScrollableFrame(frame)
        form_scroll.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        form_scroll.grid_columnconfigure(0, weight=1)

        # Delay Events Excel File
        ctk.CTkLabel(form_scroll, text="Delay Events Excel File:", anchor="w").pack(fill="x", padx=15, pady=(10, 0))
        f1 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f1.pack(fill="x", padx=15)
        self.ind_excel_entry = ctk.CTkEntry(f1, placeholder_text="Select .xlsx file...")
        self.ind_excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f1,
            text="Browse",
            width=80,
            command=lambda: self.browse_file(self.ind_excel_entry, [("Excel files", "*.xlsx *.xls")]),
        ).pack(side="right")

        # Database File
        ctk.CTkLabel(form_scroll, text="Primavera SQLite Database (.db):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        f2 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f2.pack(fill="x", padx=15)
        self.ind_db_entry = ctk.CTkEntry(f2, placeholder_text="Select SQLite db...")
        self.ind_db_entry.insert(0, DEFAULT_DB_PATH)
        self.ind_db_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f2, text="Browse", width=80, command=lambda: self.browse_file(self.ind_db_entry, [("SQLite DB", "*.db")])
        ).pack(side="right")

        # Baseline ID
        ctk.CTkLabel(form_scroll, text="Baseline Project ID (PROJ_SHORT_NAME):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        self.ind_project_entry = ctk.CTkEntry(form_scroll, placeholder_text="e.g. MAHSR T3-CWP-R6-1")
        self.ind_project_entry.pack(fill="x", padx=15)

        # Reported Activities / Milestone Selection Mode
        ctk.CTkLabel(form_scroll, text="Reported Activities / Milestone Mode:", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        self.ind_mahsr_var = ctk.StringVar(value="standard")
        f3 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f3.pack(fill="x", padx=15)
        ctk.CTkRadioButton(
            f3, text="Use standard MS-XX pattern", variable=self.ind_mahsr_var, value="standard"
        ).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(
            f3,
            text="Import specified activities list (or custom pattern) from Excel",
            variable=self.ind_mahsr_var,
            value="custom",
        ).pack(side="left")

        f3b = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f3b.pack(fill="x", padx=15, pady=(10, 0))
        self.ind_ms_excel_entry = ctk.CTkEntry(
            f3b, placeholder_text="Select specified activities list Excel file (or custom pattern Excel)..."
        )
        self.ind_ms_excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f3b,
            text="Browse",
            width=80,
            command=lambda: self.browse_file(
                self.ind_ms_excel_entry,
                [("Excel files", "*.xlsx *.xls")],
                on_selected=lambda: self.ind_mahsr_var.set("custom"),
            ),
        ).pack(side="right")

        # Default Remarks
        ctk.CTkLabel(form_scroll, text="Default Remarks (optional):", anchor="w").pack(fill="x", padx=15, pady=(15, 0))
        self.ind_remarks_entry = ctk.CTkEntry(
            form_scroll, placeholder_text="Optional remarks applied to all created delay activities"
        )
        self.ind_remarks_entry.pack(fill="x", padx=15, pady=(0, 15))

        # Launch controls
        ind_btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        ind_btn_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        ind_btn_frame.grid_columnconfigure(0, weight=4)
        ind_btn_frame.grid_columnconfigure(1, weight=1)

        self.ind_run_btn = ctk.CTkButton(
            ind_btn_frame,
            text="Run Individual EOT Analysis",
            height=40,
            font=("Segoe UI", 14, "bold"),
            fg_color=("#28A745", "#218838"),
            hover_color=("#218838", "#1E7E34"),
            command=self.launch_individual_eot,
        )
        self.ind_run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.ind_reset_btn = ctk.CTkButton(
            ind_btn_frame,
            text="Reset Form",
            height=40,
            font=("Segoe UI", 12),
            fg_color=("#3B8ED0", "#1F6AA5"),
            hover_color=("#2C72A4", "#18507C"),
            text_color="#FFFFFF",
            command=self.reset_individual_form,
        )
        self.ind_reset_btn.grid(row=0, column=1, sticky="ew")

        self.ind_log_box, self.ind_stdin, self.ind_kill_btn = self.create_terminal_widget(frame, row=3)

    def create_comb_frame(self):
        frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frames["combined"] = frame

        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Combined EOT Analysis", font=("Segoe UI", 20, "bold"), anchor="w").grid(
            row=0, column=0, sticky="ew", pady=(0, 15)
        )

        form_scroll = ctk.CTkScrollableFrame(frame)
        form_scroll.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        form_scroll.grid_columnconfigure(0, weight=1)

        # Delay Events Excel File
        ctk.CTkLabel(form_scroll, text="Delay Events Excel File:", anchor="w").pack(fill="x", padx=15, pady=(10, 0))
        f1 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f1.pack(fill="x", padx=15)
        self.comb_excel_entry = ctk.CTkEntry(f1, placeholder_text="Select .xlsx file...")
        self.comb_excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f1,
            text="Browse",
            width=80,
            command=lambda: self.browse_file(self.comb_excel_entry, [("Excel files", "*.xlsx *.xls")]),
        ).pack(side="right")

        # Database File
        ctk.CTkLabel(form_scroll, text="Primavera SQLite Database (.db):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        f2 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f2.pack(fill="x", padx=15)
        self.comb_db_entry = ctk.CTkEntry(f2, placeholder_text="Select SQLite db...")
        self.comb_db_entry.insert(0, DEFAULT_DB_PATH)
        self.comb_db_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f2, text="Browse", width=80, command=lambda: self.browse_file(self.comb_db_entry, [("SQLite DB", "*.db")])
        ).pack(side="right")

        # Baseline ID
        ctk.CTkLabel(form_scroll, text="Baseline Project ID (PROJ_SHORT_NAME):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        self.comb_project_entry = ctk.CTkEntry(form_scroll, placeholder_text="e.g. MAHSR T3-CWP-R6-1")
        self.comb_project_entry.pack(fill="x", padx=15)

        # Reported Activities / Milestone Selection Mode
        ctk.CTkLabel(form_scroll, text="Reported Activities / Milestone Mode:", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        self.comb_mahsr_var = ctk.StringVar(value="standard")
        f3 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f3.pack(fill="x", padx=15)
        ctk.CTkRadioButton(
            f3, text="Use standard MS-XX pattern", variable=self.comb_mahsr_var, value="standard"
        ).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(
            f3,
            text="Import specified activities list (or custom pattern) from Excel",
            variable=self.comb_mahsr_var,
            value="custom",
        ).pack(side="left")

        f3b = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f3b.pack(fill="x", padx=15, pady=(10, 0))
        self.comb_ms_excel_entry = ctk.CTkEntry(
            f3b, placeholder_text="Select specified activities list Excel file (or custom pattern Excel)..."
        )
        self.comb_ms_excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f3b,
            text="Browse",
            width=80,
            command=lambda: self.browse_file(
                self.comb_ms_excel_entry,
                [("Excel files", "*.xlsx *.xls")],
                on_selected=lambda: self.comb_mahsr_var.set("custom"),
            ),
        ).pack(side="right")

        # Default Remarks
        ctk.CTkLabel(form_scroll, text="Default Remarks (optional):", anchor="w").pack(fill="x", padx=15, pady=(15, 0))
        self.comb_remarks_entry = ctk.CTkEntry(
            form_scroll, placeholder_text="Optional remarks applied to all created delay activities"
        )
        self.comb_remarks_entry.pack(fill="x", padx=15, pady=(0, 15))

        # Launch controls
        comb_btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        comb_btn_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        comb_btn_frame.grid_columnconfigure(0, weight=4)
        comb_btn_frame.grid_columnconfigure(1, weight=1)

        self.comb_run_btn = ctk.CTkButton(
            comb_btn_frame,
            text="Run Combined EOT Analysis",
            height=40,
            font=("Segoe UI", 14, "bold"),
            fg_color=("#28A745", "#218838"),
            hover_color=("#218838", "#1E7E34"),
            command=self.launch_combined_eot,
        )
        self.comb_run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.comb_reset_btn = ctk.CTkButton(
            comb_btn_frame,
            text="Reset Form",
            height=40,
            font=("Segoe UI", 12),
            fg_color=("#3B8ED0", "#1F6AA5"),
            hover_color=("#2C72A4", "#18507C"),
            text_color="#FFFFFF",
            command=self.reset_combined_form,
        )
        self.comb_reset_btn.grid(row=0, column=1, sticky="ew")

        self.comb_log_box, self.comb_stdin, self.comb_kill_btn = self.create_terminal_widget(frame, row=3)

    def create_export_frame(self):
        frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frames["export"] = frame

        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Project Export Utility", font=("Segoe UI", 20, "bold"), anchor="w").grid(
            row=0, column=0, sticky="ew", pady=(0, 15)
        )

        form_scroll = ctk.CTkScrollableFrame(frame)
        form_scroll.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        form_scroll.grid_columnconfigure(0, weight=1)

        # Database File
        ctk.CTkLabel(form_scroll, text="Primavera SQLite Database (.db):", anchor="w").pack(
            fill="x", padx=15, pady=(10, 0)
        )
        f2 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f2.pack(fill="x", padx=15)
        self.exp_db_entry = ctk.CTkEntry(f2, placeholder_text="Select SQLite db...")
        self.exp_db_entry.insert(0, DEFAULT_DB_PATH)
        self.exp_db_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f2, text="Browse", width=80, command=lambda: self.browse_file(self.exp_db_entry, [("SQLite DB", "*.db")])
        ).pack(side="right")

        # Baseline ID
        ctk.CTkLabel(form_scroll, text="Baseline Project ID (PROJ_SHORT_NAME):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        self.exp_project_entry = ctk.CTkEntry(form_scroll, placeholder_text="e.g. MAHSR T3-CWP-R6-1")
        self.exp_project_entry.pack(fill="x", padx=15)

        # Log Path File
        ctk.CTkLabel(form_scroll, text="Log File Path (optional, for export mode 1):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        f1 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f1.pack(fill="x", padx=15)
        self.exp_log_entry = ctk.CTkEntry(
            f1, placeholder_text="Select or paste log file path (defaults to eot_process_logs.log)..."
        )
        self.exp_log_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f1, text="Browse", width=80, command=lambda: self.browse_file(self.exp_log_entry, [("Log files", "*.log")])
        ).pack(side="right")

        # Export Mode Radio Buttons
        ctk.CTkLabel(form_scroll, text="Export Mode:", anchor="w").pack(fill="x", padx=15, pady=(15, 0))
        self.exp_mode_var = ctk.StringVar(value="1")
        f3 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f3.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkRadioButton(
            f3, text="1. Export all delay projects scheduled in last run", variable=self.exp_mode_var, value="1"
        ).pack(anchor="w", pady=5)
        ctk.CTkRadioButton(
            f3,
            text="2. Select specific delay project serial numbers interactively",
            variable=self.exp_mode_var,
            value="2",
        ).pack(anchor="w", pady=5)
        ctk.CTkRadioButton(f3, text="3. Export all projects in database", variable=self.exp_mode_var, value="3").pack(
            anchor="w", pady=5
        )

        # Launch controls
        exp_btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        exp_btn_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        exp_btn_frame.grid_columnconfigure(0, weight=4)
        exp_btn_frame.grid_columnconfigure(1, weight=1)

        self.exp_run_btn = ctk.CTkButton(
            exp_btn_frame,
            text="Run Export Utility",
            fg_color=("#28A745", "#218838"),
            hover_color=("#218838", "#1E7E34"),
            height=40,
            font=("Segoe UI", 14, "bold"),
            command=self.launch_export_utility,
        )
        self.exp_run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.exp_reset_btn = ctk.CTkButton(
            exp_btn_frame,
            text="Reset Form",
            height=40,
            font=("Segoe UI", 12),
            fg_color=("#3B8ED0", "#1F6AA5"),
            hover_color=("#2C72A4", "#18507C"),
            text_color="#FFFFFF",
            command=self.reset_export_form,
        )
        self.exp_reset_btn.grid(row=0, column=1, sticky="ew")

        self.exp_log_box, self.exp_stdin, self.exp_kill_btn = self.create_terminal_widget(frame, row=3)

    def create_del_frame(self):
        frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frames["delete"] = frame

        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Project Deletion Utility", font=("Segoe UI", 20, "bold"), anchor="w").grid(
            row=0, column=0, sticky="ew", pady=(0, 15)
        )

        form_scroll = ctk.CTkScrollableFrame(frame)
        form_scroll.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        form_scroll.grid_columnconfigure(0, weight=1)

        # Database File
        ctk.CTkLabel(form_scroll, text="Primavera SQLite Database (.db):", anchor="w").pack(
            fill="x", padx=15, pady=(10, 0)
        )
        f2 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f2.pack(fill="x", padx=15)
        self.del_db_entry = ctk.CTkEntry(f2, placeholder_text="Select SQLite db...")
        self.del_db_entry.insert(0, DEFAULT_DB_PATH)
        self.del_db_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f2, text="Browse", width=80, command=lambda: self.browse_file(self.del_db_entry, [("SQLite DB", "*.db")])
        ).pack(side="right")

        # Baseline ID
        ctk.CTkLabel(form_scroll, text="Baseline Project ID (PROJ_SHORT_NAME):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        self.del_project_entry = ctk.CTkEntry(form_scroll, placeholder_text="e.g. MAHSR T3-CWP-R6-1")
        self.del_project_entry.pack(fill="x", padx=15)

        # Log Path File
        ctk.CTkLabel(form_scroll, text="Log File Path (optional, for deletion mode 1):", anchor="w").pack(
            fill="x", padx=15, pady=(15, 0)
        )
        f1 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f1.pack(fill="x", padx=15)
        self.del_log_entry = ctk.CTkEntry(
            f1, placeholder_text="Select or paste log file path (defaults to eot_process_logs.log)..."
        )
        self.del_log_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(
            f1, text="Browse", width=80, command=lambda: self.browse_file(self.del_log_entry, [("Log files", "*.log")])
        ).pack(side="right")

        # Deletion Mode Radio Buttons
        ctk.CTkLabel(form_scroll, text="Deletion Mode:", anchor="w").pack(fill="x", padx=15, pady=(15, 0))
        self.del_mode_var = ctk.StringVar(value="1")
        f3 = ctk.CTkFrame(form_scroll, fg_color="transparent")
        f3.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkRadioButton(f3, text="1. Delete last EOT run only", variable=self.del_mode_var, value="1").pack(
            anchor="w", pady=5
        )
        ctk.CTkRadioButton(
            f3,
            text="2. Select specific delay project serial numbers interactively",
            variable=self.del_mode_var,
            value="2",
        ).pack(anchor="w", pady=5)
        ctk.CTkRadioButton(
            f3,
            text="3. Scan & delete ALL delay projects (matching _DE, _VE, _EOT)",
            variable=self.del_mode_var,
            value="3",
        ).pack(anchor="w", pady=5)

        # Launch controls
        del_btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        del_btn_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        del_btn_frame.grid_columnconfigure(0, weight=4)
        del_btn_frame.grid_columnconfigure(1, weight=1)

        self.del_run_btn = ctk.CTkButton(
            del_btn_frame,
            text="Run Deletion Utility",
            fg_color=("#28A745", "#218838"),
            hover_color=("#218838", "#1E7E34"),
            height=40,
            font=("Segoe UI", 14, "bold"),
            command=self.launch_deletion_utility,
        )
        self.del_run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.del_reset_btn = ctk.CTkButton(
            del_btn_frame,
            text="Reset Form",
            height=40,
            font=("Segoe UI", 12),
            fg_color=("#3B8ED0", "#1F6AA5"),
            hover_color=("#2C72A4", "#18507C"),
            text_color="#FFFFFF",
            command=self.reset_delete_form,
        )
        self.del_reset_btn.grid(row=0, column=1, sticky="ew")

        self.del_log_box, self.del_stdin, self.del_kill_btn = self.create_terminal_widget(frame, row=3)

    def create_terminal_widget(self, parent_frame, row):
        terminal_frame = ctk.CTkFrame(
            parent_frame, corner_radius=10, border_width=1, border_color=("#CCCCCC", "#333333")
        )
        terminal_frame.grid(row=row, column=0, sticky="nsew", pady=(5, 0))
        parent_frame.grid_rowconfigure(row, weight=1)

        terminal_frame.grid_rowconfigure(1, weight=1)
        terminal_frame.grid_columnconfigure(0, weight=1)

        header_bar = ctk.CTkFrame(terminal_frame, height=30, fg_color=("#EAEAEA", "#1E1E1E"))
        header_bar.grid(row=0, column=0, sticky="ew")
        header_bar.grid_propagate(False)

        ctk.CTkLabel(
            header_bar, text="Console Output Logs", font=("Segoe UI", 11, "bold"), text_color=("#666666", "#888888")
        ).pack(side="left", padx=10)

        kill_btn = ctk.CTkButton(
            header_bar,
            text="Abort / Kill Process",
            font=("Segoe UI", 10, "bold"),
            fg_color="#8F1D1D",
            hover_color="#6F1212",
            width=120,
            height=20,
            corner_radius=4,
            command=self.kill_active_process,
        )
        kill_btn.pack(side="right", padx=10, pady=5)

        log_box = ctk.CTkTextbox(
            terminal_frame,
            font=("Consolas", 12),
            fg_color=("#F5F5F5", "#121212"),
            text_color=("#1A1A1A", "#E0E0E0"),
        )
        log_box.grid(row=1, column=0, sticky="nsew", padx=5, pady=(5, 0))

        input_bar = ctk.CTkFrame(terminal_frame, height=35, fg_color="transparent")
        input_bar.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        input_bar.grid_propagate(False)

        stdin_entry = ctk.CTkEntry(input_bar, placeholder_text="Type interactive command / response here...")
        stdin_entry.pack(side="left", fill="both", expand=True, padx=(0, 5))
        stdin_entry.bind("<Return>", lambda event: self.send_stdin_from_box())

        send_btn = ctk.CTkButton(input_bar, text="Send Input", width=80, command=self.send_stdin_from_box)
        send_btn.pack(side="right", fill="y")

        return log_box, stdin_entry, kill_btn

    # ─────────────────────────────────────────────────────────────────────────────
    # INTERACTION LOGIC
    # ─────────────────────────────────────────────────────────────────────────────

    @property
    def active_stdin_entry(self):
        if self.current_page == "individual":
            return self.ind_stdin
        elif self.current_page == "combined":
            return self.comb_stdin
        elif self.current_page == "export":
            return self.exp_stdin
        elif self.current_page == "delete":
            return self.del_stdin
        return None

    def browse_file(self, entry, filetypes, on_selected=None):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)
            if callable(on_selected):
                on_selected()

    def browse_folder(self, entry):
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def write_log(self, text, log_box, append_newline=True):
        if not log_box:
            return
        log_box.configure(state="normal")
        log_box.insert("end", text)
        if append_newline:
            log_box.insert("end", "\n")
        log_box.configure(state="disabled")
        log_box.see("end")

    def thread_safe_log(self, char, log_box):
        self.after(0, self.write_log, char, log_box, False)

    def send_stdin_from_box(self):
        entry = self.active_stdin_entry
        if not entry:
            return
        text = entry.get().strip()
        if self.runner and self.runner.is_running:
            self.runner.send_input(text)
            entry.delete(0, "end")

    def kill_active_process(self):
        if self.runner and self.runner.is_running:
            log_box = getattr(self.runner, "log_box", None)
            self.runner.kill()
            if log_box:
                self.write_log("\n[Process terminated by user]", log_box)

    def start_script_execution(self, cmd, log_box, auto_responses=None):
        if log_box:
            log_box.configure(state="normal")
            log_box.delete("1.0", "end")
            log_box.configure(state="disabled")

        for btn_name in [
            "ind_run_btn",
            "comb_run_btn",
            "exp_run_btn",
            "del_run_btn",
            "ind_reset_btn",
            "comb_reset_btn",
            "exp_reset_btn",
            "del_reset_btn",
        ]:
            if hasattr(self, btn_name):
                getattr(self, btn_name).configure(state="disabled")

        self.runner = ProcessRunner(
            cmd=cmd,
            on_output_func=lambda char: self.thread_safe_log(char, log_box),
            on_finished_func=self.on_process_finished,
            auto_responses=auto_responses,
        )
        self.runner.log_box = log_box
        self.runner.start()

    def on_process_finished(self):
        for btn_name in [
            "ind_run_btn",
            "comb_run_btn",
            "exp_run_btn",
            "del_run_btn",
            "ind_reset_btn",
            "comb_reset_btn",
            "exp_reset_btn",
            "del_reset_btn",
        ]:
            if hasattr(self, btn_name):
                btn = getattr(self, btn_name)
                self.after(0, lambda b=btn: b.configure(state="normal"))

        if self.runner:
            rc = self.runner.process.poll() if self.runner.process else None
            log_box = getattr(self.runner, "log_box", None)

            discrepancies = None
            log_text = ""
            if log_box:
                log_text = log_box.get("1.0", "end")
                discrepancies = self.extract_discrepancies(log_text)

            if self.current_page in ("export", "delete"):
                has_error_keyword = False
                if log_text:
                    lower_log = log_text.lower()
                    error_keywords = [
                        "no valid selections made",
                        "no matching projects found",
                        "no delay/eot projects found",
                        "no eot run sessions found",
                        "export cancelled",
                        "deletion cancelled",
                        "invalid choice",
                        "[error]",
                        "traceback",
                    ]
                    has_error_keyword = any(k in lower_log for k in error_keywords)

                if rc == 0 and not has_error_keyword:
                    self.after(
                        0, lambda: messagebox.showinfo("Process Status", "The process is completed successfully.")
                    )
                elif rc is not None and rc != -9:
                    self.after(0, lambda: messagebox.showerror("Process Status", "Unexpected error occurred."))
                elif rc == 0 and has_error_keyword:
                    self.after(0, lambda: messagebox.showerror("Process Status", "Unexpected error occurred."))
            else:
                if rc == 0:
                    if discrepancies:
                        self.after(0, lambda: self.show_discrepancies_popup(discrepancies))
                    else:
                        self.after(
                            0, lambda: messagebox.showinfo("Process Status", "The process is completed successfully.")
                        )
                elif rc is not None and rc != -9:
                    self.after(
                        0,
                        lambda: messagebox.showerror(
                            "Process Status", f"Error occurred, please check console. Exit code: {rc}"
                        ),
                    )

    def extract_discrepancies(self, log_text):
        pattern = r"SUMMARY OF SKIPPED ROWS WITH DISCREPANCIES\s*={5,}\s*(.*?)\s*={5,}"
        match = re.search(pattern, log_text, re.DOTALL)
        if match:
            content = match.group(1).strip()
            if "Total rows skipped: 0" in content or "No rows were skipped" in content:
                return None
            return content
        return None

    def show_discrepancies_popup(self, discrepancies_text):
        dialog_win = ctk.CTkToplevel(self)
        dialog_win.title("Process Status - Discrepancies")
        dialog_win.geometry("600x420")
        dialog_win.minsize(500, 350)

        dialog_win.focus_set()
        dialog_win.grab_set()
        dialog_win.attributes("-topmost", True)

        dialog_win.grid_rowconfigure(2, weight=1)
        dialog_win.grid_columnconfigure(0, weight=1)

        header_lbl = ctk.CTkLabel(
            dialog_win,
            text="The process completed with some discrepancies/skipped rows.",
            font=("Segoe UI", 14, "bold"),
            text_color="#E65100",
            anchor="w",
            justify="left",
        )
        header_lbl.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))

        inst_lbl = ctk.CTkLabel(
            dialog_win,
            text="Skipped Rows (you can select and copy the text below):",
            font=("Segoe UI", 11, "italic"),
            text_color="gray",
            anchor="w",
        )
        inst_lbl.grid(row=1, column=0, sticky="ew", padx=20, pady=(5, 5))

        text_box = ctk.CTkTextbox(
            dialog_win,
            font=("Consolas", 12),
            fg_color=("white", "#1C1C1C"),
            text_color=("black", "white"),
        )
        text_box.grid(row=2, column=0, sticky="nsew", padx=20, pady=(5, 15))
        text_box.insert("1.0", discrepancies_text)
        text_box.configure(state="disabled")

        btn_frame = ctk.CTkFrame(dialog_win, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 15))
        btn_frame.grid_columnconfigure(0, weight=1)

        close_btn = ctk.CTkButton(
            btn_frame,
            text="Exit",
            fg_color="#28A745",
            hover_color="#218838",
            command=dialog_win.destroy,
            font=("Segoe UI", 11, "bold"),
        )
        close_btn.pack(pady=5)

        dialog_win.update_idletasks()
        width = dialog_win.winfo_width()
        height = dialog_win.winfo_height()
        x = (dialog_win.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog_win.winfo_screenheight() // 2) - (height // 2)
        dialog_win.geometry(f"+{x}+{y}")

        dialog_win.protocol("WM_DELETE_WINDOW", dialog_win.destroy)

    def reset_individual_form(self):
        self.ind_excel_entry.delete(0, "end")
        self.ind_db_entry.delete(0, "end")
        self.ind_db_entry.insert(0, DEFAULT_DB_PATH)
        self.ind_project_entry.delete(0, "end")
        self.ind_ms_excel_entry.delete(0, "end")
        self.ind_remarks_entry.delete(0, "end")
        self.ind_mahsr_var.set("standard")

        self.ind_log_box.configure(state="normal")
        self.ind_log_box.delete("1.0", "end")
        self.ind_log_box.configure(state="disabled")

    def reset_combined_form(self):
        self.comb_excel_entry.delete(0, "end")
        self.comb_db_entry.delete(0, "end")
        self.comb_db_entry.insert(0, DEFAULT_DB_PATH)
        self.comb_project_entry.delete(0, "end")
        if hasattr(self, "comb_ms_excel_entry"):
            self.comb_ms_excel_entry.delete(0, "end")
        if hasattr(self, "comb_mahsr_var"):
            self.comb_mahsr_var.set("standard")
        self.comb_remarks_entry.delete(0, "end")

        self.comb_log_box.configure(state="normal")
        self.comb_log_box.delete("1.0", "end")
        self.comb_log_box.configure(state="disabled")

    def reset_export_form(self):
        self.exp_db_entry.delete(0, "end")
        self.exp_db_entry.insert(0, DEFAULT_DB_PATH)
        self.exp_project_entry.delete(0, "end")
        self.exp_log_entry.delete(0, "end")
        self.exp_mode_var.set("1")

        self.exp_log_box.configure(state="normal")
        self.exp_log_box.delete("1.0", "end")
        self.exp_log_box.configure(state="disabled")

    def reset_delete_form(self):
        self.del_db_entry.delete(0, "end")
        self.del_db_entry.insert(0, DEFAULT_DB_PATH)
        self.del_project_entry.delete(0, "end")
        self.del_log_entry.delete(0, "end")
        self.del_mode_var.set("1")

        self.del_log_box.configure(state="normal")
        self.del_log_box.delete("1.0", "end")
        self.del_log_box.configure(state="disabled")

    # ─────────────────────────────────────────────────────────────────────────────
    # LAUNCHERS
    # ─────────────────────────────────────────────────────────────────────────────

    def launch_individual_eot(self):
        excel_path = self.ind_excel_entry.get().strip().strip("'\"")
        db_path = self.ind_db_entry.get().strip().strip("'\"")
        source_project = self.ind_project_entry.get().strip()
        remarks = self.ind_remarks_entry.get().strip()
        milestone_mode = self.ind_mahsr_var.get()
        ms_config = self.ind_ms_excel_entry.get().strip().strip("'\"")

        if ms_config and milestone_mode != "custom":
            milestone_mode = "custom"
            self.ind_mahsr_var.set("custom")

        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return
        if not db_path or not os.path.exists(db_path):
            messagebox.showerror("Error", "Please select a valid database (.db) file.")
            return
        if not source_project:
            messagebox.showerror("Error", "Please enter the Baseline Project ID.")
            return
        if milestone_mode == "custom" and (not ms_config or not os.path.exists(ms_config)):
            messagebox.showerror(
                "Error", "Please select a valid Specified Activities / Custom Milestone Excel file."
            )
            return

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
            "--milestone-mode",
            milestone_mode,
        ]
        if milestone_mode == "custom" and ms_config:
            cmd.extend(["--milestone-config", ms_config])
        if remarks:
            cmd.extend(["--default-remarks", remarks])

        auto_responses = {
            "Enter the Primavera SQLite database path (default:": db_path,
            "Enter the baseline project ID (PROJ_SHORT_NAME):": source_project,
            "Enter the path of the delay events Excel file:": excel_path,
            "Enter path to the specified activities / custom milestone Excel file:": ms_config,
        }

        self.start_script_execution(cmd, self.ind_log_box, auto_responses)

    def launch_combined_eot(self):
        excel_path = self.comb_excel_entry.get().strip().strip("'\"")
        db_path = self.comb_db_entry.get().strip().strip("'\"")
        source_project = self.comb_project_entry.get().strip()
        remarks = self.comb_remarks_entry.get().strip()
        milestone_mode = self.comb_mahsr_var.get() if hasattr(self, "comb_mahsr_var") else "standard"
        ms_config = (
            self.comb_ms_excel_entry.get().strip().strip("'\"") if hasattr(self, "comb_ms_excel_entry") else ""
        )

        if ms_config and milestone_mode != "custom":
            milestone_mode = "custom"
            if hasattr(self, "comb_mahsr_var"):
                self.comb_mahsr_var.set("custom")

        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return
        if not db_path or not os.path.exists(db_path):
            messagebox.showerror("Error", "Please select a valid database (.db) file.")
            return
        if not source_project:
            messagebox.showerror("Error", "Please enter the Baseline Project ID.")
            return
        if milestone_mode == "custom" and (not ms_config or not os.path.exists(ms_config)):
            messagebox.showerror(
                "Error", "Please select a valid Specified Activities / Custom Milestone Excel file."
            )
            return

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
            "--milestone-mode",
            milestone_mode,
        ]
        if milestone_mode == "custom" and ms_config:
            cmd.extend(["--milestone-config", ms_config])
        if remarks:
            cmd.extend(["--default-remarks", remarks])

        auto_responses = {
            "Enter the Primavera SQLite database path (default:": db_path,
            "Enter the baseline project ID (PROJ_SHORT_NAME):": source_project,
            "Enter the path of the delay events Excel file:": excel_path,
            "Enter path to the specified activities / custom milestone Excel file:": ms_config,
        }
        self.start_script_execution(cmd, self.comb_log_box, auto_responses)

    def launch_export_utility(self):
        db_path = self.exp_db_entry.get().strip().strip("'\"")
        source_project = self.exp_project_entry.get().strip()
        log_path = self.exp_log_entry.get().strip().strip("'\"")
        choice = self.exp_mode_var.get()

        if not db_path or not os.path.exists(db_path):
            messagebox.showerror("Error", "Please select a valid database (.db) file.")
            return
        if not source_project:
            messagebox.showerror("Error", "Please enter the Baseline Project ID.")
            return
        if choice == "1":
            if not log_path and os.path.exists(DEFAULT_LOG_PATH):
                log_path = DEFAULT_LOG_PATH
            if not log_path or not os.path.exists(log_path):
                messagebox.showerror(
                    "Error",
                    "Option 1 (Export projects scheduled in last run) requires a valid, existing log file. Please select or paste your log file path (.log).",
                )
                return

        cmd = [
            sys.executable,
            "-u",
            EXPORT_EOT_SCRIPT,
            "--db-path",
            db_path,
            "--project",
            source_project,
            "--mode",
            choice,
        ]
        if choice == "1" and log_path:
            cmd.extend(["--log-path", log_path])

        auto_responses = {}

        def ask_export_confirm():
            buf = self.exp_log_box.get("1.0", "end")
            lines = [line.strip() for line in buf.split("\n")]
            projects_list = []
            found_start = False
            for line in lines:
                if "project(s) to export:" in line:
                    found_start = True
                    continue
                if found_start:
                    if "Do you want to export these projects" in line:
                        break
                    if line:
                        projects_list.append(line)

            if projects_list:
                msg = "Are you sure you want to export these projects to XER format?\n\n" + "\n".join(projects_list)
            else:
                msg = "Are you sure you want to export these projects to XER format?"

            res = messagebox.askyesno("Confirm Project Export", msg)
            if res:
                self.runner.send_input("y")
            else:
                self.runner.send_input("n")

        auto_responses["Do you want to export these projects to XER format? (y/n):"] = lambda: self.after(
            0, ask_export_confirm
        )

        def ask_export_serials():
            buf = self.exp_log_box.get("1.0", "end")
            lines = [line.strip() for line in buf.split("\n")]
            projects = []
            found_start = False
            for line in lines:
                if "Found" in line and "project(s):" in line:
                    found_start = True
                    continue
                if found_start:
                    if "Enter the serial numbers" in line:
                        break
                    m = re.match(r"^(\d+)\s*\.\s*(.+)$", line)
                    if m:
                        serial = m.group(1)
                        name = m.group(2).strip()
                        projects.append((serial, name))

            if not projects:
                dialog = ctk.CTkInputDialog(
                    text="Enter the serial numbers of the projects to export (comma-separated, e.g. 1,3,4):",
                    title="Select Projects",
                )
                res = dialog.get_input()
                self.runner.send_input(res if res else "")
                return

            dialog_win = ctk.CTkToplevel(self)
            dialog_win.title("Select Projects to Export")
            dialog_win.geometry("500x500")
            dialog_win.minsize(400, 400)

            dialog_win.focus_set()
            dialog_win.grab_set()

            dialog_win.grid_rowconfigure(2, weight=1)
            dialog_win.grid_columnconfigure(0, weight=1)

            header_lbl = ctk.CTkLabel(
                dialog_win,
                text=f"Scanned Projects ({len(projects)} found)\nSelect the projects you want to export:",
                font=("Segoe UI", 13, "bold"),
                anchor="w",
                justify="left",
            )
            header_lbl.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))

            def toggle_select_all():
                any_unchecked = any(not var.get() for var in checkbox_vars.values())
                new_val = True if any_unchecked else False
                for var in checkbox_vars.values():
                    var.set(new_val)
                select_all_btn.configure(text="Deselect All" if new_val else "Select All")

            select_all_btn = ctk.CTkButton(
                dialog_win,
                text="Select All",
                width=120,
                height=28,
                font=("Segoe UI", 11),
                command=toggle_select_all,
            )
            select_all_btn.grid(row=1, column=0, sticky="w", padx=20, pady=(5, 10))

            scroll_frame = ctk.CTkScrollableFrame(dialog_win)
            scroll_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 15))
            scroll_frame.grid_columnconfigure(0, weight=1)

            checkbox_vars = {}
            for idx, (serial, name) in enumerate(projects):
                var = ctk.BooleanVar(value=False)
                checkbox_vars[serial] = var
                cb = ctk.CTkCheckBox(
                    scroll_frame,
                    text=f"{serial}. {name}",
                    variable=var,
                    font=("Segoe UI", 12),
                )
                cb.pack(anchor="w", fill="x", padx=10, pady=6)

            def on_submit():
                selections = [serial for serial, var in checkbox_vars.items() if var.get()]
                if not selections:
                    messagebox.showwarning(
                        "No Selection",
                        "Please select at least one project to export, or click Cancel.",
                        parent=dialog_win,
                    )
                    return
                self.runner.send_input(",".join(selections))
                dialog_win.destroy()

            def on_cancel():
                self.runner.send_input("")
                dialog_win.destroy()

            btn_frame = ctk.CTkFrame(dialog_win, fg_color="transparent")
            btn_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 15))
            btn_frame.grid_columnconfigure((0, 1), weight=1)

            cancel_btn = ctk.CTkButton(
                btn_frame,
                text="Cancel",
                fg_color="#333333",
                hover_color="#444444",
                command=on_cancel,
            )
            cancel_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

            export_btn = ctk.CTkButton(
                btn_frame,
                text="Export Selected",
                fg_color="#28A745",
                hover_color="#218838",
                command=on_submit,
            )
            export_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

            dialog_win.protocol("WM_DELETE_WINDOW", on_cancel)

        auto_responses["Enter the serial numbers of the projects to export (comma-separated, e.g. 1,3,4):"] = (
            lambda: self.after(0, ask_export_serials)
        )

        def ask_output_directory():
            path = filedialog.askdirectory(title="Select Output Directory for XER Files")
            self.runner.send_input(path if path else "")

        auto_responses["Output Directory:"] = lambda: self.after(0, ask_output_directory)

        self.start_script_execution(cmd, self.exp_log_box, auto_responses)

    def launch_deletion_utility(self):
        db_path = self.del_db_entry.get().strip().strip("'\"")
        source_project = self.del_project_entry.get().strip()
        log_path = self.del_log_entry.get().strip().strip("'\"")
        choice = self.del_mode_var.get()

        if not db_path or not os.path.exists(db_path):
            messagebox.showerror("Error", "Please select a valid database (.db) file.")
            return
        if not source_project:
            messagebox.showerror("Error", "Please enter the Baseline Project ID.")
            return
        if choice == "1":
            if not log_path and os.path.exists(DEFAULT_LOG_PATH):
                log_path = DEFAULT_LOG_PATH
            if not log_path or not os.path.exists(log_path):
                messagebox.showerror(
                    "Error",
                    "Option 1 (Delete last EOT run only) requires a valid, existing log file. Please select or paste your log file path (.log).",
                )
                return

        cmd = [sys.executable, "-u", DELETE_PROJECTS_SCRIPT]

        auto_responses = {
            "Enter the Primavera SQLite database path (default:": db_path if db_path else "",
            "Log File Path:": log_path if log_path else "",
            "Enter the baseline project ID (PROJ_SHORT_NAME):": source_project,
            "Enter choice (1, 2, or 3):": choice,
        }

        def ask_user_confirm():
            buf = self.del_log_box.get("1.0", "end")
            lines = [line.strip() for line in buf.split("\n")]
            projects_list = []
            found_start = False
            for line in lines:
                if "project(s) to delete in database:" in line:
                    found_start = True
                    continue
                if found_start:
                    if "Do you want to permanently delete" in line:
                        break
                    if line:
                        projects_list.append(line)

            if projects_list:
                msg = "Are you sure you want to permanently delete these projects?\n\n" + "\n".join(projects_list)
            else:
                msg = "Are you sure you want to permanently delete these projects?"

            res = messagebox.askyesno("Confirm Project Deletion", msg)
            if res:
                self.runner.send_input("y")
            else:
                self.runner.send_input("n")

        auto_responses["Do you want to permanently delete these projects from the database? (y/n):"] = (
            lambda: self.after(0, ask_user_confirm)
        )

        def ask_serials():
            buf = self.del_log_box.get("1.0", "end")
            lines = [line.strip() for line in buf.split("\n")]
            projects = []
            found_start = False
            for line in lines:
                if "Found" in line and "delay project(s)" in line and "associated with" in line:
                    found_start = True
                    continue
                if found_start:
                    if "Enter the serial numbers" in line:
                        break
                    m = re.match(r"^(\d+)\s*\.\s*(.+)$", line)
                    if m:
                        serial = m.group(1)
                        name = m.group(2).strip()
                        projects.append((serial, name))

            if not projects:
                dialog = ctk.CTkInputDialog(
                    text="Enter the serial numbers of the projects you want to delete (comma-separated, e.g. 1,3,4):",
                    title="Select Projects",
                )
                res = dialog.get_input()
                self.runner.send_input(res if res else "")
                return

            dialog_win = ctk.CTkToplevel(self)
            dialog_win.title("Select Projects to Delete")
            dialog_win.geometry("500x500")
            dialog_win.minsize(400, 400)

            dialog_win.focus_set()
            dialog_win.grab_set()

            dialog_win.grid_rowconfigure(2, weight=1)
            dialog_win.grid_columnconfigure(0, weight=1)

            header_lbl = ctk.CTkLabel(
                dialog_win,
                text=f"Scanned Projects ({len(projects)} found)\nSelect the projects you want to delete:",
                font=("Segoe UI", 13, "bold"),
                anchor="w",
                justify="left",
            )
            header_lbl.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))

            def toggle_select_all():
                any_unchecked = any(not var.get() for var in checkbox_vars.values())
                new_val = True if any_unchecked else False
                for var in checkbox_vars.values():
                    var.set(new_val)
                select_all_btn.configure(text="Deselect All" if new_val else "Select All")

            select_all_btn = ctk.CTkButton(
                dialog_win,
                text="Select All",
                width=120,
                height=28,
                font=("Segoe UI", 11),
                command=toggle_select_all,
            )
            select_all_btn.grid(row=1, column=0, sticky="w", padx=20, pady=(5, 10))

            scroll_frame = ctk.CTkScrollableFrame(dialog_win)
            scroll_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 15))
            scroll_frame.grid_columnconfigure(0, weight=1)

            checkbox_vars = {}
            for idx, (serial, name) in enumerate(projects):
                var = ctk.BooleanVar(value=False)
                checkbox_vars[serial] = var
                cb = ctk.CTkCheckBox(
                    scroll_frame,
                    text=f"{serial}. {name}",
                    variable=var,
                    font=("Segoe UI", 12),
                )
                cb.pack(anchor="w", fill="x", padx=10, pady=6)

            def on_submit():
                selections = [serial for serial, var in checkbox_vars.items() if var.get()]
                if not selections:
                    messagebox.showwarning(
                        "No Selection",
                        "Please select at least one project to delete, or click Cancel.",
                        parent=dialog_win,
                    )
                    return
                self.runner.send_input(",".join(selections))
                dialog_win.destroy()

            def on_cancel():
                self.runner.send_input("")
                dialog_win.destroy()

            btn_frame = ctk.CTkFrame(dialog_win, fg_color="transparent")
            btn_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 15))
            btn_frame.grid_columnconfigure((0, 1), weight=1)

            cancel_btn = ctk.CTkButton(
                btn_frame,
                text="Cancel",
                fg_color="#333333",
                hover_color="#444444",
                command=on_cancel,
            )
            cancel_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

            delete_btn = ctk.CTkButton(
                btn_frame,
                text="Delete Selected",
                fg_color="#A12222",
                hover_color="#801a1a",
                command=on_submit,
            )
            delete_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

            dialog_win.protocol("WM_DELETE_WINDOW", on_cancel)

        auto_responses["Enter the serial numbers of the projects you want to delete (comma-separated, e.g. 1,3,4):"] = (
            lambda: self.after(0, ask_serials)
        )

        self.start_script_execution(cmd, self.del_log_box, auto_responses)


if __name__ == "__main__":
    app = App()
    app.mainloop()
