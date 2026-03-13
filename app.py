import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from drive_check import check_drive
from engine import AttributeConflictError, get_device_id_exclusion_status, run_pipeline
from version import APP_NAME, BUILD, VERSION

CONFIG_FILE = os.path.expanduser("~/.stardraw_tool_config.json")
APPROVED_ROOT = "TechTeam Resources"
LOGO_FILE = "Twisted_Large_Logo Positive.png"
WINDOWS_ICON_FILE = "Twisted_Icon.ico"

BG = "#0c1117"
FG = "#8e979e"
MUTED = "#94a0ad"
ACCENT = "#ff6a1a"
ACCENT_DARK = "#d95714"
TEAL = "#7ec6bd"
PANEL = "#141b24"
PANEL_ALT = "#1a2430"
BORDER = "#2a3847"
FIELD_BG = "#0b1219"
LOG_BG = "#081018"
SECONDARY = "#d6deea"
SECONDARY_HOVER = "#bec9d7"
ERROR = "#ff7b72"
SUCCESS = "#67d5a2"
WARNING = "#ffd166"


def get_resource_path(filename):
    """Resolve bundled resources correctly in both source and packaged app modes."""
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return os.path.join(bundle_dir, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


def load_config():
    """Load saved user preferences, such as the last destination and archive paths."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(data):
    """Persist user preferences so repeat runs need less manual input."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)


class App:
    """Tkinter desktop app for selecting files, running the update, and showing progress."""

    def __init__(self, root):
        """Initialize the main window, style system, and background task plumbing."""
        self.root = root
        self.root.title(f"{APP_NAME} v{VERSION}")
        self.root.geometry("980x700")
        self.root.minsize(900, 640)
        self.root.configure(bg=BG)
        self.apply_window_icon()

        self.config_data = load_config()
        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.file_entries = []
        self.browse_buttons = []
        self.log_lines = []
        self.pending_conflicts = []
        self.pending_source = None
        self.pending_dest = None
        self.pending_archive_dir = None
        self.conflict_window = None

        self.lock_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready. Select the source export and destination workbook.")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_text_var = tk.StringVar(value="Idle")
        self.logo_image = None

        self.build_ui()
        self.refresh_exclusion_status()
        self.root.after(100, self.process_log_queue)

    def apply_window_icon(self):
        """Apply the Windows ICO file to the app window when available."""
        icon_path = get_resource_path(WINDOWS_ICON_FILE)
        if os.name != "nt" or not os.path.exists(icon_path):
            return

        try:
            self.root.iconbitmap(icon_path)
        except tk.TclError:
            pass

    def build_ui(self):
        """Create the main GUI layout, controls, status banner, and scrolling log area."""
        shell = tk.Frame(self.root, bg=BG)
        shell.pack(fill="both", expand=True)

        self.outer_canvas = tk.Canvas(
            shell,
            bg=BG,
            highlightthickness=0,
            bd=0,
        )
        self.outer_canvas.pack(side="left", fill="both", expand=True)

        outer_scrollbar = tk.Scrollbar(shell, orient="vertical", command=self.outer_canvas.yview)
        outer_scrollbar.pack(side="right", fill="y")
        self.outer_canvas.configure(yscrollcommand=outer_scrollbar.set)

        outer = tk.Frame(self.outer_canvas, bg=BG)
        self.outer_canvas_window = self.outer_canvas.create_window((0, 0), window=outer, anchor="nw")

        outer.bind("<Configure>", self.on_outer_frame_configure)
        self.outer_canvas.bind("<Configure>", self.on_outer_canvas_configure)
        self.outer_canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        hero = tk.Frame(outer, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        hero.pack(fill="x", padx=18, pady=(18, 14))

        hero_band = tk.Frame(hero, bg=ACCENT, height=6)
        hero_band.pack(fill="x")

        hero_top = tk.Frame(hero, bg=PANEL)
        hero_top.pack(fill="x", padx=16, pady=(14, 6))

        logo_label = self.create_logo_label(hero_top)
        if logo_label:
            logo_label.pack(side="left", anchor="n", padx=(0, 14))

        title_block = tk.Frame(hero_top, bg=PANEL)
        title_block.pack(side="left", fill="x", expand=True)

        eyebrow = tk.Label(
            title_block,
            text="STARDRAW ATTRIBUTE WORKFLOW",
            bg=PANEL,
            fg=TEAL,
            font=("Avenir Next", 9, "bold"),
            anchor="w",
        )
        eyebrow.pack(fill="x", pady=(2, 2))

        title = tk.Label(
            title_block,
            text=APP_NAME,
            bg=PANEL,
            fg=FG,
            font=("Avenir Next", 20, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 2))

        subtitle = tk.Label(
            title_block,
            text=f"Version {VERSION}  |  Build {BUILD}",
            bg=PANEL,
            fg=ACCENT,
            font=("Avenir Next", 10, "bold"),
            anchor="w",
        )
        subtitle.pack(fill="x", pady=(0, 2))

        intro = tk.Label(
            hero,
            text="Updates Stardraw model attributes from the source product export and writes safely back to the shared workbook.",
            bg=PANEL,
            fg=MUTED,
            font=("Avenir Next", 11),
            anchor="w",
            justify="left",
            wraplength=760,
        )
        intro.pack(fill="x", padx=16, pady=(0, 14))

        hero_meta = tk.Frame(hero, bg=PANEL)
        hero_meta.pack(fill="x", padx=16, pady=(0, 14))

        for label_text, fill, text_color in [
            ("Shared workbook safe", "#1f3342", "#dbe8f3"),
            ("Conflict chooser enabled", "#17342f", SUCCESS),
            ("Archive backup before replace", "#432a1c", WARNING),
        ]:
            pill = tk.Label(
                hero_meta,
                text=label_text,
                bg=fill,
                fg=text_color,
                font=("Avenir Next", 9, "bold"),
                padx=10,
                pady=5,
            )
            pill.pack(side="left", padx=(0, 8))

        form = tk.Frame(outer, bg=BG)
        form.pack(fill="x", padx=18)

        self.source_entry = self.create_file_field(form, "Source Products Export", "Choose the latest products export workbook.")
        self.dest_entry = self.create_file_field(form, "Destination Attributes File", "Choose the shared Stardraw attributes workbook.")
        self.archive_entry = self.create_folder_field(form, "Archive Backup Folder", "Choose where backup copies should be written before the live file is replaced.")

        if "destination" in self.config_data:
            self.dest_entry.insert(0, self.config_data["destination"])
        if "archive_dir" in self.config_data:
            self.archive_entry.insert(0, self.config_data["archive_dir"])

        options_panel = tk.Frame(outer, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        options_panel.pack(fill="x", padx=18, pady=(12, 10))

        options_title = tk.Label(
            options_panel,
            text="Safety Checks",
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 11, "bold"),
            anchor="w",
        )
        options_title.pack(fill="x", padx=14, pady=(10, 4))

        lock_check = tk.Checkbutton(
            options_panel,
            text="Only allow destination files inside the approved TechTeam Google Drive path",
            variable=self.lock_var,
            bg=PANEL_ALT,
            fg=FG,
            activebackground=PANEL_ALT,
            activeforeground=FG,
            selectcolor=PANEL,
            font=("Avenir Next", 10),
            anchor="w",
        )
        lock_check.pack(fill="x", padx=14, pady=(0, 10))
        self.lock_check = lock_check

        controls = tk.Frame(outer, bg=BG)
        controls.pack(fill="x", padx=18, pady=(0, 10))

        self.run_btn = tk.Button(
            controls,
            text="Run Update",
            bg=ACCENT,
            fg="black",
            activebackground=ACCENT_DARK,
            activeforeground="black",
            font=("Avenir Next", 12, "bold"),
            padx=24,
            pady=9,
            relief="flat",
            command=self.start_pipeline,
            cursor="hand2",
        )
        self.run_btn.pack(side="left")

        self.status_label = tk.Label(
            controls,
            textvariable=self.status_var,
            bg=BG,
            fg=MUTED,
            font=("Avenir Next", 9, "bold"),
            anchor="w",
            justify="left",
        )
        self.status_label.pack(side="left", fill="x", expand=True, padx=(12, 0))

        progress_panel = tk.Frame(outer, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        progress_panel.pack(fill="x", padx=18, pady=(0, 10))

        progress_header = tk.Frame(progress_panel, bg=PANEL_ALT)
        progress_header.pack(fill="x", padx=14, pady=(10, 6))

        progress_title = tk.Label(
            progress_header,
            text="Run Progress",
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 11, "bold"),
            anchor="w",
        )
        progress_title.pack(side="left")

        self.progress_value_label = tk.Label(
            progress_header,
            textvariable=self.progress_text_var,
            bg=PANEL_ALT,
            fg=ACCENT,
            font=("Avenir Next", 10, "bold"),
            anchor="e",
        )
        self.progress_value_label.pack(side="right")

        style = ttk.Style()
        style.theme_use(style.theme_use())
        style.configure(
            "Stardraw.Horizontal.TProgressbar",
            troughcolor=FIELD_BG,
            background=ACCENT,
            bordercolor=PANEL_ALT,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            thickness=18,
        )

        self.progress_bar = ttk.Progressbar(
            progress_panel,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            style="Stardraw.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x", padx=14, pady=(0, 8))

        self.progress_stage_label = tk.Label(
            progress_panel,
            text="Waiting to start.",
            bg=PANEL_ALT,
            fg=MUTED,
            font=("Avenir Next", 9),
            anchor="w",
        )
        self.progress_stage_label.pack(fill="x", padx=14, pady=(0, 10))

        info_panel = tk.Frame(outer, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        info_panel.pack(fill="x", padx=18, pady=(0, 10))

        info_title = tk.Label(
            info_panel,
            text="Runtime Checks",
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 11, "bold"),
            anchor="w",
        )
        info_title.pack(fill="x", padx=14, pady=(10, 4))

        self.exclusion_label = tk.Label(
            info_panel,
            text="Checking Device ID exclusion file...",
            bg=PANEL_ALT,
            fg=MUTED,
            font=("Avenir Next", 9),
            anchor="w",
            justify="left",
            wraplength=760,
        )
        self.exclusion_label.pack(fill="x", padx=14, pady=(0, 10))

        log_panel = tk.Frame(outer, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        log_panel.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        log_header = tk.Frame(log_panel, bg=PANEL)
        log_header.pack(fill="x", padx=14, pady=(10, 6))

        log_title = tk.Label(
            log_header,
            text="Run Log",
            bg=PANEL,
            fg=FG,
            font=("Avenir Next", 11, "bold"),
            anchor="w",
        )
        log_title.pack(side="left")

        log_hint = tk.Label(
            log_header,
            text="Progress and validation messages appear here during the update.",
            bg=PANEL,
            fg=MUTED,
            font=("Avenir Next", 9),
            anchor="e",
        )
        log_hint.pack(side="right")

        log_frame = tk.Frame(log_panel, bg=PANEL)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")

        self.log = tk.Text(
            log_frame,
            bg=LOG_BG,
            fg=FG,
            insertbackground=FG,
            selectbackground="#334155",
            relief="flat",
            padx=10,
            pady=10,
            yscrollcommand=scrollbar.set,
            font=("Menlo", 9),
            wrap="word",
            height=16,
        )
        self.log.pack(fill="both", expand=True)
        scrollbar.config(command=self.log.yview)

        self.log.tag_configure("error", foreground=ERROR)
        self.log.tag_configure("success", foreground=SUCCESS)
        self.log.tag_configure("warning", foreground=WARNING)
        self.log.tag_configure("info", foreground=FG)

    def create_logo_label(self, parent):
        """Load and render the company logo when the PNG resource is available."""
        logo_path = get_resource_path(LOGO_FILE)
        if not os.path.exists(logo_path):
            return None

        try:
            raw_logo = tk.PhotoImage(file=logo_path)
        except tk.TclError:
            return None

        target_height = 54
        shrink = max(1, int(round(raw_logo.height() / target_height)))
        self.logo_image = raw_logo.subsample(shrink, shrink)

        return tk.Label(
            parent,
            image=self.logo_image,
            bg=PANEL,
            bd=0,
            highlightthickness=0,
        )

    def create_file_field(self, parent, label_text, helper_text):
        """Render a labeled file picker row and return its entry widget."""
        frame = tk.Frame(parent, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill="x", pady=8)

        accent_bar = tk.Frame(frame, bg=ACCENT, width=6)
        accent_bar.pack(side="left", fill="y")

        content = tk.Frame(frame, bg=PANEL_ALT)
        content.pack(side="left", fill="both", expand=True)

        label = tk.Label(
            content,
            text=label_text,
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 12, "bold"),
            anchor="w",
        )
        label.pack(fill="x", padx=18, pady=(14, 2))

        helper = tk.Label(
            content,
            text=helper_text,
            bg=PANEL_ALT,
            fg=MUTED,
            font=("Avenir Next", 10),
            anchor="w",
        )
        helper.pack(fill="x", padx=18, pady=(0, 10))

        row = tk.Frame(content, bg=PANEL_ALT)
        row.pack(fill="x", padx=18, pady=(0, 16))

        entry = tk.Entry(
            row,
            bg=FIELD_BG,
            fg="#d9e1e8",
            insertbackground="#d9e1e8",
            relief="flat",
            font=("Menlo", 11),
            disabledbackground=FIELD_BG,
            disabledforeground=MUTED,
        )
        entry.pack(side="left", fill="x", expand=True, ipady=8)

        btn = tk.Button(
            row,
            text="Browse",
            bg=ACCENT,
            fg="black",
            activebackground=ACCENT_DARK,
            activeforeground="black",
            relief="flat",
            font=("Avenir Next", 11, "bold"),
            padx=16,
            pady=8,
            command=lambda: self.browse(entry),
            cursor="hand2",
        )
        btn.pack(side="left", padx=(12, 0))

        self.file_entries.append(entry)
        self.browse_buttons.append(btn)
        return entry

    def create_folder_field(self, parent, label_text, helper_text):
        """Render a labeled folder picker row and return its entry widget."""
        frame = tk.Frame(parent, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill="x", pady=8)

        accent_bar = tk.Frame(frame, bg=ACCENT, width=6)
        accent_bar.pack(side="left", fill="y")

        content = tk.Frame(frame, bg=PANEL_ALT)
        content.pack(side="left", fill="both", expand=True)

        label = tk.Label(
            content,
            text=label_text,
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 12, "bold"),
            anchor="w",
        )
        label.pack(fill="x", padx=18, pady=(14, 2))

        helper = tk.Label(
            content,
            text=helper_text,
            bg=PANEL_ALT,
            fg=MUTED,
            font=("Avenir Next", 10),
            anchor="w",
        )
        helper.pack(fill="x", padx=18, pady=(0, 10))

        row = tk.Frame(content, bg=PANEL_ALT)
        row.pack(fill="x", padx=18, pady=(0, 16))

        entry = tk.Entry(
            row,
            bg=FIELD_BG,
            fg="#d9e1e8",
            insertbackground="#d9e1e8",
            relief="flat",
            font=("Menlo", 11),
            disabledbackground=FIELD_BG,
            disabledforeground=MUTED,
        )
        entry.pack(side="left", fill="x", expand=True, ipady=8)

        btn = tk.Button(
            row,
            text="Browse",
            bg=ACCENT,
            fg="black",
            activebackground=ACCENT_DARK,
            activeforeground="black",
            relief="flat",
            font=("Avenir Next", 11, "bold"),
            padx=16,
            pady=8,
            command=lambda: self.browse_folder(entry),
            cursor="hand2",
        )
        btn.pack(side="left", padx=(12, 0))

        self.file_entries.append(entry)
        self.browse_buttons.append(btn)
        return entry

    def browse(self, entry):
        """Open a file picker and write the selected workbook path into the target entry."""
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def browse_folder(self, entry):
        """Open a folder picker and write the selected directory path into the target entry."""
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def on_outer_frame_configure(self, event):
        """Keep the scroll region in sync with the full window content height."""
        self.outer_canvas.configure(scrollregion=self.outer_canvas.bbox("all"))

    def on_outer_canvas_configure(self, event):
        """Resize the scrollable content to match the visible canvas width."""
        self.outer_canvas.itemconfigure(self.outer_canvas_window, width=event.width)

    def on_mousewheel(self, event):
        """Allow mouse-wheel scrolling for the full app layout."""
        self.outer_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def append_log(self, msg):
        """Append a message to the on-screen log and keep the newest line visible."""
        self.log_lines.append(msg)
        lower_msg = msg.lower()
        tag = "info"

        if "error" in lower_msg:
            tag = "error"
        elif "warning" in lower_msg:
            tag = "warning"
        elif "complete" in lower_msg or "passed" in lower_msg or "updated:" in lower_msg:
            tag = "success"

        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)

    def log_message(self, msg):
        """Queue a log message from worker code so UI updates stay on the main thread."""
        self.log_queue.put(("log", msg))

    def process_log_queue(self):
        """Drain queued worker messages and apply them safely to the Tkinter widgets."""
        try:
            while True:
                event, payload = self.log_queue.get_nowait()

                if event == "log":
                    self.append_log(payload)
                elif event == "progress":
                    self.on_progress_update(*payload)
                elif event == "conflicts":
                    self.on_pipeline_conflicts(payload)
                elif event == "success":
                    self.on_pipeline_success(payload)
                elif event == "error":
                    self.on_pipeline_error(payload)
        except queue.Empty:
            pass

        self.root.after(100, self.process_log_queue)

    def set_status(self, message, color=MUTED):
        """Update the single-line run status banner."""
        self.status_var.set(message)
        self.status_label.configure(fg=color)

    def set_progress(self, percent, stage):
        """Update the progress bar and stage label."""
        clamped = max(0, min(100, percent))
        self.progress_var.set(clamped)
        self.progress_text_var.set(f"{int(clamped)}%")
        self.progress_stage_label.configure(text=stage)

    def on_progress_update(self, percent, stage):
        """Apply a progress event emitted by the background worker."""
        self.set_progress(percent, stage)

    def refresh_exclusion_status(self):
        """Show whether the Device ID exclusion file was found and how many entries were loaded."""
        status = get_device_id_exclusion_status()
        if status["found"]:
            count = status["count"]
            noun = "entry" if count == 1 else "entries"
            self.exclusion_label.configure(
                text=f"Device ID exclusion file loaded: {count} {noun} from {status['path']}",
                fg=SUCCESS,
            )
        else:
            self.exclusion_label.configure(
                text=f"Device ID exclusion file not found: {status['path']}",
                fg=WARNING,
            )

    def set_running_state(self, is_running):
        """Enable or disable interactive controls while the pipeline is running."""
        run_state = "disabled" if is_running else "normal"
        entry_state = "disabled" if is_running else "normal"

        self.run_btn.configure(state=run_state)
        self.lock_check.configure(state=run_state)

        for entry in self.file_entries:
            entry.configure(state=entry_state)

        for btn in self.browse_buttons:
            btn.configure(state=run_state)

    def start_pipeline(self):
        """Validate inputs and launch the engine without freezing the UI."""
        if self.worker_thread and self.worker_thread.is_alive():
            return

        source = self.source_entry.get().strip()
        dest = self.dest_entry.get().strip()
        archive_dir = self.archive_entry.get().strip()

        if not source or not dest or not archive_dir:
            self.set_status("Source, destination, and archive folder are all required.", ERROR)
            messagebox.showerror("Error", "Please select the source file, destination file, and archive folder.")
            return

        if self.lock_var.get() and APPROVED_ROOT not in dest:
            self.set_status("Destination failed the approved path safety check.", WARNING)
            messagebox.showwarning(
                "Warning",
                "Destination not inside approved TechTeam path.",
            )
            return

        self.launch_pipeline(source, dest, archive_dir=archive_dir)

    def launch_pipeline(self, source, dest, conflict_resolutions=None, reset_log=True, archive_dir=None):
        """Start a background pipeline run, optionally applying chosen conflict resolutions."""
        self.pending_source = source
        self.pending_dest = dest
        self.pending_archive_dir = archive_dir

        if self.conflict_window and self.conflict_window.winfo_exists():
            self.conflict_window.destroy()
            self.conflict_window = None

        if reset_log:
            self.log.delete("1.0", tk.END)
            self.log_lines = []
            self.append_log(f"Starting {APP_NAME} v{VERSION}")
        else:
            self.append_log("Retrying update with selected conflict values.")

        self.refresh_exclusion_status()
        self.set_progress(0, "Preparing to run")
        self.set_status("Running update. The window will stay responsive while processing.", WARNING)
        self.set_running_state(True)
        self.pending_conflicts = []

        self.worker_thread = threading.Thread(
            target=self.run_pipeline_worker,
            args=(source, dest, conflict_resolutions or {}, archive_dir),
            daemon=True,
        )
        self.worker_thread.start()

    def run_pipeline_worker(self, source, dest, conflict_resolutions, archive_dir):
        """Execute preflight checks and the update engine on a background thread."""
        try:
            self.log_queue.put(("progress", (8, "Checking Google Drive availability")))
            check_drive(os.path.dirname(dest))
            self.log_message("Drive check passed.")

            run_pipeline(
                source,
                dest,
                self.log_message,
                self.queue_progress_update,
                conflict_resolutions,
                archive_dir,
            )
            save_config({"destination": dest, "archive_dir": archive_dir})

            self.log_queue.put(("success", "Update completed successfully."))
        except AttributeConflictError as exc:
            self.log_queue.put(("conflicts", exc.conflicts))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def queue_progress_update(self, percent, stage):
        """Queue a progress update from the engine worker."""
        self.log_queue.put(("progress", (percent, stage)))

    def on_pipeline_success(self, message):
        """Restore the idle UI state and report a successful update."""
        self.set_running_state(False)
        self.set_progress(100, "Finished")
        self.pending_conflicts = []
        self.set_status("Update complete. The destination workbook was updated successfully.", SUCCESS)
        self.append_log("GREEN TICK - UPDATE COMPLETE")
        messagebox.showinfo("Success", message)

    def on_pipeline_conflicts(self, conflicts):
        """Restore the idle UI state and let the user resolve detected conflicts."""
        self.set_running_state(False)
        self.pending_conflicts = conflicts
        self.progress_stage_label.configure(text="Waiting for conflict selections")
        self.set_status("Conflicts need your input before the update can continue.", WARNING)
        self.show_conflict_chooser(conflicts)

    def on_pipeline_error(self, message):
        """Restore the idle UI state and surface a pipeline failure to the user."""
        self.set_running_state(False)
        if self.progress_var.get() < 100:
            self.progress_stage_label.configure(text="Stopped due to an error")
        self.set_status("Update failed. Review the log and try again.", ERROR)
        self.append_log("ERROR: " + message)
        messagebox.showerror("Error", message)

    def show_conflict_chooser(self, conflicts):
        """Open a dedicated window for browsing detected conflicts."""
        if not conflicts:
            return

        selections = {}
        active_conflict_index = {"value": None}

        window = tk.Toplevel(self.root)
        window.title("Conflict Chooser")
        window.geometry("980x620")
        window.configure(bg=BG)
        self.conflict_window = window
        window.transient(self.root)
        window.protocol("WM_DELETE_WINDOW", lambda: self.close_conflict_window(window))

        header = tk.Frame(window, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        header.pack(fill="x", padx=16, pady=16)

        title = tk.Label(
            header,
            text="Conflict Chooser",
            bg=PANEL,
            fg=FG,
            font=("Avenir Next", 16, "bold"),
            anchor="w",
        )
        title.pack(fill="x", padx=14, pady=(12, 4))

        hint = tk.Label(
            header,
            text="Choose one value for each conflict, then apply the selections to continue the update.",
            bg=PANEL,
            fg=MUTED,
            font=("Avenir Next", 10),
            anchor="w",
            justify="left",
        )
        hint.pack(fill="x", padx=14, pady=(0, 12))

        body = tk.Frame(window, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        left = tk.Frame(body, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        left.pack(side="left", fill="y")

        right = tk.Frame(body, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))

        list_label = tk.Label(
            left,
            text="Detected Conflicts",
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 11, "bold"),
            anchor="w",
        )
        list_label.pack(fill="x", padx=12, pady=(10, 6))

        list_frame = tk.Frame(left, bg=PANEL_ALT)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        list_scroll = tk.Scrollbar(list_frame)
        list_scroll.pack(side="right", fill="y")

        conflict_list = tk.Listbox(
            list_frame,
            bg="#11161c",
            fg=FG,
            selectbackground="#334155",
            selectforeground=FG,
            relief="flat",
            font=("Menlo", 9),
            width=34,
            exportselection=False,
            yscrollcommand=list_scroll.set,
        )
        conflict_list.pack(side="left", fill="both", expand=True)
        list_scroll.config(command=conflict_list.yview)

        detail_header = tk.Label(
            right,
            text="Choose Value",
            bg=PANEL,
            fg=FG,
            font=("Avenir Next", 11, "bold"),
            anchor="w",
        )
        detail_header.pack(fill="x", padx=12, pady=(10, 6))

        selected_conflict_var = tk.StringVar(value="Select a conflict to choose which value should win.")
        selected_value_var = tk.StringVar(value="No value selected yet.")

        conflict_title_label = tk.Label(
            right,
            textvariable=selected_conflict_var,
            bg=PANEL,
            fg=MUTED,
            font=("Avenir Next", 10),
            anchor="w",
            justify="left",
            wraplength=600,
        )
        conflict_title_label.pack(fill="x", padx=12, pady=(0, 10))

        option_frame = tk.Frame(right, bg=PANEL)
        option_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        option_scroll = tk.Scrollbar(option_frame)
        option_scroll.pack(side="right", fill="y")

        option_list = tk.Listbox(
            option_frame,
            bg="#12151a",
            fg=FG,
            selectbackground=ACCENT,
            selectforeground="black",
            relief="flat",
            font=("Menlo", 10),
            exportselection=False,
            yscrollcommand=option_scroll.set,
        )
        option_list.pack(fill="both", expand=True)
        option_scroll.config(command=option_list.yview)

        selection_bar = tk.Frame(right, bg=PANEL_ALT, highlightbackground=BORDER, highlightthickness=1)
        selection_bar.pack(fill="x", padx=12, pady=(0, 12))

        selection_title = tk.Label(
            selection_bar,
            text="Current Selection",
            bg=PANEL_ALT,
            fg=FG,
            font=("Avenir Next", 10, "bold"),
            anchor="w",
        )
        selection_title.pack(fill="x", padx=12, pady=(8, 2))

        selection_value = tk.Label(
            selection_bar,
            textvariable=selected_value_var,
            bg=PANEL_ALT,
            fg=SUCCESS,
            font=("Avenir Next", 10),
            anchor="w",
            justify="left",
            wraplength=600,
        )
        selection_value.pack(fill="x", padx=12, pady=(0, 8))

        footer = tk.Frame(window, bg=BG)
        footer.pack(fill="x", padx=16, pady=(0, 16))

        summary_var = tk.StringVar(value=f"Selected 0 of {len(conflicts)} conflicts.")

        summary_label = tk.Label(
            footer,
            textvariable=summary_var,
            bg=BG,
            fg=MUTED,
            font=("Avenir Next", 10, "bold"),
            anchor="w",
        )
        summary_label.pack(side="left")

        cancel_btn = tk.Button(
            footer,
            text="Close",
            bg=SECONDARY,
            fg="black",
            activebackground=SECONDARY_HOVER,
            activeforeground="black",
            relief="flat",
            font=("Avenir Next", 10, "bold"),
            padx=14,
            pady=7,
            command=window.destroy,
            cursor="hand2",
        )
        cancel_btn.pack(side="right")

        apply_btn = tk.Button(
            footer,
            text="Apply Selected Values",
            bg=ACCENT,
            fg="black",
            activebackground=ACCENT_DARK,
            activeforeground="black",
            relief="flat",
            font=("Avenir Next", 10, "bold"),
            padx=14,
            pady=7,
            cursor="hand2",
        )
        apply_btn.pack(side="right", padx=(0, 10))

        def update_summary():
            summary_var.set(f"Selected {len(selections)} of {len(conflicts)} conflicts.")

        def conflict_display_label(conflict):
            key = conflict_key(conflict)
            prefix = "[Chosen] " if key in selections else "[Open] "
            return prefix + conflict["title"]

        def conflict_key(conflict):
            return (conflict["model_norm"], conflict["field_name"])

        def render_conflict(index):
            active_conflict_index["value"] = index
            conflict = conflicts[index]
            selected_conflict_var.set(conflict["title"])
            option_list.delete(0, tk.END)

            if conflict["options"]:
                for option in conflict["options"]:
                    option_list.insert(tk.END, option["detail"])
            else:
                option_list.insert(tk.END, "No detail lines were captured for this conflict.")

            chosen_value = selections.get(conflict_key(conflict))
            if chosen_value is not None:
                for option_index, option in enumerate(conflict["options"]):
                    if option["value"] == chosen_value:
                        option_list.selection_set(option_index)
                        option_list.activate(option_index)
                        option_list.see(option_index)
                        break
                selected_value_var.set(f"Selected value: {chosen_value}")
            else:
                selected_value_var.set("No value selected yet.")

        def choose_current_option(_event=None):
            option_selection = option_list.curselection()

            if active_conflict_index["value"] is None or not option_selection:
                return

            conflict = conflicts[active_conflict_index["value"]]
            chosen_option = conflict["options"][option_selection[0]]
            chosen_value = chosen_option["value"]
            selections[conflict_key(conflict)] = chosen_value
            selected_value_var.set(f"Selected value: {chosen_value}")
            conflict_list.delete(active_conflict_index["value"])
            conflict_list.insert(active_conflict_index["value"], conflict_display_label(conflict))
            conflict_list.selection_set(active_conflict_index["value"])
            conflict_list.activate(active_conflict_index["value"])
            update_summary()

        def apply_selections():
            missing = [conflict["title"] for conflict in conflicts if conflict_key(conflict) not in selections]
            if missing:
                messagebox.showwarning(
                    "Selections Required",
                    "Choose a value for every conflict before continuing.",
                    parent=window,
                )
                return

            window.destroy()
            self.conflict_window = None
            self.launch_pipeline(
                self.pending_source,
                self.pending_dest,
                conflict_resolutions=selections,
                reset_log=False,
                archive_dir=self.pending_archive_dir,
            )

        for conflict in conflicts:
            conflict_list.insert(tk.END, conflict_display_label(conflict))

        def on_select(_event):
            selection = conflict_list.curselection()
            if selection:
                render_conflict(selection[0])

        conflict_list.bind("<<ListboxSelect>>", on_select)
        option_list.bind("<<ListboxSelect>>", choose_current_option)
        option_list.bind("<ButtonRelease-1>", choose_current_option)
        option_list.bind("<Double-Button-1>", choose_current_option)
        apply_btn.configure(command=apply_selections)

        if conflicts:
            conflict_list.selection_set(0)
            render_conflict(0)
        update_summary()

    def close_conflict_window(self, window):
        """Clear the active conflict window reference when the chooser is closed."""
        if window.winfo_exists():
            window.destroy()
        self.conflict_window = None


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
