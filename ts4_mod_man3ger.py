import json
import os
import shutil
import subprocess
import sys
import zipfile
import webbrowser
import hashlib
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import rarfile
except Exception:
    rarfile = None

# ============================================================
# TS4 Mod Man3ger — Beta v0.1.1
# ============================================================
APP_TITLE = "TS4 Mod Man3ger"
APP_VERSION = "Beta v0.1.1"
APP_TITLEBAR = f"{APP_TITLE} — {APP_VERSION}"

# ============================================================
# SOCIAL / BRANDING LINKS
# ============================================================
SOCIALS = {
    "Discord Forever": "https://discord.gg/QHYwBZ4QqJ",
    "TikTok": "https://www.tiktok.com/@blazedkiwi0",
    "YouTube": "https://www.youtube.com/@Blazedkiwi0",
    "X": "https://x.com/DGotchu21632",
    "GitHub": "https://github.com/blazedgaming",
}

# ============================================================
# BUTTON LABEL SETTINGS
# ============================================================
BTN_SCAN = "Scan"
BTN_PREVIEW = "Preview"
BTN_PROCESS = "Process"
BTN_FINALIZE = "Finalize"
BTN_AUDIT = "Create Audit"
BTN_LOCK = "Lock"
BTN_UNLOCK = "Unlock"
BTN_CLEAN = "Clean Mode"

# ============================================================
# FILE SETTINGS
# ============================================================
APP_FOLDER_NAME = "TS4_Mod_Man3ger"
CONFIG_FILENAME = "ts4_mod_man3ger_config.json"

VALID_MOD_EXTENSIONS = {".package", ".ts4script"}
VALID_ARCHIVE_EXTENSIONS = {".zip", ".rar", ".7z"}
RAR_7Z_EXTENSIONS = {".rar", ".7z"}

SAFE_CACHE_FILES = {
    "localthumbcache.package",
}
SAFE_CACHE_FOLDERS = {
    "cache",
    "cachestr",
    "onlinethumbnailcache",
}

DEFAULT_CONFIG = {
    "mods_folder": "",
    "downloads_folder": "",
    "archive_folder": "",
    "backup_folder": "",
    "reports_folder": "",
    "live_mods_folder": "",
    "test_mods_folder": "",
    "current_preset": "Live",
    "is_locked": False,
    "always_prompt_backup_before_import": True,
    "last_audit_file": "",
    "last_backup_folder": "",
    "removed_duplicates_folder": "",
    "last_removed_duplicates_folder": "",
}


class TS4ModMan3ger(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLEBAR)
        self.geometry("1460x900")
        self.minsize(1240, 780)

        self.app_root = self.get_app_root()
        self.config_path = self.app_root / CONFIG_FILENAME
        self.ensure_app_dirs()

        self.config_data = self.load_config()
        self.scan_items = []

        self.mods_var = tk.StringVar()
        self.downloads_var = tk.StringVar()
        self.archive_var = tk.StringVar()
        self.removed_dupes_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="LIVE MODE")
        self.locked_var = tk.StringVar(value="UNLOCKED")
        self.current_preset_var = tk.StringVar(value="Live")
        self.status_text = tk.StringVar(value="Ready")
        self.last_backup_var = tk.StringVar(value="No backup yet")
        self.last_audit_var = tk.StringVar(value="No audit yet")
        self.summary_var = tk.StringVar(value="No session activity yet.")
        self.url_var = tk.StringVar()
        self.url_label_var = tk.StringVar()

        self.build_base_style()

        if (
            not self.config_data.get("mods_folder")
            or not self.config_data.get("downloads_folder")
            or not self.config_data.get("archive_folder")
        ):
            self.run_setup_wizard()

        self.apply_config_to_vars()
        self.build_main_ui()
        self.refresh_status()

    # ============================================================
    # App folders / config
    # ============================================================
    def get_app_root(self) -> Path:
        root = Path.home() / APP_FOLDER_NAME
        root.mkdir(parents=True, exist_ok=True)
        return root

    def ensure_app_dirs(self):
        for name in ["backups", "reports", "temp_extract"]:
            (self.app_root / name).mkdir(parents=True, exist_ok=True)

    def get_backup_folder(self) -> Path:
        raw = self.config_data.get("backup_folder", "")
        path = Path(raw) if raw else self.app_root / "backups"
        self.config_data["backup_folder"] = str(path)
        path.mkdir(parents=True, exist_ok=True)
        return path

    def get_reports_folder(self) -> Path:
        raw = self.config_data.get("reports_folder", "")
        path = Path(raw) if raw else self.app_root / "reports"
        self.config_data["reports_folder"] = str(path)
        path.mkdir(parents=True, exist_ok=True)
        return path

    def get_temp_extract_folder(self) -> Path:
        path = self.app_root / "temp_extract"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def load_config(self):
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                merged = DEFAULT_CONFIG.copy()
                merged.update(data)
                return merged
            except Exception:
                return DEFAULT_CONFIG.copy()
        return DEFAULT_CONFIG.copy()

    def save_config(self):
        self.config_path.write_text(json.dumps(self.config_data, indent=2), encoding="utf-8")

    def apply_config_to_vars(self):
        self.mods_var.set(self.config_data.get("mods_folder", ""))
        self.downloads_var.set(self.config_data.get("downloads_folder", ""))
        self.archive_var.set(self.config_data.get("archive_folder", ""))
        self.removed_dupes_var.set(self.config_data.get("removed_duplicates_folder", ""))
        self.current_preset_var.set(self.config_data.get("current_preset", "Live"))

        live = self.config_data.get("live_mods_folder", "")
        current = self.config_data.get("mods_folder", "")
        self.mode_var.set("TEST MODE" if live and current and Path(current) != Path(live) else "LIVE MODE")
        self.locked_var.set("LOCKED" if self.config_data.get("is_locked") else "UNLOCKED")
        self.last_backup_var.set(self.config_data.get("last_backup_folder") or "No backup yet")
        self.last_audit_var.set(self.config_data.get("last_audit_file") or "No audit yet")

    def update_config_from_vars(self):
        self.config_data["mods_folder"] = self.mods_var.get().strip()
        self.config_data["downloads_folder"] = self.downloads_var.get().strip()
        self.config_data["archive_folder"] = self.archive_var.get().strip()
        self.config_data["removed_duplicates_folder"] = self.removed_dupes_var.get().strip()
        self.save_config()

    # ============================================================
    # Style
    # ============================================================
    def build_base_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        base_font = ("Georgia", 12)
        bold_font = ("Georgia", 12, "bold")
        header_font = ("Georgia", 20, "bold")
        subheader_font = ("Georgia", 13, "bold")

        self.option_add("*Font", base_font)

        style.configure("TLabel", font=base_font)
        style.configure("TButton", font=bold_font, padding=(12, 8))
        style.configure("TEntry", font=base_font)
        style.configure("TMenubutton", font=bold_font)
        style.configure("TLabelframe", padding=12)
        style.configure("TLabelframe.Label", font=subheader_font)

        style.configure("Header.TLabel", font=header_font)
        style.configure("SubHeader.TLabel", font=subheader_font)
        style.configure("Card.TLabelframe", padding=12)
        style.configure("Card.TLabelframe.Label", font=subheader_font)
        style.configure("Quick.TButton", font=("Georgia", 12, "bold"), padding=(16, 10))

    # ============================================================
    # Wizard
    # ============================================================
    def run_setup_wizard(self):
        wizard = tk.Toplevel(self)
        wizard.title(f"{APP_TITLE} Setup Wizard")
        wizard.geometry("780x360")
        wizard.resizable(False, False)
        wizard.transient(self)
        wizard.grab_set()

        mods_var = tk.StringVar(value=self.config_data.get("mods_folder", ""))
        downloads_var = tk.StringVar(value=self.config_data.get("downloads_folder", ""))
        archive_var = tk.StringVar(value=self.config_data.get("archive_folder", ""))

        outer = ttk.Frame(wizard, padding=20)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=f"{APP_TITLE} — First Setup", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        ttk.Label(outer, text="Choose the three main folders this app will use.").pack(anchor="w", pady=(0, 18))

        self.add_setup_row(outer, "Where is your Mods folder?", mods_var)
        self.add_setup_row(outer, "Where is your Downloads folder?", downloads_var)
        self.add_setup_row(outer, "Where should archived mods go?", archive_var)

        ttk.Label(outer, text="Missing folders can be created automatically.").pack(anchor="w", pady=(14, 0))

        button_row = ttk.Frame(outer)
        button_row.pack(fill="x", pady=(22, 0))

        def finish():
            mods = mods_var.get().strip()
            downloads = downloads_var.get().strip()
            archive = archive_var.get().strip()
            if not mods or not downloads or not archive:
                messagebox.showerror("Missing Paths", "All three paths are required.", parent=wizard)
                return
            for raw_path in [mods, downloads, archive]:
                p = Path(raw_path)
                if not p.exists():
                    if messagebox.askyesno("Create Folder", f"This folder does not exist:\n\n{raw_path}\n\nCreate it now?", parent=wizard):
                        p.mkdir(parents=True, exist_ok=True)
                    else:
                        return
            self.config_data["mods_folder"] = mods
            self.config_data["downloads_folder"] = downloads
            self.config_data["archive_folder"] = archive
            archive_parent = Path(archive).parent if Path(archive).parent != Path("") else Path(archive)
            removed_folder = archive_parent / "@Removed Duplicates"
            removed_folder.mkdir(parents=True, exist_ok=True)
            self.config_data["removed_duplicates_folder"] = str(removed_folder)
            self.config_data["live_mods_folder"] = self.config_data.get("live_mods_folder") or mods
            self.config_data["backup_folder"] = str(self.app_root / "backups")
            self.config_data["reports_folder"] = str(self.app_root / "reports")
            self.config_data["current_preset"] = "Live"
            self.save_config()
            wizard.destroy()

        ttk.Button(button_row, text="Save Setup", command=finish).pack(side="right")
        ttk.Button(button_row, text="Cancel", command=wizard.destroy).pack(side="right", padx=(0, 10))
        self.wait_window(wizard)

    def add_setup_row(self, parent, label_text, variable):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=9)
        ttk.Label(row, text=label_text, width=30).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=10)
        ttk.Button(row, text="Browse", command=lambda v=variable: self.browse_folder_to_var(v)).pack(side="left")

    def browse_folder_to_var(self, variable):
        chosen = filedialog.askdirectory(title="Select Folder")
        if chosen:
            variable.set(chosen)

    # ============================================================
    # UI
    # ============================================================
    def build_main_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        self.build_menu_bar()
        self.build_header()
        self.build_quick_actions()
        self.build_body()
        self.build_status_bar()

    def build_menu_bar(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Run Setup Wizard Again", command=self.run_setup_wizard_again)
        file_menu.add_separator()
        file_menu.add_command(label="Set Mods Folder", command=lambda: self.pick_path("mods"))
        file_menu.add_command(label="Set Downloads Folder", command=lambda: self.pick_path("downloads"))
        file_menu.add_command(label="Set Archive Folder", command=lambda: self.pick_path("archive"))
        file_menu.add_command(label="Set Removed Duplicates Folder", command=lambda: self.pick_path("removed"))
        file_menu.add_separator()
        file_menu.add_command(label="Switch to Live Preset", command=self.switch_to_live_preset)
        file_menu.add_command(label="Switch to Test Preset", command=self.switch_to_test_preset)
        file_menu.add_separator()
        file_menu.add_command(label="Open Mods Folder", command=lambda: self.open_folder(self.mods_var.get()))
        file_menu.add_command(label="Open Downloads Folder", command=lambda: self.open_folder(self.downloads_var.get()))
        file_menu.add_command(label="Open Archive Folder", command=lambda: self.open_folder(self.archive_var.get()))
        file_menu.add_command(label="Open Removed Duplicates Folder", command=lambda: self.open_folder(self.removed_dupes_var.get()))
        file_menu.add_command(label="Open Reports Folder", command=lambda: self.open_folder(str(self.get_reports_folder())))
        file_menu.add_command(label="Open Backup Folder", command=lambda: self.open_folder(str(self.get_backup_folder())))
        file_menu.add_separator()
        file_menu.add_command(label="Save Settings", command=self.save_settings_action)
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        lock_menu = tk.Menu(menubar, tearoff=0)
        lock_menu.add_command(label="Lock Current Mods Folder", command=self.lock_current_folder)
        lock_menu.add_command(label="Unlock Current Mods Folder", command=self.unlock_current_folder)
        lock_menu.add_command(label="View Lock Status", command=self.show_lock_status)
        menubar.add_cascade(label="Lock", menu=lock_menu)

        backup_menu = tk.Menu(menubar, tearoff=0)
        backup_menu.add_command(label="Create Full Backup", command=self.create_backup)
        backup_menu.add_command(label="Quick Backup", command=self.create_backup)
        backup_menu.add_command(label="Restore Backup", command=self.placeholder_action)
        backup_menu.add_command(label="Open Backup Folder", command=lambda: self.open_folder(str(self.get_backup_folder())))
        menubar.add_cascade(label="Backup", menu=backup_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Scan Downloads", command=self.scan_downloads)
        tools_menu.add_command(label="Import Preview", command=self.preview_import)
        tools_menu.add_command(label="Process Downloads", command=self.process_downloads)
        tools_menu.add_command(label="Finalize", command=self.finalize_cleanup)
        tools_menu.add_command(label="Create Mod Audit", command=self.create_mod_audit)
        tools_menu.add_command(label="Open Last Audit", command=self.open_last_audit)
        tools_menu.add_command(label="Export Log", command=self.export_log)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        updates_menu = tk.Menu(menubar, tearoff=0)
        updates_menu.add_command(label="Check Mod Update", command=self.placeholder_action)
        updates_menu.add_command(label="Analyze Link", command=self.analyze_link)
        updates_menu.add_command(label="Download File", command=self.download_file)
        updates_menu.add_command(label="Open Mod Page", command=self.open_link_in_browser)
        menubar.add_cascade(label="Updates", menu=updates_menu)

        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label=f"About {APP_TITLE}", command=self.show_about)
        about_menu.add_command(label="Open HTML Help Guide", command=self.open_help_file)
        about_menu.add_command(label="Version Info", command=lambda: messagebox.showinfo("Version", APP_TITLEBAR))
        menubar.add_cascade(label="About", menu=about_menu)

        self.config(menu=menubar)

    def build_header(self):
        header = ttk.Frame(self, padding=(16, 14))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=1)

        title_frame = ttk.Frame(header)
        title_frame.grid(row=0, column=0, sticky="w")
        ttk.Label(title_frame, text=APP_TITLE, style="Header.TLabel").pack(anchor="w")
        ttk.Label(title_frame, text=APP_VERSION).pack(anchor="w", pady=(2, 0))

        status_frame = ttk.Frame(header)
        status_frame.grid(row=0, column=1, sticky="e")
        ttk.Label(status_frame, textvariable=self.mode_var, style="SubHeader.TLabel").pack(anchor="e", pady=(0, 4))
        ttk.Label(status_frame, textvariable=self.locked_var, style="SubHeader.TLabel").pack(anchor="e")

        banner = ttk.LabelFrame(self, text="Status Banner", style="Card.TLabelframe")
        banner.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))
        banner.columnconfigure(1, weight=1)
        banner.columnconfigure(3, weight=1)
        banner.columnconfigure(5, weight=1)

        ttk.Label(banner, text="Mods Folder", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Label(banner, textvariable=self.mods_var).grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=3)
        ttk.Label(banner, text="Downloads", style="SubHeader.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=3)
        ttk.Label(banner, textvariable=self.downloads_var).grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=3)
        ttk.Label(banner, text="Archive", style="SubHeader.TLabel").grid(row=0, column=4, sticky="w", padx=(0, 8), pady=3)
        ttk.Label(banner, textvariable=self.archive_var).grid(row=0, column=5, sticky="ew", pady=3)

    def build_quick_actions(self):
        quick = ttk.Frame(self, padding=(16, 0, 16, 12))
        quick.grid(row=2, column=0, sticky="ew")
        ttk.Button(quick, text=BTN_SCAN, style="Quick.TButton", command=self.scan_downloads).pack(side="left", padx=(0, 8))
        ttk.Button(quick, text=BTN_PREVIEW, style="Quick.TButton", command=self.preview_import).pack(side="left", padx=(0, 8))
        ttk.Button(quick, text=BTN_PROCESS, style="Quick.TButton", command=self.process_downloads).pack(side="left", padx=(0, 8))
        ttk.Button(quick, text=BTN_FINALIZE, style="Quick.TButton", command=self.finalize_cleanup).pack(side="left", padx=(0, 8))
        ttk.Button(quick, text=BTN_AUDIT, style="Quick.TButton", command=self.create_mod_audit).pack(side="left", padx=(0, 8))
        ttk.Button(quick, text=BTN_CLEAN, style="Quick.TButton", command=self.clean_duplicate_content).pack(side="left", padx=(0, 8))
        self.lock_toggle_btn = ttk.Button(quick, text=BTN_UNLOCK if self.config_data.get("is_locked") else BTN_LOCK, style="Quick.TButton", command=self.toggle_lock)
        self.lock_toggle_btn.pack(side="left")

    def build_body(self):
        body = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        body.grid(row=3, column=0, sticky="nsew", padx=16, pady=(0, 12))
        left = ttk.Frame(body, padding=2)
        right = ttk.Frame(body, padding=2)
        body.add(left, weight=4)
        body.add(right, weight=2)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        self.build_path_panel(left)
        self.build_link_panel(left)
        self.build_preview_panel(left)
        self.build_session_panel(right)
        self.build_notes_panel(right)
        self.build_log_panel(right)

    def build_path_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Path Summary", style="Card.TLabelframe")
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)
        self.add_path_row(frame, 0, "Mods Folder", self.mods_var, lambda: self.pick_path("mods"), lambda: self.open_folder(self.mods_var.get()))
        self.add_path_row(frame, 1, "Downloads Folder", self.downloads_var, lambda: self.pick_path("downloads"), lambda: self.open_folder(self.downloads_var.get()))
        self.add_path_row(frame, 2, "Archive Folder", self.archive_var, lambda: self.pick_path("archive"), lambda: self.open_folder(self.archive_var.get()))
        self.add_path_row(frame, 3, "Removed Duplicates", self.removed_dupes_var, lambda: self.pick_path("removed"), lambda: self.open_folder(self.removed_dupes_var.get()))
        ttk.Label(frame, text="Current Preset", style="SubHeader.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Label(frame, textvariable=self.current_preset_var).grid(row=4, column=1, sticky="w", pady=(12, 0))

    def add_path_row(self, parent, row_idx, label_text, variable, browse_cmd, open_cmd):
        ttk.Label(parent, text=label_text, style="SubHeader.TLabel").grid(row=row_idx, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row_idx, column=1, sticky="ew", padx=10, pady=6)
        ttk.Button(parent, text="Browse", command=browse_cmd).grid(row=row_idx, column=2, padx=(0, 6), pady=6)
        ttk.Button(parent, text="Open", command=open_cmd).grid(row=row_idx, column=3, pady=6)

    def build_link_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Download Intake", style="Card.TLabelframe")
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Link / URL", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.url_var).grid(row=0, column=1, sticky="ew", padx=10, pady=6)
        ttk.Label(frame, text="Optional Label", style="SubHeader.TLabel").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.url_label_var).grid(row=1, column=1, sticky="ew", padx=10, pady=6)
        btns = ttk.Frame(frame)
        btns.grid(row=0, column=2, rowspan=2, sticky="ns")
        ttk.Button(btns, text="Analyze Link", command=self.analyze_link).pack(fill="x", pady=2)
        ttk.Button(btns, text="Download File", command=self.download_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Open in Browser", command=self.open_link_in_browser).pack(fill="x", pady=2)

    def build_preview_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Scan Results / Import Preview", style="Card.TLabelframe")
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        columns = ("name", "type", "source", "destination", "action")
        self.preview_tree = ttk.Treeview(frame, columns=columns, show="headings", height=16)
        for col, text, width in [
            ("name", "File Name", 240),
            ("type", "Type", 120),
            ("source", "Source Path", 320),
            ("destination", "Destination", 250),
            ("action", "Action", 160),
        ]:
            self.preview_tree.heading(col, text=text)
            self.preview_tree.column(col, width=width, anchor="w")
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=scroll_y.set)
        scroll_y.grid(row=0, column=1, sticky="ns")

    def build_session_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Session Summary", style="Card.TLabelframe")
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Summary", style="SubHeader.TLabel").grid(row=0, column=0, sticky="nw")
        ttk.Label(frame, textvariable=self.summary_var, wraplength=420, justify="left").grid(row=0, column=1, sticky="w")
        ttk.Label(frame, text="Last Backup", style="SubHeader.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Label(frame, textvariable=self.last_backup_var, wraplength=420, justify="left").grid(row=1, column=1, sticky="w", pady=(12, 0))
        ttk.Label(frame, text="Last Audit", style="SubHeader.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(frame, textvariable=self.last_audit_var, wraplength=420, justify="left").grid(row=2, column=1, sticky="w", pady=(8, 0))

    def build_notes_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Notes", style="Card.TLabelframe")
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(frame, text=(
            "Beta v0.1.1 additions:\n\n"
            "- duplicate detection by filename and content hash\n"
            "- .rar and .7z archive support if extractor exists\n"
            "- archive support status shown in scan/preview\n"
            "- full placeholder methods included to avoid startup crashes"
        ), justify="left", wraplength=420).pack(anchor="w")

    def build_log_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Activity Log", style="Card.TLabelframe")
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(frame, wrap="word", height=18, font=("Georgia", 12))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll_y.set)
        scroll_y.grid(row=0, column=1, sticky="ns")

    def build_status_bar(self):
        ttk.Label(self, textvariable=self.status_text, relief="sunken", anchor="w").grid(row=4, column=0, sticky="ew")

    # ============================================================
    # Logging / helpers
    # ============================================================
    def log(self, message):
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.status_text.set(message)

    def clear_preview(self):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

    def add_preview_row(self, name, file_type, source, destination, action):
        self.preview_tree.insert("", "end", values=(name, file_type, source, destination, action))

    def hash_file(self, file_path: Path):
        try:
            h = hashlib.md5()
            with file_path.open("rb") as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
            return h.hexdigest()
        except Exception:
            return None

    def get_archive_support_text(self, ext: str):
        ext = ext.lower()
        if ext == ".zip":
            return "archive"
        if ext in RAR_7Z_EXTENSIONS:
            return "archive" if self.has_7z_support() or self.has_rarfile_support() else "archive-missing-tool"
        return "archive"

    def find_7z_executable(self):
        candidates = [
            shutil.which("7z"),
            shutil.which("7za"),
            shutil.which("7zr"),
            r"C:\\Program Files\\7-Zip\\7z.exe",
            r"C:\\Program Files (x86)\\7-Zip\\7z.exe",
        ]
        for candidate in candidates:
            if candidate and Path(candidate).exists():
                return str(candidate)
        return None

    def has_7z_support(self):
        return self.find_7z_executable() is not None

    def has_rarfile_support(self):
        if rarfile is None:
            return False
        try:
            if self.find_7z_executable():
                rarfile.UNRAR_TOOL = self.find_7z_executable()
            return True
        except Exception:
            return False

    def extract_archive(self, archive_path: Path, target_dir: Path):
        ext = archive_path.suffix.lower()
        if ext == ".zip":
            with zipfile.ZipFile(archive_path, "r") as zf:
                zf.extractall(target_dir)
            return True, "zip extracted"

        exe = self.find_7z_executable()
        if exe:
            cmd = [exe, "x", str(archive_path), f"-o{str(target_dir)}", "-y"]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                return True, f"{ext} extracted with 7-Zip"
            return False, result.stderr.strip() or result.stdout.strip() or "7-Zip extraction failed"

        if ext == ".rar" and rarfile is not None:
            try:
                rf = rarfile.RarFile(str(archive_path))
                rf.extractall(str(target_dir))
                return True, "rar extracted with rarfile"
            except Exception as exc:
                return False, str(exc)

        return False, f"No extractor available for {ext}. Install 7-Zip and add it to PATH."

    # ============================================================
    # Core helper methods
    # ============================================================
    def require_paths(self):
        mods_folder = self.mods_var.get().strip()
        downloads_folder = self.downloads_var.get().strip()
        archive_folder = self.archive_var.get().strip()
        if not mods_folder:
            messagebox.showerror("Missing Mods Folder", "Set your Mods folder first.")
            return None
        if not downloads_folder:
            messagebox.showerror("Missing Downloads Folder", "Set your Downloads folder first.")
            return None
        if not archive_folder:
            messagebox.showerror("Missing Archive Folder", "Set your Archive folder first.")
            return None
        mods_path = Path(mods_folder)
        downloads_path = Path(downloads_folder)
        archive_path = Path(archive_folder)
        if not downloads_path.exists():
            messagebox.showerror("Downloads Folder Missing", "The selected Downloads folder does not exist.")
            return None
        mods_path.mkdir(parents=True, exist_ok=True)
        archive_path.mkdir(parents=True, exist_ok=True)
        return mods_path, downloads_path, archive_path

    def get_removed_duplicates_folder(self) -> Path:
        removed_folder = self.removed_dupes_var.get().strip() or self.config_data.get("removed_duplicates_folder", "").strip()
        if not removed_folder:
            archive_folder = self.archive_var.get().strip() or self.config_data.get("archive_folder", "").strip()
            if archive_folder:
                removed_folder = str(Path(archive_folder).parent / "@Removed Duplicates")
            else:
                removed_folder = str(self.app_root / "removed_duplicates")
        removed_path = Path(removed_folder)
        removed_path.mkdir(parents=True, exist_ok=True)
        self.removed_dupes_var.set(str(removed_path))
        self.config_data["removed_duplicates_folder"] = str(removed_path)
        self.save_config()
        return removed_path

    def find_duplicate_content_groups(self, mods_path: Path):
        duplicate_hash_map = {}
        for file_path in mods_path.rglob("*"):
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() not in VALID_MOD_EXTENSIONS:
                continue
            file_hash = self.hash_file(file_path)
            if file_hash:
                duplicate_hash_map.setdefault(file_hash, []).append(file_path)
        return {file_hash: paths for file_hash, paths in duplicate_hash_map.items() if len(paths) > 1}

    def choose_keep_file(self, paths):
        def score(p: Path):
            lower_parts = [part.lower() for part in p.parts]
            penalty = 0
            joined = " ".join(lower_parts)
            if "@removed duplicates" in joined:
                penalty += 100000
            if "optional packages" in joined:
                penalty += 5000
            if "@archive" in joined or "archive" in joined:
                penalty += 3000
            try:
                mtime = p.stat().st_mtime
            except Exception:
                mtime = 0
            return (penalty, -mtime, len(str(p)))
        return sorted(paths, key=score)[0]

    def show_removed_duplicates_window(self, removed_items, removed_folder: Path):
        win = tk.Toplevel(self)
        win.title("Removed Duplicate Mods")
        win.geometry("980x620")
        win.transient(self)

        outer = ttk.Frame(win, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        ttk.Label(
            outer,
            text=f"Moved {len(removed_items)} duplicate files to:\n{removed_folder}",
            style="SubHeader.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        cols = ("removed_name", "kept_name", "removed_from", "kept_at")
        tree = ttk.Treeview(outer, columns=cols, show="headings")
        tree.heading("removed_name", text="Removed File")
        tree.heading("kept_name", text="Kept File")
        tree.heading("removed_from", text="Original Location")
        tree.heading("kept_at", text="Kept At")
        tree.column("removed_name", width=220, anchor="w")
        tree.column("kept_name", width=220, anchor="w")
        tree.column("removed_from", width=250, anchor="w")
        tree.column("kept_at", width=250, anchor="w")

        ybar = ttk.Scrollbar(outer, orient="vertical", command=tree.yview)
        xbar = ttk.Scrollbar(outer, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        tree.grid(row=1, column=0, sticky="nsew")
        ybar.grid(row=1, column=1, sticky="ns")
        xbar.grid(row=2, column=0, sticky="ew")

        for item in removed_items:
            tree.insert(
                "",
                "end",
                values=(
                    item["removed_name"],
                    item["kept_name"],
                    item["removed_from"],
                    item["kept_at"],
                ),
            )

        btns = ttk.Frame(outer)
        btns.grid(row=3, column=0, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Open Removed Folder", command=lambda: self.open_folder(str(removed_folder))).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="left")

    def clean_duplicate_content(self):
        if self.config_data.get("is_locked"):
            messagebox.showwarning("Folder Locked", "Unlock the current Mods folder before running Clean Mode.")
            return

        mods_folder = self.mods_var.get().strip()
        if not mods_folder:
            messagebox.showerror("Missing Mods Folder", "Set your Mods folder first.")
            return

        mods_path = Path(mods_folder)
        if not mods_path.exists():
            messagebox.showerror("Missing Mods Folder", "The selected Mods folder does not exist.")
            return

        duplicate_groups = self.find_duplicate_content_groups(mods_path)
        duplicate_file_total = sum(len(paths) - 1 for paths in duplicate_groups.values())
        if duplicate_file_total <= 0:
            messagebox.showinfo("Clean Mode", "No duplicate content files were found.")
            self.status_text.set("Clean Mode found no duplicate content.")
            return

        removed_folder = self.get_removed_duplicates_folder()
        confirm_text = (
            f"Clean Mode found {duplicate_file_total} duplicate content files across {len(duplicate_groups)} groups.\n\n"
            f"It will keep the newest/best candidate in each group and move the others into:\n\n{removed_folder}\n\n"
            f"Continue?"
        )
        if not messagebox.askokcancel("Clean Mode Confirm", confirm_text):
            self.status_text.set("Clean Mode cancelled.")
            return

        timestamp_root = removed_folder / f"removed_{self.get_timestamp()}"
        timestamp_root.mkdir(parents=True, exist_ok=True)

        removed_items = []
        errors = []

        for file_hash, paths in duplicate_groups.items():
            keep_path = self.choose_keep_file(paths)
            for dup_path in paths:
                if dup_path == keep_path:
                    continue
                relative = dup_path.relative_to(mods_path)
                target = timestamp_root / relative
                target.parent.mkdir(parents=True, exist_ok=True)

                if target.exists():
                    target = target.parent / f"{target.stem}_{self.get_timestamp()}{target.suffix}"

                try:
                    shutil.move(str(dup_path), str(target))
                    removed_items.append({
                        "removed_name": dup_path.name,
                        "kept_name": keep_path.name,
                        "removed_from": str(relative),
                        "kept_at": str(keep_path.relative_to(mods_path)),
                    })
                    self.log(f"Clean Mode moved duplicate: {dup_path} -> {target}")
                except Exception as exc:
                    errors.append(f"{dup_path} -> {exc}")
                    self.log(f"Clean Mode failed moving duplicate: {dup_path} -> {exc}")

        self.config_data["last_removed_duplicates_folder"] = str(timestamp_root)
        self.save_config()

        if errors:
            messagebox.showwarning(
                "Clean Mode Completed With Errors",
                f"Moved {len(removed_items)} duplicate files.\n\nSome files failed:\n\n" + "\n".join(errors[:15])
            )
        else:
            messagebox.showinfo(
                "Clean Mode Complete",
                f"Moved {len(removed_items)} duplicate files into:\n\n{timestamp_root}"
            )

        if removed_items:
            self.summary_var.set(f"Clean Mode moved {len(removed_items)} duplicate files.")
            self.status_text.set(f"Clean Mode complete. Removed {len(removed_items)} duplicate files.")
            self.show_removed_duplicates_window(removed_items, timestamp_root)
        else:
            self.summary_var.set("Clean Mode completed with no files moved.")
            self.status_text.set("Clean Mode completed with no files moved.")

    def get_sims_root_from_mods(self) -> Path | None:
        mods_folder = self.mods_var.get().strip()
        return Path(mods_folder).parent if mods_folder else None

    def prompt_duplicate_action(self, filename):
        choice = messagebox.askyesnocancel(
            "Duplicate Detected",
            f"A duplicate was found:\n\n{filename}\n\nYes = Replace existing file\nNo = Rename incoming copy\nCancel = Skip this file"
        )
        if choice is True:
            return "replace"
        if choice is False:
            return "rename"
        return "skip"

    def make_renamed_path(self, original_path: Path) -> Path:
        stem = original_path.stem
        suffix = original_path.suffix
        parent = original_path.parent
        counter = 1
        while True:
            candidate = parent / f"{stem}_imported_{counter}{suffix}"
            if not candidate.exists():
                return candidate
            counter += 1

    def create_backup(self):
        paths = self.require_paths()
        if not paths:
            return None
        mods_path, _, _ = paths
        backup_root = self.get_backup_folder()
        backup_name = f"mods_backup_{Path(mods_path).name}_{self.get_timestamp()}"
        backup_path = backup_root / backup_name
        try:
            shutil.copytree(mods_path, backup_path)
            self.config_data["last_backup_folder"] = str(backup_path)
            self.last_backup_var.set(str(backup_path))
            self.save_config()
            self.log(f"Backup created: {backup_path}")
            return backup_path
        except Exception as exc:
            messagebox.showerror("Backup Failed", f"Could not create backup.\n\n{exc}")
            self.log(f"Backup failed: {exc}")
            return None

    def get_timestamp(self):
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 ** 2:
            return f"{size_bytes / 1024:.2f} KB"
        if size_bytes < 1024 ** 3:
            return f"{size_bytes / (1024 ** 2):.2f} MB"
        return f"{size_bytes / (1024 ** 3):.2f} GB"

    def html_escape(self, value):
        return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

    # ============================================================
    # Help / socials
    # ============================================================
    def create_help_file(self):
        help_path = self.get_reports_folder() / "help.html"
        social_html = "".join(
            f'<p><a href="{url}" target="_blank"><strong style="font-size:1.2em;">{self.html_escape(name)}</strong></a><br><small>{self.html_escape(url)}</small></p>'
            for name, url in SOCIALS.items()
        )
        extractor_status = "7-Zip detected" if self.has_7z_support() else "7-Zip not detected"
        html = f"""
        <html><head><meta charset=\"utf-8\"><title>{self.html_escape(APP_TITLE)} Help</title>
        <style>
        body {{ font-family: Georgia, 'Times New Roman', serif; background: #0f172a; color: white; padding: 30px; line-height: 1.6; }}
        h1 {{ font-size: 30px; margin-top: 0; }}
        h2 {{ font-size: 22px; }}
        p, li, small {{ font-size: 18px; }}
        a {{ color: #38bdf8; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        .card {{ background: #1e293b; padding: 20px; border-radius: 12px; margin-bottom: 20px; border: 1px solid #334155; }}
        </style></head><body>
        <h1>{self.html_escape(APP_TITLE)} — Help</h1>
        <div class=\"card\"><h2>How to Use</h2>
        <p>1. Set your Mods, Downloads, and Archive folders.</p>
        <p>2. Click <strong>Scan</strong> to inspect incoming files.</p>
        <p>3. Click <strong>Preview</strong> to review actions.</p>
        <p>4. Click <strong>Process</strong> to import valid Sims files.</p>
        <p>5. Click <strong>Finalize</strong> to clear safe Sims 4 cache files.</p>
        <p>6. Click <strong>Create Audit</strong> to generate a HTML report of installed mods.</p>
        <p><strong>Archive support:</strong> ZIP always works. RAR / 7Z work when 7-Zip is installed. Current status: {self.html_escape(extractor_status)}.</p>
        </div>
        <div class=\"card\"><h2>Creator Links</h2>{social_html}</div>
        </body></html>
        """
        help_path.write_text(html, encoding="utf-8")
        return help_path

    def open_help_file(self):
        try:
            help_file = self.create_help_file()
            webbrowser.open(help_file.as_uri())
            self.status_text.set("Opened help guide.")
        except Exception as exc:
            messagebox.showerror("Help Failed", f"Could not open help file.\n\n{exc}")

    # ============================================================
    # Actions
    # ============================================================
    def refresh_status(self):
        self.update_config_from_vars()
        live = self.config_data.get("live_mods_folder", "")
        current = self.mods_var.get().strip()
        self.mode_var.set("TEST MODE" if live and current and Path(current) != Path(live) else "LIVE MODE")
        self.locked_var.set("LOCKED" if self.config_data.get("is_locked") else "UNLOCKED")
        if hasattr(self, "lock_toggle_btn"):
            self.lock_toggle_btn.configure(text=BTN_UNLOCK if self.config_data.get("is_locked") else BTN_LOCK)
        self.status_text.set(f"Preset: {self.current_preset_var.get()} | {self.mode_var.get()} | {self.locked_var.get()}")

    def run_setup_wizard_again(self):
        self.run_setup_wizard()
        self.apply_config_to_vars()
        self.refresh_status()
        self.status_text.set("Setup wizard completed again.")

    def save_settings_action(self):
        self.update_config_from_vars()
        self.status_text.set("Settings saved.")
        messagebox.showinfo("Saved", "Settings saved.")

    def pick_path(self, kind):
        chosen = filedialog.askdirectory(title=f"Select {kind.title()} Folder")
        if not chosen:
            return
        if kind == "mods":
            self.mods_var.set(chosen)
        elif kind == "downloads":
            self.downloads_var.set(chosen)
        elif kind == "archive":
            self.archive_var.set(chosen)
        elif kind == "removed":
            self.removed_dupes_var.set(chosen)
        self.refresh_status()
        self.status_text.set(f"{kind.title()} folder updated.")

    def open_folder(self, path):
        if not path:
            messagebox.showwarning("Missing Path", "No folder is set yet.")
            return
        p = Path(path)
        p.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(p))
            elif sys.platform == "darwin":
                os.system(f'open "{p}"')
            else:
                os.system(f'xdg-open "{p}"')
        except Exception as exc:
            messagebox.showerror("Open Folder Failed", str(exc))

    def switch_to_live_preset(self):
        live = self.config_data.get("live_mods_folder", "")
        if not live:
            messagebox.showwarning("Missing Live Preset", "No live Mods folder has been saved yet.")
            return
        self.mods_var.set(live)
        self.current_preset_var.set("Live")
        self.config_data["current_preset"] = "Live"
        self.refresh_status()
        self.status_text.set("Switched to Live preset.")

    def switch_to_test_preset(self):
        test_path = self.config_data.get("test_mods_folder", "")
        if not test_path:
            chosen = filedialog.askdirectory(title="Choose Test Mods Folder")
            if not chosen:
                return
            self.config_data["test_mods_folder"] = chosen
            test_path = chosen
        self.mods_var.set(test_path)
        self.current_preset_var.set("Test")
        self.config_data["current_preset"] = "Test"
        self.refresh_status()
        self.status_text.set("Switched to Test preset.")

    def lock_current_folder(self):
        self.config_data["is_locked"] = True
        self.save_config()
        self.refresh_status()
        self.status_text.set("Current Mods folder locked inside app.")

    def unlock_current_folder(self):
        self.config_data["is_locked"] = False
        self.save_config()
        self.refresh_status()
        self.status_text.set("Current Mods folder unlocked inside app.")

    def toggle_lock(self):
        if self.config_data.get("is_locked"):
            self.unlock_current_folder()
        else:
            self.lock_current_folder()

    def show_lock_status(self):
        messagebox.showinfo("Lock Status", f"Current state: {self.locked_var.get()}")

    def show_about(self):
        try:
            help_file = self.create_help_file()
            webbrowser.open(help_file.as_uri())
        except Exception:
            pass
        social_lines = "\n".join(f"{name}: {url}" for name, url in SOCIALS.items())
        messagebox.showinfo(f"About {APP_TITLE}", f"{APP_TITLEBAR}\n\nHelp guide opened in browser.\n\n{social_lines}")

    def analyze_link(self):
        url = self.url_var.get().strip()
        label = self.url_label_var.get().strip()
        if not url:
            messagebox.showwarning("Missing Link", "Paste a link or file URL first.")
            return
        info = [f"URL: {url}", f"Label: {label if label else 'None'}"]
        lower_url = url.lower()
        if lower_url.endswith(".zip"):
            info.append("Detected type: Direct ZIP file")
        elif lower_url.endswith(".package"):
            info.append("Detected type: Direct PACKAGE file")
        elif lower_url.endswith(".ts4script"):
            info.append("Detected type: Direct TS4SCRIPT file")
        elif lower_url.endswith(".rar") or lower_url.endswith(".7z"):
            info.append("Detected type: Direct RAR / 7Z file")
            info.append("Will process if 7-Zip is installed.")
        else:
            info.append("Detected type: Page / unknown link")
            info.append("This will likely need browser or later parser support.")
        messagebox.showinfo("Link Analysis", "\n".join(info))
        self.status_text.set("Link analyzed.")

    def download_file(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning("Missing Link", "Paste a link first.")
            return
        messagebox.showinfo("Download Placeholder", "Direct download wiring is still not added in this build.")
        self.status_text.set("Download placeholder triggered.")

    def open_link_in_browser(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning("Missing Link", "Paste a link first.")
            return
        try:
            webbrowser.open(url)
            self.status_text.set("Opened link in browser.")
        except Exception as exc:
            messagebox.showerror("Open Link Failed", str(exc))

    def scan_downloads(self):
        paths = self.require_paths()
        if not paths:
            return
        mods_path, downloads_path, _ = paths
        self.scan_items = []
        self.clear_preview()

        existing_name_map = {}
        existing_hash_map = {}
        for p in mods_path.rglob("*"):
            if p.is_file() and p.suffix.lower() in VALID_MOD_EXTENSIONS:
                existing_name_map.setdefault(p.name.lower(), []).append(p)
                h = self.hash_file(p)
                if h:
                    existing_hash_map.setdefault(h, []).append(p)

        archives = 0
        valid_mods = 0
        junk = 0
        duplicate_name_count = 0
        duplicate_content_count = 0
        unsupported_archives = 0

        self.log("Starting scan of Downloads folder...")

        for file_path in downloads_path.rglob("*"):
            if not file_path.is_file():
                continue
            ext = file_path.suffix.lower()
            destination = str(mods_path)
            action = "Skip"
            file_type = "junk"

            if ext in VALID_ARCHIVE_EXTENSIONS:
                archives += 1
                file_type = self.get_archive_support_text(ext)
                if file_type == "archive-missing-tool":
                    action = "Needs 7-Zip"
                    unsupported_archives += 1
                else:
                    action = "Archive"
            elif ext in VALID_MOD_EXTENSIONS:
                valid_mods += 1
                file_type = ext
                name_dup = file_path.name.lower() in existing_name_map
                file_hash = self.hash_file(file_path)
                content_dup = bool(file_hash and file_hash in existing_hash_map)
                if name_dup and content_dup:
                    duplicate_name_count += 1
                    duplicate_content_count += 1
                    action = "Dup Name+Hash"
                elif name_dup:
                    duplicate_name_count += 1
                    action = "Dup Name"
                elif content_dup:
                    duplicate_content_count += 1
                    action = "Dup Hash"
                else:
                    action = "Import"
            else:
                junk += 1
                file_type = "junk"
                action = "Skip"

            item = {
                "name": file_path.name,
                "type": file_type,
                "source": str(file_path),
                "destination": destination,
                "action": action,
            }
            self.scan_items.append(item)
            self.add_preview_row(item["name"], item["type"], item["source"], item["destination"], item["action"])

        summary_text = (
            f"Files scanned: {len(self.scan_items)} | Archives: {archives} | Valid mods: {valid_mods} | "
            f"Junk: {junk} | Dup names: {duplicate_name_count} | Dup content: {duplicate_content_count} | "
            f"Unsupported archives: {unsupported_archives}"
        )
        self.summary_var.set(summary_text)
        self.log("Scan complete.")
        self.log(summary_text)
        if not self.scan_items:
            messagebox.showinfo("Scan Complete", "No files were found in the selected Downloads folder.")
        else:
            messagebox.showinfo("Scan Complete", summary_text)

    def preview_import(self):
        if not self.scan_items:
            self.scan_downloads()
            if not self.scan_items:
                return
        imports = sum(1 for item in self.scan_items if item["action"] == "Import")
        warnings = sum(1 for item in self.scan_items if "Dup" in item["action"] or item["action"] == "Needs 7-Zip")
        archives = sum(1 for item in self.scan_items if str(item["type"]).startswith("archive"))
        skipped = sum(1 for item in self.scan_items if item["action"] == "Skip")
        msg = (
            f"Preview ready.\n\nImport: {imports}\nWarnings: {warnings}\nArchives: {archives}\nSkipped: {skipped}\n\n"
            f"Review the preview table for full details."
        )
        self.log("Preview refreshed.")
        messagebox.showinfo("Import Preview", msg)

    def process_downloads(self):
        if self.config_data.get("is_locked"):
            messagebox.showwarning("Folder Locked", "Unlock the current Mods folder before processing downloads.")
            return
        paths = self.require_paths()
        if not paths:
            return
        mods_path, _, archive_path = paths
        if not self.scan_items:
            self.scan_downloads()
            if not self.scan_items:
                return

        if self.config_data.get("always_prompt_backup_before_import", True):
            choice = messagebox.askyesnocancel(
                "Backup Before Import",
                "Create backup before import?\n\nYes = Create backup now\nNo = Continue without backup\nCancel = Stop processing"
            )
            if choice is None:
                self.log("Process cancelled by user.")
                return
            if choice and self.create_backup() is None:
                return

        temp_extract = self.get_temp_extract_folder()
        if temp_extract.exists():
            shutil.rmtree(temp_extract, ignore_errors=True)
        temp_extract.mkdir(parents=True, exist_ok=True)

        imported_count = 0
        archived_count = 0
        skipped_count = 0
        extracted_count = 0
        duplicate_count = 0

        existing_hashes = {}
        for p in mods_path.rglob("*"):
            if p.is_file() and p.suffix.lower() in VALID_MOD_EXTENSIONS:
                h = self.hash_file(p)
                if h:
                    existing_hashes.setdefault(h, []).append(p)

        self.log("Starting download processing...")

        def import_single_file(source_file: Path):
            nonlocal imported_count, skipped_count, duplicate_count
            dest_file = mods_path / source_file.name
            source_hash = self.hash_file(source_file)
            same_content_exists = bool(source_hash and source_hash in existing_hashes)

            if same_content_exists:
                duplicate_count += 1
                action = self.prompt_duplicate_action(f"{source_file.name} (same content)")
            elif dest_file.exists():
                duplicate_count += 1
                action = self.prompt_duplicate_action(source_file.name)
            else:
                action = "import"

            if action == "skip":
                skipped_count += 1
                self.log(f"Skipped duplicate: {source_file.name}")
                return
            if action == "replace":
                try:
                    if dest_file.exists():
                        dest_file.unlink()
                    shutil.copy2(source_file, dest_file)
                    imported_count += 1
                    if source_hash:
                        existing_hashes.setdefault(source_hash, []).append(dest_file)
                    self.log(f"Replaced existing file: {dest_file.name}")
                except Exception as exc:
                    skipped_count += 1
                    self.log(f"Failed replacing file {source_file.name}: {exc}")
                return
            if action == "rename":
                renamed_dest = self.make_renamed_path(dest_file)
                try:
                    shutil.copy2(source_file, renamed_dest)
                    imported_count += 1
                    if source_hash:
                        existing_hashes.setdefault(source_hash, []).append(renamed_dest)
                    self.log(f"Imported renamed copy: {renamed_dest.name}")
                except Exception as exc:
                    skipped_count += 1
                    self.log(f"Failed importing renamed file {source_file.name}: {exc}")
                return

            try:
                shutil.copy2(source_file, dest_file)
                imported_count += 1
                if source_hash:
                    existing_hashes.setdefault(source_hash, []).append(dest_file)
                self.log(f"Imported file: {dest_file.name}")
            except Exception as exc:
                skipped_count += 1
                self.log(f"Failed importing file {source_file.name}: {exc}")

        for item in self.scan_items:
            source_path = Path(item["source"])
            if not source_path.exists():
                skipped_count += 1
                self.log(f"Missing source file, skipped: {source_path}")
                continue
            ext = source_path.suffix.lower()

            if ext in VALID_ARCHIVE_EXTENSIONS:
                extract_target = temp_extract / source_path.stem
                extract_target.mkdir(parents=True, exist_ok=True)
                ok, detail = self.extract_archive(source_path, extract_target)
                if not ok:
                    skipped_count += 1
                    self.log(f"Failed extracting archive {source_path.name}: {detail}")
                    continue
                extracted_count += 1
                self.log(f"Extracted archive: {source_path.name} ({detail})")
                for extracted_file in extract_target.rglob("*"):
                    if extracted_file.is_file() and extracted_file.suffix.lower() in VALID_MOD_EXTENSIONS:
                        import_single_file(extracted_file)
                try:
                    archive_dest = archive_path / source_path.name
                    if archive_dest.exists():
                        archive_dest = archive_path / f"{source_path.stem}_{self.get_timestamp()}{source_path.suffix}"
                    shutil.move(str(source_path), str(archive_dest))
                    archived_count += 1
                    self.log(f"Archived source archive: {archive_dest.name}")
                except Exception as exc:
                    self.log(f"Failed archiving source archive {source_path.name}: {exc}")
            elif ext in VALID_MOD_EXTENSIONS:
                import_single_file(source_path)
                try:
                    archive_dest = archive_path / source_path.name
                    if archive_dest.exists():
                        archive_dest = archive_path / f"{source_path.stem}_{self.get_timestamp()}{source_path.suffix}"
                    shutil.move(str(source_path), str(archive_dest))
                    archived_count += 1
                    self.log(f"Archived source mod file: {archive_dest.name}")
                except Exception as exc:
                    self.log(f"Failed archiving source mod file {source_path.name}: {exc}")
            else:
                skipped_count += 1

        summary_text = (
            f"Imported: {imported_count} | Archived: {archived_count} | Extracted archives: {extracted_count} | "
            f"Duplicates handled: {duplicate_count} | Skipped: {skipped_count}"
        )
        self.summary_var.set(summary_text)
        self.log("Processing complete.")
        self.log(summary_text)
        if temp_extract.exists():
            shutil.rmtree(temp_extract, ignore_errors=True)
        self.scan_downloads()
        messagebox.showinfo("Process Complete", summary_text)

    def finalize_cleanup(self):
        sims_root = self.get_sims_root_from_mods()
        if not sims_root:
            messagebox.showerror("Missing Mods Folder", "Set your Mods folder first.")
            return
        if not sims_root.exists():
            messagebox.showerror("Invalid Sims Folder", "The folder above your Mods folder does not exist.")
            return
        if not messagebox.askyesno("Finalize", "Clear safe Sims 4 cache files now?\n\nThis will remove:\n- localthumbcache.package\n- cache\n- cachestr\n- onlinethumbnailcache"):
            self.log("Finalize cancelled by user.")
            return
        removed_files = 0
        removed_folders = 0
        missing_items = 0
        for filename in SAFE_CACHE_FILES:
            target = sims_root / filename
            if target.exists() and target.is_file():
                try:
                    target.unlink()
                    removed_files += 1
                    self.log(f"Removed cache file: {target}")
                except Exception as exc:
                    self.log(f"Failed removing cache file {target}: {exc}")
            else:
                missing_items += 1
        for folder_name in SAFE_CACHE_FOLDERS:
            target = sims_root / folder_name
            if target.exists() and target.is_dir():
                try:
                    shutil.rmtree(target, ignore_errors=True)
                    removed_folders += 1
                    self.log(f"Removed cache folder: {target}")
                except Exception as exc:
                    self.log(f"Failed removing cache folder {target}: {exc}")
            else:
                missing_items += 1
        summary_text = f"Finalize complete | Files removed: {removed_files} | Folders removed: {removed_folders} | Already missing: {missing_items}"
        self.summary_var.set(summary_text)
        self.log(summary_text)
        messagebox.showinfo("Finalize Complete", summary_text)

    def create_mod_audit(self):
        mods_folder = self.mods_var.get().strip()
        if not mods_folder:
            messagebox.showerror("Missing Mods Folder", "Set your Mods folder first.")
            return
        mods_path = Path(mods_folder)
        if not mods_path.exists():
            messagebox.showerror("Missing Mods Folder", "The selected Mods folder does not exist.")
            return

        reports_folder = self.get_reports_folder()
        timestamp = self.get_timestamp()
        audit_path = reports_folder / f"mod_audit_{timestamp}.html"
        self.log("Starting mod audit...")

        files = []
        duplicate_name_map = {}
        duplicate_hash_map = {}
        total_files = 0
        total_packages = 0
        total_scripts = 0
        total_size = 0

        for file_path in mods_path.rglob("*"):
            if not file_path.is_file():
                continue
            total_files += 1
            ext = file_path.suffix.lower()
            if ext == ".package":
                total_packages += 1
            elif ext == ".ts4script":
                total_scripts += 1
            try:
                size = file_path.stat().st_size
            except Exception:
                size = 0
            total_size += size
            rel_path = file_path.relative_to(mods_path)
            modified = datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            duplicate_name_map.setdefault(file_path.name.lower(), []).append(str(rel_path))
            if ext in VALID_MOD_EXTENSIONS:
                h = self.hash_file(file_path)
                if h:
                    duplicate_hash_map.setdefault(h, []).append(str(rel_path))
            files.append({
                "name": file_path.name,
                "ext": ext,
                "relative_path": str(rel_path),
                "size": size,
                "size_text": self.format_size(size),
                "modified": modified,
                "hash": self.hash_file(file_path) if ext in VALID_MOD_EXTENSIONS else None,
            })

        duplicate_names = {name: paths for name, paths in duplicate_name_map.items() if len(paths) > 1}
        duplicate_hashes = {h: paths for h, paths in duplicate_hash_map.items() if len(paths) > 1}
        duplicate_count = len(duplicate_names)
        duplicate_hash_count = len(duplicate_hashes)
        files.sort(key=lambda x: x["relative_path"].lower())

        duplicate_name_rows = "".join(
            f"<tr><td>{self.html_escape(name)}</td><td>{'<br>'.join(self.html_escape(p) for p in paths)}</td></tr>"
            for name, paths in sorted(duplicate_names.items())
        ) or "<tr><td colspan='2'>No duplicate filenames were found.</td></tr>"

        duplicate_hash_rows = "".join(
            f"<tr><td>{self.html_escape(h)}</td><td>{'<br>'.join(self.html_escape(p) for p in paths)}</td></tr>"
            for h, paths in sorted(duplicate_hashes.items())
        ) or "<tr><td colspan='2'>No duplicate file contents were found.</td></tr>"

        file_rows = []
        for item in files:
            dup_name_flag = "Yes" if item["name"].lower() in duplicate_names else "No"
            dup_hash_flag = "Yes" if item["hash"] and item["hash"] in duplicate_hashes else "No"
            script_flag = "Yes" if item["ext"] == ".ts4script" else "No"
            file_rows.append(f"""
                <tr>
                    <td>{self.html_escape(item['name'])}</td>
                    <td>{self.html_escape(item['ext'])}</td>
                    <td>{self.html_escape(item['relative_path'])}</td>
                    <td>{self.html_escape(item['size_text'])}</td>
                    <td>{self.html_escape(item['modified'])}</td>
                    <td>{dup_name_flag}</td>
                    <td>{dup_hash_flag}</td>
                    <td>{script_flag}</td>
                </tr>
            """)

        html = f"""<!DOCTYPE html>
<html lang=\"en\"><head><meta charset=\"UTF-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
<title>{self.html_escape(APP_TITLE)} Audit</title>
<style>
body {{ font-family: Georgia, 'Times New Roman', serif; background: #111827; color: #e5e7eb; margin: 0; padding: 24px; }}
.container {{ max-width: 1400px; margin: 0 auto; }}
h1, h2, h3 {{ margin-top: 0; }}
.meta {{ margin-bottom: 20px; color: #cbd5e1; font-size: 18px; }}
.grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 24px; }}
.card {{ background: #1f2937; border: 1px solid #374151; border-radius: 14px; padding: 16px; }}
.card p {{ font-size: 1.5rem; font-weight: bold; margin: 8px 0 0 0; }}
.section {{ background: #1f2937; border: 1px solid #374151; border-radius: 14px; padding: 18px; margin-bottom: 24px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 1rem; }}
th, td {{ border-bottom: 1px solid #374151; padding: 10px; text-align: left; vertical-align: top; }}
th {{ background: #0f172a; }}
tr:hover {{ background: #172033; }}
</style></head><body>
<div class=\"container\">
<h1>{self.html_escape(APP_TITLE)} — Installed Mods Audit</h1>
<div class=\"meta\">
<div><strong>Version:</strong> {self.html_escape(APP_VERSION)}</div>
<div><strong>Generated:</strong> {self.html_escape(timestamp)}</div>
<div><strong>Mods Folder:</strong> {self.html_escape(str(mods_path))}</div>
</div>
<div class=\"grid\">
<div class=\"card\"><h3>Total Files</h3><p>{total_files}</p></div>
<div class=\"card\"><h3>.package</h3><p>{total_packages}</p></div>
<div class=\"card\"><h3>.ts4script</h3><p>{total_scripts}</p></div>
<div class=\"card\"><h3>Total Size</h3><p>{self.format_size(total_size)}</p></div>
<div class=\"card\"><h3>Dup Names</h3><p>{duplicate_count}</p></div>
<div class=\"card\"><h3>Dup Content</h3><p>{duplicate_hash_count}</p></div>
</div>
<div class=\"section\"><h2>Duplicate Filename Warnings</h2><table><thead><tr><th>Filename</th><th>Locations</th></tr></thead><tbody>{duplicate_name_rows}</tbody></table></div>
<div class=\"section\"><h2>Duplicate Content Warnings</h2><table><thead><tr><th>Hash</th><th>Locations</th></tr></thead><tbody>{duplicate_hash_rows}</tbody></table></div>
<div class=\"section\"><h2>Installed Files</h2><table><thead><tr>
<th>File Name</th><th>Extension</th><th>Relative Path</th><th>Size</th><th>Modified</th><th>Dup Name?</th><th>Dup Content?</th><th>Script Mod?</th>
</tr></thead><tbody>{''.join(file_rows)}</tbody></table></div>
</div></body></html>"""

        try:
            audit_path.write_text(html, encoding="utf-8")
            self.config_data["last_audit_file"] = str(audit_path)
            self.last_audit_var.set(str(audit_path))
            self.save_config()
            summary_text = (
                f"Audit created | Files: {total_files} | Packages: {total_packages} | Scripts: {total_scripts} | "
                f"Dup Names: {duplicate_count} | Dup Content: {duplicate_hash_count}"
            )
            self.summary_var.set(summary_text)
            self.log(f"Audit created: {audit_path}")
            self.log(summary_text)
            webbrowser.open(audit_path.as_uri())
            messagebox.showinfo("Audit Created", f"HTML audit created successfully.\n\n{audit_path}")
        except Exception as exc:
            messagebox.showerror("Audit Failed", f"Could not create audit.\n\n{exc}")
            self.log(f"Audit failed: {exc}")

    def open_last_audit(self):
        last_audit = self.config_data.get("last_audit_file", "")
        if not last_audit:
            messagebox.showwarning("No Audit", "No audit file has been created yet.")
            return
        audit_path = Path(last_audit)
        if not audit_path.exists():
            messagebox.showerror("Missing Audit", "The saved audit file no longer exists.")
            return
        try:
            webbrowser.open(audit_path.as_uri())
            self.status_text.set("Opened last audit.")
        except Exception as exc:
            messagebox.showerror("Open Audit Failed", str(exc))

    def export_log(self):
        reports_folder = self.get_reports_folder()
        log_path = reports_folder / f"log_{self.get_timestamp()}.txt"
        try:
            content = self.log_text.get("1.0", "end").strip()
            log_path.write_text(content, encoding="utf-8")
            self.log(f"Log exported: {log_path}")
            messagebox.showinfo("Log Exported", f"Log saved to:\n\n{log_path}")
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Could not export log.\n\n{exc}")

    def placeholder_action(self):
        self.status_text.set("This feature is not wired yet in this build.")
        messagebox.showinfo("Coming Soon", "This feature is not wired yet in this build.")


if __name__ == "__main__":
    app = TS4ModMan3ger()
    app.mainloop()
