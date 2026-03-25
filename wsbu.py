# =============================================================================
#  Windhawk Service Management Utility - Version 2.5.0
#  Author: scorpion421
#  Description: A tool for backing up and restoring Windhawk configurations.
# =============================================================================

import ctypes
import datetime
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import winreg
import zipfile
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk
from tkinter import ttk

# ---------------------------------------------------------------------------
# Application constants
# ---------------------------------------------------------------------------
APP_VERSION   = "2.5.0"
APP_TITLE     = f"Windhawk Service Management Utility v{APP_VERSION}"

WINDHAWK_REGISTRY_KEY   = r"SOFTWARE\Windhawk"
WINDHAWK_SERVICE_NAME   = "Windhawk"
WINDHAWK_ROOT_SENTINELS = ("ModsSource", os.path.join("Engine", "Mods"), "windhawk.exe")

DEFAULT_WINDHAWK_ROOT = os.path.expandvars(r"%programdata%\Windhawk")
DEFAULT_BACKUP_FOLDER = os.path.expandvars(r"%userprofile%\Documents\Windhawk_Backup")
DEFAULT_MAX_BACKUPS   = 10

CONFIG_DIR  = os.path.expandvars(r"%appdata%\Windhawk_Backup_Utility")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

PAD = 8  # Universal spacing unit used throughout the UI

# =============================================================================
#                            CORE LOGIC (BACKEND)
# =============================================================================

# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------

def load_config() -> dict:
    """Loads settings from the JSON config file, merging with built-in defaults."""
    defaults = {
        "windhawk_root": DEFAULT_WINDHAWK_ROOT,
        "backup_folder": DEFAULT_BACKUP_FOLDER,
        "portable":      False,
        "max_backups":   DEFAULT_MAX_BACKUPS,
    }
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
            stored = json.load(fh)
        defaults.update(stored)
    except (OSError, json.JSONDecodeError):
        pass
    return defaults


def save_config(cfg: dict) -> None:
    """Persists settings to the JSON config file. Failure is non-fatal."""
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Backup catalogue helpers
# ---------------------------------------------------------------------------

def _format_size(size_bytes: int) -> str:
    """Formats a byte count as a human-readable string (B / KB / MB / GB)."""
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes //= 1024
    return f"{size_bytes:.1f} TB"


def list_backups(backup_folder: str) -> list[dict]:
    """
    Scans the backup folder for archives and returns metadata for each,
    newest first. Reads manifest.json from inside each ZIP if available.
    """
    results: list[dict] = []
    if not os.path.isdir(backup_folder):
        return results

    names = sorted(
        (n for n in os.listdir(backup_folder)
         if n.startswith("windhawk-backup_") and n.endswith(".zip")),
        reverse=True,
    )
    for name in names:
        full_path = os.path.join(backup_folder, name)
        try:
            size  = os.path.getsize(full_path)
            mtime = os.path.getmtime(full_path)
            dt    = datetime.datetime.fromtimestamp(mtime)

            manifest:  dict     = {}
            mod_count: int | None = None
            try:
                with zipfile.ZipFile(full_path, "r") as zf:
                    names = zf.namelist()
                    if "manifest.json" in names:
                        manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
                    else:
                        # Legacy archive: normalize separators first (Windows
                        # shutil.make_archive uses backslashes in ZIP entries).
                        normalized = [n.replace("\\", "/") for n in names]
                        mod_count = sum(
                            1 for n in normalized
                            if n.startswith("ModsSource/") and n.endswith(".wh.cpp")
                        )
            except Exception:
                pass

            # Prefer manifest value, fall back to counted value, then "-"
            mods_display = str(
                manifest.get("mod_count", mod_count)
                if "mod_count" in manifest or mod_count is not None
                else "-"
            )

            results.append({
                "name":  name,
                "path":  full_path,
                "date":  dt.strftime("%Y-%m-%d  %H:%M:%S"),
                "size":  _format_size(size),
                "kind":  "Portable" if manifest.get("portable") else "Standard",
                "mods":  mods_display,
            })
        except OSError:
            continue
    return results


def create_manifest(windhawk_root: str, portable: bool) -> dict:
    """Builds a metadata dict to be stored as manifest.json inside the archive."""
    mods: list[str] = []
    mods_dir = os.path.join(windhawk_root, "ModsSource")
    if os.path.isdir(mods_dir):
        mods = [f for f in os.listdir(mods_dir) if f.endswith(".wh.cpp")]
    # Strip the .wh.cpp suffix for readability in the manifest mod list.
    mod_names = [f[:-7] for f in mods]
    return {
        "app_version":   APP_VERSION,
        "created":       datetime.datetime.now().isoformat(timespec="seconds"),
        "windhawk_root": windhawk_root,
        "portable":      portable,
        "mods":          mod_names,
        "mod_count":     len(mod_names),
    }


def rotate_backups(backup_folder: str, max_backups: int) -> list[str]:
    """
    Deletes the oldest backup archives when the total exceeds max_backups.
    A value of 0 disables rotation entirely.
    Returns the list of deleted filenames.
    """
    if max_backups <= 0 or not os.path.isdir(backup_folder):
        return []
    archives = sorted(
        f for f in os.listdir(backup_folder)
        if f.startswith("windhawk-backup_") and f.endswith(".zip")
    )
    to_delete = archives[:-max_backups] if len(archives) > max_backups else []
    deleted: list[str] = []
    for name in to_delete:
        try:
            os.remove(os.path.join(backup_folder, name))
            deleted.append(name)
        except OSError:
            pass
    return deleted


# ---------------------------------------------------------------------------
# System helpers
# ---------------------------------------------------------------------------

def registry_key_exists(key_path: str) -> bool:
    """Returns True if the given HKLM registry key exists, False otherwise."""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path):
            return True
    except OSError:
        return False


def is_admin() -> bool:
    """Returns True if the process is running with administrator privileges."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def run_as_admin() -> bool:
    """
    Re-launches this script with elevated privileges via ShellExecute.
    Paths containing spaces are quoted correctly.
    Returns True if the elevation request was submitted successfully.
    """
    try:
        args   = " ".join(f'"{a}"' for a in sys.argv)
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, args, None, 1
        )
        return int(result) > 32
    except Exception:
        return False


def validate_windhawk_root(path: str) -> bool:
    """
    Returns True if at least one known Windhawk sentinel exists inside the
    given root path, preventing operations on obviously wrong directories.
    """
    return any(os.path.exists(os.path.join(path, s)) for s in WINDHAWK_ROOT_SENTINELS)


def _run_sc(action: str) -> tuple[bool, str]:
    """Runs 'sc <action> <service>' and returns (success, combined_output)."""
    try:
        r = subprocess.run(
            ["sc", action, WINDHAWK_SERVICE_NAME],
            capture_output=True, text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return r.returncode == 0, (r.stdout + r.stderr).strip()
    except OSError as exc:
        return False, str(exc)


def stop_windhawk_service() -> tuple[bool, str]:
    """Stops the Windhawk Windows service. Returns (success, message)."""
    ok, out = _run_sc("stop")
    if ok:
        return True, "Status: Windhawk service stopped."
    if "1062" in out or "not started" in out.lower():
        return True, "Info: Windhawk service was not running - no action needed."
    return False, f"Warning: Could not stop Windhawk service: {out}"


def start_windhawk_service() -> tuple[bool, str]:
    """Starts the Windhawk Windows service. Returns (success, message)."""
    ok, out = _run_sc("start")
    if ok:
        return True, "Status: Windhawk service restarted."
    return False, f"Warning: Could not restart Windhawk service: {out}"


# ---------------------------------------------------------------------------
# Backup / Restore operations
# ---------------------------------------------------------------------------

def execute_backup_operation(
    windhawk_root: str,
    backup_folder: str,
    portable:      bool = False,
    max_backups:   int  = DEFAULT_MAX_BACKUPS,
) -> tuple[bool, str]:
    """
    Backs up Windhawk mod sources, compiled mods, a manifest.json, and
    (unless portable) the registry key into a timestamped ZIP archive.

    The Windhawk service is stopped before file access and restarted
    afterwards via try/finally. The resulting archive is validated with
    zipfile.testzip(). Old backups are rotated if max_backups > 0.

    Returns (success, log_text).
    """
    log: list[str] = []

    if not validate_windhawk_root(windhawk_root):
        return False, (
            f"ERROR: Not a valid Windhawk installation:\n{windhawk_root}\n"
            f"Expected at least one of: {', '.join(WINDHAWK_ROOT_SENTINELS)}"
        )

    try:
        os.makedirs(backup_folder, exist_ok=True)
    except OSError as exc:
        return False, f"ERROR: Could not create backup folder: {exc}"

    timestamp    = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_base = os.path.join(backup_folder, f"windhawk-backup_{timestamp}")

    if not portable:
        ok, msg = stop_windhawk_service()
        log.append(msg)
        if not ok:
            log.append("Warning: Proceeding despite service stop issue. Files may be locked.")

    try:
        with tempfile.TemporaryDirectory() as stage_dir:

            # Step 1 - Stage mod directories
            for rel, src in {
                "ModsSource":                   os.path.join(windhawk_root, "ModsSource"),
                os.path.join("Engine", "Mods"): os.path.join(windhawk_root, "Engine", "Mods"),
            }.items():
                dst = os.path.join(stage_dir, rel)
                if os.path.isdir(src):
                    try:
                        shutil.copytree(src, dst)
                        log.append(f"Status: '{rel}' staged.")
                    except OSError as exc:
                        log.append(f"Warning: Could not stage '{rel}': {exc}")
                else:
                    log.append(f"Warning: Not found, skipping: {src}")

            # Step 2 - Write manifest
            try:
                manifest_path = os.path.join(stage_dir, "manifest.json")
                with open(manifest_path, "w", encoding="utf-8") as fh:
                    json.dump(create_manifest(windhawk_root, portable), fh, indent=2)
                log.append("Status: Manifest written.")
            except OSError as exc:
                log.append(f"Warning: Could not write manifest: {exc}")

            # Step 3 - Export registry key
            if portable:
                log.append("Info: Portable mode - registry export skipped.")
            else:
                reg_file = os.path.join(stage_dir, "Windhawk.reg")
                try:
                    subprocess.run(
                        ["reg", "export", f"HKLM\\{WINDHAWK_REGISTRY_KEY}", reg_file, "/y"],
                        check=True, capture_output=True, text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW,
                    )
                    log.append("Status: Registry exported.")
                except subprocess.CalledProcessError as exc:
                    log.append(f"ERROR: Registry export failed: {exc.stderr.strip()}")
                    return False, "\n".join(log)

            # Step 4 - Create archive
            try:
                shutil.make_archive(archive_base, "zip", stage_dir)
            except OSError as exc:
                log.append(f"ERROR: Archive creation failed: {exc}")
                return False, "\n".join(log)

            # Step 5 - Validate archive integrity
            archive_path = f"{archive_base}.zip"
            try:
                with zipfile.ZipFile(archive_path, "r") as zf:
                    bad = zf.testzip()
                if bad is not None:
                    log.append(f"ERROR: Archive corrupt - bad entry: {bad}")
                    return False, "\n".join(log)
                log.append("Status: Archive integrity verified.")
            except zipfile.BadZipFile as exc:
                log.append(f"ERROR: Archive is not a valid ZIP: {exc}")
                return False, "\n".join(log)

            log.append(f"\nOperation Complete: Archive created at:\n{archive_path}")

    except OSError as exc:
        log.append(f"ERROR: Staging directory error: {exc}")
        return False, "\n".join(log)

    finally:
        if not portable:
            _, msg = start_windhawk_service()
            log.append(msg)

    # Step 6 - Rotate old backups
    deleted = rotate_backups(backup_folder, max_backups)
    for name in deleted:
        log.append(f"Info: Rotation - deleted old backup: {name}")

    return True, "\n".join(log)


def _resolve_nested_source(path: str) -> tuple[str, bool]:
    """
    Detects and resolves one level of same-name nesting inside a directory.

    If 'path' contains a direct subdirectory whose name matches the last
    component of 'path' (e.g. ModsSource/ModsSource or Engine/Mods/Mods),
    that subdirectory is returned as the real source so the restore does not
    reproduce the nesting on disk.

    Returns (resolved_path, was_nested).
    """
    folder_name = os.path.basename(path)
    nested = os.path.join(path, folder_name)
    if os.path.isdir(nested):
        return nested, True
    return path, False


def execute_restore_operation(
    windhawk_root: str,
    archive_path:  str,
    portable:      bool = False,
) -> tuple[bool, str]:
    """
    Restores mod sources, compiled mods, and (unless portable) registry
    settings from a previously created ZIP archive.

    The Windhawk service is stopped before file access and restarted
    afterwards via try/finally.

    Returns (success, log_text).
    """
    log: list[str] = []

    if not validate_windhawk_root(windhawk_root):
        return False, (
            f"ERROR: Not a valid Windhawk installation:\n{windhawk_root}\n"
            f"Expected at least one of: {', '.join(WINDHAWK_ROOT_SENTINELS)}"
        )

    if not portable:
        ok, msg = stop_windhawk_service()
        log.append(msg)
        if not ok:
            log.append("Warning: Proceeding despite service stop issue. Files may be locked.")

    try:
        with tempfile.TemporaryDirectory() as stage_dir:

            # Step 1 - Extract archive
            try:
                shutil.unpack_archive(archive_path, stage_dir)
                log.append(f"Status: '{os.path.basename(archive_path)}' extracted.")
            except Exception as exc:
                log.append(f"ERROR: Extraction failed: {exc}")
                return False, "\n".join(log)

            # Step 2 - Restore mod directories
            for label, (src, dst) in {
                "ModsSource": (
                    os.path.join(stage_dir, "ModsSource"),
                    os.path.join(windhawk_root, "ModsSource"),
                ),
                os.path.join("Engine", "Mods"): (
                    os.path.join(stage_dir, "Engine", "Mods"),
                    os.path.join(windhawk_root, "Engine", "Mods"),
                ),
            }.items():
                if os.path.isdir(src):
                    real_src, was_nested = _resolve_nested_source(src)
                    if was_nested:
                        log.append(
                            f"Info: Nested structure detected in '{label}' - "
                            f"using inner folder as source to prevent duplication."
                        )
                    try:
                        shutil.copytree(real_src, dst, dirs_exist_ok=True)
                        log.append(f"Status: '{label}' restored.")
                    except OSError as exc:
                        log.append(f"Warning: Could not restore '{label}': {exc}")
                else:
                    log.append(f"Warning: '{label}' not found in archive, skipping.")

            # Step 3 - Import registry key
            if portable:
                log.append("Info: Portable mode - registry import skipped.")
            else:
                reg_file = os.path.join(stage_dir, "Windhawk.reg")
                if os.path.isfile(reg_file):
                    try:
                        subprocess.run(
                            ["reg", "import", reg_file],
                            check=True, capture_output=True, text=True,
                            creationflags=subprocess.CREATE_NO_WINDOW,
                        )
                        log.append("Status: Registry imported.")
                    except subprocess.CalledProcessError as exc:
                        log.append(f"ERROR: Registry import failed: {exc.stderr.strip()}")
                        return False, "\n".join(log)
                else:
                    log.append("Warning: Registry file not found in archive, skipping.")

    except OSError as exc:
        log.append(f"ERROR: Staging directory error: {exc}")
        return False, "\n".join(log)

    finally:
        if not portable:
            _, msg = start_windhawk_service()
            log.append(msg)

    log.append("\nOperation Complete: Restore finished successfully.")
    return True, "\n".join(log)


# =============================================================================
#                       GRAPHICAL USER INTERFACE
# =============================================================================

class WindhawkManagerApp:
    """Main application window."""

    LOG_COLOURS: dict[str, str] = {
        "info":    "RoyalBlue",
        "success": "ForestGreen",
        "warning": "DarkOrange",
        "error":   "Crimson",
    }

    # Treeview column definitions: heading text, pixel width, anchor
    TV_COLUMNS: dict[str, tuple[str, int, str]] = {
        "date": ("Date / Time",  172, "w"),
        "size": ("Size",          68, "e"),
        "kind": ("Type",          74, "center"),
        "mods": ("Mods",          48, "center"),
        "name": ("Archive Name", 300, "w"),
    }

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("820x680")
        self.root.minsize(740, 580)

        self._apply_style()

        self._cfg = load_config()

        outer = ttk.Frame(root, padding=PAD)
        outer.pack(fill=tk.BOTH, expand=True)

        self._build_config_section(outer)
        self._build_archive_section(outer)
        self._build_log_section(outer)
        self._build_status_bar(root)

        self._configure_log_tags()
        self._apply_config()
        self._refresh_backup_list()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Styling
    # ------------------------------------------------------------------

    def _apply_style(self) -> None:
        s = ttk.Style()
        s.theme_use("vista")

        s.configure("Treeview", rowheight=23, font=("Segoe UI", 9))
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        s.map("Treeview",
              background=[("selected", "#CCE4F7")],
              foreground=[("selected", "#000000")])

        s.configure("Accent.Horizontal.TProgressbar",
                    troughcolor="#E4E4E4", background="#3A9BD5", thickness=5)

        s.configure("Status.TLabel",
                    font=("Segoe UI", 8), foreground="#555555", background="#F0F0F0")
        s.configure("StatusBar.TFrame", background="#F0F0F0", relief="sunken")

    # ------------------------------------------------------------------
    # UI builders
    # ------------------------------------------------------------------

    def _build_config_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Configuration", padding=PAD)
        frame.pack(fill=tk.X, pady=(0, PAD))
        frame.columnconfigure(1, weight=1)

        lbl = {"sticky": "w", "padx": (0, PAD), "pady": 4}
        ent = {"sticky": "ew", "pady": 4}

        # Windhawk root
        self.windhawk_path_var = tk.StringVar()
        ttk.Label(frame, text="Windhawk Root:").grid(row=0, column=0, **lbl)
        ttk.Entry(frame, textvariable=self.windhawk_path_var).grid(row=0, column=1, **ent)
        ttk.Button(frame, text="Browse...", width=10,
                   command=self._select_windhawk_path).grid(row=0, column=2, padx=(PAD, 0), pady=4)

        # Backup folder
        self.backup_path_var = tk.StringVar()
        ttk.Label(frame, text="Backup Folder:").grid(row=1, column=0, **lbl)
        ttk.Entry(frame, textvariable=self.backup_path_var).grid(row=1, column=1, **ent)
        ttk.Button(frame, text="Browse...", width=10,
                   command=self._select_backup_path).grid(row=1, column=2, padx=(PAD, 0), pady=4)

        # Options row
        opts = ttk.Frame(frame)
        opts.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 0))

        ttk.Label(opts, text="Keep last").pack(side=tk.LEFT)

        self.max_backups_var = tk.IntVar(value=DEFAULT_MAX_BACKUPS)
        ttk.Spinbox(opts, from_=0, to=99, width=4,
                    textvariable=self.max_backups_var).pack(side=tk.LEFT, padx=(4, 4))

        ttk.Label(opts, text="backups  (0 = unlimited)").pack(side=tk.LEFT, padx=(0, PAD * 2))

        self.portable_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Portable installation",
                         variable=self.portable_var,
                         command=self._on_portable_toggled).pack(side=tk.LEFT, padx=(0, PAD))

        ttk.Button(opts, text="Auto-Detect", width=11,
                   command=self._auto_detect_portable).pack(side=tk.LEFT)

    def _build_archive_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Backup Archives", padding=PAD)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, PAD))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # Treeview with scrollbar
        tv_frame = ttk.Frame(frame)
        tv_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        tv_frame.columnconfigure(0, weight=1)
        tv_frame.rowconfigure(0, weight=1)

        cols = list(self.TV_COLUMNS.keys())
        self.tree = ttk.Treeview(
            tv_frame, columns=cols, show="headings",
            selectmode="browse", height=8,
        )
        for col, (heading, width, anchor) in self.TV_COLUMNS.items():
            self.tree.heading(col, text=heading,
                              command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=width, minwidth=40, anchor=anchor)

        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd",  background="#F2F6FA")
        self.tree.bind("<Double-1>", lambda _e: self._show_preview())

        # Action buttons
        btn = ttk.Frame(frame)
        btn.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(PAD, 0))

        self.backup_button = ttk.Button(
            btn, text="Create Backup", width=15, command=self._run_backup)
        self.backup_button.pack(side=tk.LEFT)

        self.restore_button = ttk.Button(
            btn, text="Restore Selected", width=15, command=self._restore_selected)
        self.restore_button.pack(side=tk.LEFT, padx=(PAD, 0))

        self.delete_button = ttk.Button(
            btn, text="Delete Selected", width=15, command=self._delete_selected)
        self.delete_button.pack(side=tk.LEFT, padx=(PAD, 0))

        ttk.Button(btn, text="Refresh", width=9,
                   command=self._refresh_backup_list).pack(side=tk.RIGHT)

    def _build_log_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Operation Log", padding=PAD)
        frame.pack(fill=tk.X, pady=(0, PAD))
        frame.columnconfigure(0, weight=1)

        hdr = ttk.Frame(frame)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Button(hdr, text="Export Log...", width=12,
                   command=self._export_log).pack(side=tk.RIGHT)

        self.log_widget = scrolledtext.ScrolledText(
            frame, height=7, wrap=tk.WORD, state=tk.DISABLED,
            font=("Consolas", 9), relief="flat",
            background="#FAFAFA", borderwidth=1,
        )
        self.log_widget.grid(row=1, column=0, sticky="ew")

        self.progressbar = ttk.Progressbar(
            frame, mode="indeterminate", length=200,
            style="Accent.Horizontal.TProgressbar",
        )
        self.progressbar.grid(row=2, column=0, sticky="ew", pady=(6, 0))

    def _build_status_bar(self, parent: tk.Tk) -> None:
        bar = ttk.Frame(parent, style="StatusBar.TFrame", height=22)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bar, textvariable=self.status_var,
                  style="Status.TLabel").pack(side=tk.LEFT, padx=(PAD, 0), pady=2)

    # ------------------------------------------------------------------
    # Config helpers
    # ------------------------------------------------------------------

    def _apply_config(self) -> None:
        self.windhawk_path_var.set(self._cfg.get("windhawk_root", DEFAULT_WINDHAWK_ROOT))
        self.backup_path_var.set(self._cfg.get("backup_folder",  DEFAULT_BACKUP_FOLDER))
        self.portable_var.set(self._cfg.get("portable",          False))
        self.max_backups_var.set(self._cfg.get("max_backups",    DEFAULT_MAX_BACKUPS))
        self.log("Info: Configuration loaded.", "info")

    def _collect_config(self) -> dict:
        return {
            "windhawk_root": self.windhawk_path_var.get().strip(),
            "backup_folder": self.backup_path_var.get().strip(),
            "portable":      self.portable_var.get(),
            "max_backups":   self.max_backups_var.get(),
        }

    def _on_close(self) -> None:
        save_config(self._collect_config())
        self.root.destroy()

    # ------------------------------------------------------------------
    # Treeview helpers
    # ------------------------------------------------------------------

    def _refresh_backup_list(self) -> None:
        self.tree.delete(*self.tree.get_children())
        backups = list_backups(self.backup_path_var.get().strip())
        for i, b in enumerate(backups):
            self.tree.insert(
                "", tk.END, iid=b["path"],
                values=(b["date"], b["size"], b["kind"], b["mods"], b["name"]),
                tags=("even" if i % 2 == 0 else "odd",),
            )
        count = len(backups)
        self._set_status(
            f"{count} backup{'s' if count != 1 else ''} found."
            if count else "No backups found in the selected folder."
        )

    def _sort_tree(self, col: str) -> None:
        """Sorts all treeview rows by the clicked column, ascending."""
        items = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children()]
        items.sort(key=lambda x: x[0])
        for i, (_val, iid) in enumerate(items):
            self.tree.move(iid, "", i)
            self.tree.item(iid, tags=("even" if i % 2 == 0 else "odd",))

    def _selected_archive_path(self) -> str | None:
        sel = self.tree.selection()
        return sel[0] if sel else None

    # ------------------------------------------------------------------
    # Browse dialogs
    # ------------------------------------------------------------------

    def _safe_initial_dir(self, path: str) -> str:
        return path if os.path.isdir(path) else os.path.expanduser("~")

    def _select_windhawk_path(self) -> None:
        path = filedialog.askdirectory(
            title="Select Windhawk Installation Directory",
            initialdir=self._safe_initial_dir(self.windhawk_path_var.get()),
        )
        if path:
            self.windhawk_path_var.set(os.path.normpath(path))

    def _select_backup_path(self) -> None:
        path = filedialog.askdirectory(
            title="Select Backup Destination Folder",
            initialdir=self._safe_initial_dir(self.backup_path_var.get()),
        )
        if path:
            self.backup_path_var.set(os.path.normpath(path))
            self._refresh_backup_list()

    # ------------------------------------------------------------------
    # Portable / auto-detect
    # ------------------------------------------------------------------

    def _auto_detect_portable(self) -> None:
        if registry_key_exists(WINDHAWK_REGISTRY_KEY):
            self.portable_var.set(False)
            self.log("Info: Registry key found - standard installation detected. Portable mode disabled.", "info")
        else:
            self.portable_var.set(True)
            self.log("Info: Registry key not found - portable installation assumed. Portable mode enabled.", "warning")

    def _on_portable_toggled(self) -> None:
        if self.portable_var.get():
            self.log("Info: Portable mode enabled - registry steps will be skipped.", "warning")
        else:
            self.log("Info: Portable mode disabled - registry steps will be included.", "info")

    # ------------------------------------------------------------------
    # Logging and status
    # ------------------------------------------------------------------

    def _configure_log_tags(self) -> None:
        """Configures log colour tags once at startup."""
        for tag, colour in self.LOG_COLOURS.items():
            self.log_widget.tag_config(tag, foreground=colour)

    def log(self, message: str, level: str = "info") -> None:
        """Appends a timestamped message to the log widget (thread-safe)."""
        ts   = datetime.datetime.now().strftime("%H:%M:%S")
        text = f"[{ts}]  {message}"

        def _write() -> None:
            self.log_widget.config(state=tk.NORMAL)
            self.log_widget.insert(tk.END, text + "\n", (level,))
            self.log_widget.see(tk.END)
            self.log_widget.config(state=tk.DISABLED)

        self.root.after(0, _write)

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))

    def _export_log(self) -> None:
        default_name = f"wsbu-log-{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        path = filedialog.asksaveasfilename(
            title="Export Operation Log",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            initialfile=default_name,
        )
        if not path:
            return
        content = self.log_widget.get("1.0", tk.END)
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(content)
            self.log(f"Info: Log exported to: {path}", "info")
        except OSError as exc:
            messagebox.showerror("Export Failed", str(exc))

    # ------------------------------------------------------------------
    # Backup preview
    # ------------------------------------------------------------------

    def _show_preview(self) -> None:
        """
        Opens a detail window for the selected backup archive.
        Reads manifest.json from inside the ZIP without extracting it.
        """
        archive = self._selected_archive_path()
        if not archive:
            return

        # --- Read manifest ---
        manifest: dict = {}
        try:
            with zipfile.ZipFile(archive, "r") as zf:
                if "manifest.json" in zf.namelist():
                    manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
        except Exception as exc:
            messagebox.showerror("Preview Failed", f"Could not read archive:\n{exc}")
            return

        # --- Build dialog ---
        win = tk.Toplevel(self.root)
        win.title(f"Backup Details  -  {os.path.basename(archive)}")
        win.geometry("480x540")
        win.minsize(400, 460)
        win.resizable(True, True)
        win.grab_set()

        outer = ttk.Frame(win, padding=PAD)
        outer.pack(fill=tk.BOTH, expand=True)

        def _row(parent: ttk.Frame, label: str, value: str, row: int) -> None:
            ttk.Label(parent, text=label, font=("Segoe UI", 9, "bold"),
                      anchor="w").grid(row=row, column=0, sticky="w",
                                       padx=(0, PAD), pady=3)
            ttk.Label(parent, text=value, anchor="w").grid(
                row=row, column=1, sticky="ew", pady=3)

        # --- Metadata grid ---
        meta = ttk.LabelFrame(outer, text="Archive Information", padding=PAD)
        meta.pack(fill=tk.X, pady=(0, PAD))
        meta.columnconfigure(1, weight=1)

        size_bytes = os.path.getsize(archive)

        if manifest:
            _row(meta, "Created:",         manifest.get("created",     "-"), 0)
            _row(meta, "Utility Version:", manifest.get("app_version", "-"), 1)
            _row(meta, "Installation:",    "Portable" if manifest.get("portable") else "Standard", 2)
            _row(meta, "Mod Count:",       str(manifest.get("mod_count", "-")), 3)
            _row(meta, "Archive Size:",    _format_size(size_bytes), 4)
        else:
            _row(meta, "Archive:",  os.path.basename(archive), 0)
            _row(meta, "Size:",     _format_size(size_bytes),  1)
            ttk.Label(meta, text="No manifest.json found in this archive (legacy backup).",
                      foreground="DarkOrange").grid(
                row=2, column=0, columnspan=2, sticky="w", pady=(PAD, 0))

        # --- Mod list ---
        mods: list[str] = manifest.get("mods", [])
        mod_frame = ttk.LabelFrame(
            outer,
            text=f"Installed Mods  ({len(mods)})" if mods else "Installed Mods",
            padding=PAD,
        )
        mod_frame.pack(fill=tk.BOTH, expand=True, pady=(0, PAD))
        mod_frame.rowconfigure(0, weight=1)
        mod_frame.columnconfigure(0, weight=1)

        if mods:
            lb_frame = ttk.Frame(mod_frame)
            lb_frame.grid(sticky="nsew")
            lb_frame.rowconfigure(0, weight=1)
            lb_frame.columnconfigure(0, weight=1)

            lb = tk.Listbox(lb_frame, font=("Consolas", 9),
                            selectmode=tk.BROWSE,
                            relief="flat", borderwidth=0,
                            background="#FAFAFA", activestyle="none",
                            highlightthickness=1, highlightcolor="#CCE4F7",
                            highlightbackground="#DDDDDD")
            sb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=lb.yview)
            lb.configure(yscrollcommand=sb.set)
            lb.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")

            for i, mod in enumerate(sorted(mods)):
                display = mod[:-7] if mod.endswith(".wh.cpp") else mod
                lb.insert(tk.END, f"  {display}")
                lb.itemconfig(i, background="#FFFFFF" if i % 2 == 0 else "#F2F6FA")
        else:
            ttk.Label(mod_frame,
                      text="No mod list available (legacy backup).",
                      foreground="DarkOrange").grid(sticky="w")

        # --- Close button ---
        ttk.Button(outer, text="Close", width=10,
                   command=win.destroy).pack(anchor="e")

    # ------------------------------------------------------------------
    # Operation control
    # ------------------------------------------------------------------

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        for btn in (self.backup_button, self.restore_button, self.delete_button):
            btn.config(state=state)
        if enabled:
            self.progressbar.stop()
            self.progressbar.config(value=0)
        else:
            self.progressbar.start(12)

    # ------------------------------------------------------------------
    # Backup
    # ------------------------------------------------------------------

    def _run_backup(self) -> None:
        cfg = self._collect_config()
        if not cfg["windhawk_root"] or not cfg["backup_folder"]:
            messagebox.showwarning(
                "Configuration Incomplete",
                "Please specify both the Windhawk root and the backup folder.",
            )
            return

        self.log("\n--- Backup started ---", "info")
        self._set_status("Backup in progress...")
        self._set_controls_enabled(False)

        def _worker() -> None:
            success, message = execute_backup_operation(
                cfg["windhawk_root"],
                cfg["backup_folder"],
                portable=cfg["portable"],
                max_backups=cfg["max_backups"],
            )
            self.root.after(0, lambda: self._on_backup_done(success, message))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_backup_done(self, success: bool, message: str) -> None:
        self._set_controls_enabled(True)
        self.log(message, "success" if success else "error")
        self._set_status("Backup completed." if success else "Backup failed - see log.")
        self._refresh_backup_list()
        if success:
            messagebox.showinfo("Backup Succeeded", "The backup completed successfully.")
        else:
            messagebox.showerror("Backup Failed", "An error occurred. Please review the log.")

    # ------------------------------------------------------------------
    # Restore
    # ------------------------------------------------------------------

    def _restore_selected(self) -> None:
        archive = self._selected_archive_path()
        if not archive:
            messagebox.showinfo("No Selection",
                                "Please select a backup from the list to restore.")
            return

        wh_path = self.windhawk_path_var.get().strip()
        if not wh_path:
            messagebox.showwarning("Configuration Incomplete",
                                   "Please specify the Windhawk root path.")
            return

        name = os.path.basename(archive)
        if not messagebox.askyesno(
            "Confirm Restore",
            f"Restore from:\n{name}\n\nThis will overwrite existing mod files. Continue?",
        ):
            return

        self.log(f"\n--- Restore started: {name} ---", "info")
        self._set_status("Restore in progress...")
        self._set_controls_enabled(False)

        portable = self.portable_var.get()

        def _worker() -> None:
            success, message = execute_restore_operation(wh_path, archive, portable)
            self.root.after(0, lambda: self._on_restore_done(success, message))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_restore_done(self, success: bool, message: str) -> None:
        self._set_controls_enabled(True)
        self.log(message, "success" if success else "error")
        self._set_status("Restore completed." if success else "Restore failed - see log.")
        if success:
            messagebox.showinfo("Restore Succeeded", "The restore completed successfully.")
        else:
            messagebox.showerror("Restore Failed", "An error occurred. Please review the log.")

    # ------------------------------------------------------------------
    # Delete
    # ------------------------------------------------------------------

    def _delete_selected(self) -> None:
        archive = self._selected_archive_path()
        if not archive:
            messagebox.showinfo("No Selection",
                                "Please select a backup from the list to delete.")
            return

        name = os.path.basename(archive)
        if not messagebox.askyesno(
            "Confirm Delete",
            f"Permanently delete:\n{name}\n\nThis cannot be undone.",
            icon="warning",
        ):
            return

        try:
            os.remove(archive)
            self.log(f"Info: Deleted backup: {name}", "info")
        except OSError as exc:
            messagebox.showerror("Delete Failed", str(exc))
            return

        self._refresh_backup_list()
        self._set_status(f"Deleted: {name}")


# =============================================================================
#                          APPLICATION ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    if not is_admin():
        if not run_as_admin():
            messagebox.showerror(
                "Elevation Failed",
                "This application requires administrator privileges and could not elevate.\n"
                "Please re-run it manually as Administrator.",
            )
        sys.exit()

    root = tk.Tk()
    WindhawkManagerApp(root)
    root.mainloop()
