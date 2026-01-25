# =============================================================================
#  Windhawk Service Management Utility - Version 2.1.1
#  Author: scorpion421
#  Description: A tool for backing up and restoring Windhawk configurations.
# =============================================================================

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import shutil
import subprocess
import datetime
import tempfile
import ctypes
import sys

# --- APPLICATION CONSTANTS ---
# Define paths with environment variables as templates
_BACKUP_FOLDER_TEMPLATE = r"%userprofile%\Documents\Windhawk_Backup"
_WINDHAWK_ROOT_TEMPLATE = r"%programdata%\Windhawk"

# Expand environment variables to get the actual, usable paths
DEFAULT_BACKUP_FOLDER = os.path.expandvars(_BACKUP_FOLDER_TEMPLATE)
DEFAULT_WINDHAWK_ROOT = os.path.expandvars(_WINDHAWK_ROOT_TEMPLATE)
WINDHAWK_REGISTRY_KEY = r"SOFTWARE\Windhawk"

# =============================================================================
#                            CORE LOGIC (BACKEND)
# =============================================================================

def is_admin():
    """Checks if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Re-launches the script with elevated (administrator) privileges."""
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

def execute_backup_operation(windhawk_root_path, backup_destination_folder):
    """Executes the complete backup process."""
    if not os.path.exists(backup_destination_folder):
        os.makedirs(backup_destination_folder)
    
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_zip_path = os.path.join(backup_destination_folder, f"windhawk-backup_{timestamp}")

    with tempfile.TemporaryDirectory() as temp_dir:
        log_messages = []
        
        # Step 1: Copy source directories
        mods_source_folder = os.path.join(windhawk_root_path, "ModsSource")
        engine_mods_folder = os.path.join(windhawk_root_path, "Engine", "Mods")
        
        if os.path.exists(mods_source_folder):
            shutil.copytree(mods_source_folder, os.path.join(temp_dir, "ModsSource"))
            log_messages.append("Status: ModsSource directory successfully staged for backup.")
        else:
            log_messages.append(f"Warning: ModsSource directory not found at: {mods_source_folder}")

        if os.path.exists(engine_mods_folder):
            shutil.copytree(engine_mods_folder, os.path.join(temp_dir, "Engine", "Mods"))
            log_messages.append("Status: Engine\\Mods directory successfully staged for backup.")
        else:
            log_messages.append(f"Warning: Engine\\Mods directory not found at: {engine_mods_folder}")
            
        # Step 2: Export registry key
        reg_export_file = os.path.join(temp_dir, "Windhawk.reg")
        try:
            # Construct full registry path for the command line tool
            full_reg_key = f"HKLM\\{WINDHAWK_REGISTRY_KEY}"
            subprocess.run(
                ['reg', 'export', full_reg_key, reg_export_file, '/y'],
                check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            log_messages.append("Status: Windhawk registry key successfully exported.")
        except subprocess.CalledProcessError as e:
            log_messages.append(f"ERROR: Registry export failed. Details: {e.stderr}")
            return False, "\n".join(log_messages)

        # Step 3: Compress the staged files into a single archive
        shutil.make_archive(backup_zip_path, 'zip', temp_dir)
        log_messages.append(f"\nOperation Complete: Backup archive created at:\n{backup_zip_path}.zip")
        
        return True, "\n".join(log_messages)

def execute_restore_operation(windhawk_root_path, backup_zip_path):
    """Executes the complete restore process."""
    with tempfile.TemporaryDirectory() as temp_dir:
        log_messages = []

        # Step 1: Extract the backup archive
        try:
            shutil.unpack_archive(backup_zip_path, temp_dir)
            log_messages.append(f"Status: Archive '{os.path.basename(backup_zip_path)}' extracted successfully.")
        except Exception as e:
            log_messages.append(f"ERROR: Failed to extract archive. Details: {e}")
            return False, "\n".join(log_messages)

        # Step 2: Copy directories to target location
        backup_mods_source = os.path.join(temp_dir, "ModsSource")
        backup_engine_mods = os.path.join(temp_dir, "Engine", "Mods")
        
        if os.path.exists(backup_mods_source):
            shutil.copytree(backup_mods_source, os.path.join(windhawk_root_path, "ModsSource"), dirs_exist_ok=True)
            log_messages.append("Status: ModsSource directory restored.")
        else:
            log_messages.append("Warning: ModsSource directory not found in backup archive.")

        if os.path.exists(backup_engine_mods):
            shutil.copytree(backup_engine_mods, os.path.join(windhawk_root_path, "Engine", "Mods"), dirs_exist_ok=True)
            log_messages.append("Status: Engine\\Mods directory restored.")
        else:
            log_messages.append("Warning: Engine\\Mods directory not found in backup archive.")
            
        # Step 3: Import registry settings
        reg_backup_file = os.path.join(temp_dir, "Windhawk.reg")
        if os.path.exists(reg_backup_file):
            try:
                subprocess.run(
                    ['reg', 'import', reg_backup_file],
                    check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
                )
                log_messages.append("Status: Windhawk registry settings imported.")
            except subprocess.CalledProcessError as e:
                log_messages.append(f"ERROR: Registry import failed. Details: {e.stderr}")
                return False, "\n".join(log_messages)
        else:
            log_messages.append("Warning: Registry file (Windhawk.reg) not found in backup archive.")
            
        log_messages.append("\nOperation Complete: Restore process finished successfully.")
        return True, "\n".join(log_messages)

# =============================================================================
#                       GRAPHICAL USER INTERFACE (FRONTEND)
# =============================================================================

class WindhawkManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windhawk Service Management Utility v2.1.1")
        self.root.geometry("700x500")
        self.root.minsize(600, 450)
        
        self.style = ttk.Style()
        self.style.theme_use('vista')

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Configuration Section ---
        config_frame = ttk.LabelFrame(main_frame, text="Configuration Paths", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)

        # Windhawk Path
        self.windhawk_path_var = tk.StringVar()
        ttk.Label(config_frame, text="Windhawk Root:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        windhawk_entry = ttk.Entry(config_frame, textvariable=self.windhawk_path_var)
        windhawk_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        ttk.Button(config_frame, text="Browse...", command=self.select_windhawk_path).grid(row=0, column=2, padx=5, pady=5)
        
        # Backup Path
        self.backup_path_var = tk.StringVar()
        ttk.Label(config_frame, text="Backup Destination:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        backup_entry = ttk.Entry(config_frame, textvariable=self.backup_path_var)
        backup_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        ttk.Button(config_frame, text="Browse...", command=self.select_backup_path).grid(row=1, column=2, padx=5, pady=5)

        # --- Operations Section ---
        ops_frame = ttk.Frame(main_frame)
        ops_frame.pack(fill=tk.X, pady=10)
        
        self.backup_button = ttk.Button(ops_frame, text="Execute Backup Operation", command=self.run_backup)
        self.backup_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X, ipady=5)

        self.restore_button = ttk.Button(ops_frame, text="Execute Restore Operation", command=self.run_restore)
        self.restore_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X, ipady=5)

        # --- Logging Section ---
        log_label = ttk.Label(main_frame, text="Operation Log:")
        log_label.pack(anchor=tk.W, pady=(10, 2))
        
        self.log_widget = scrolledtext.ScrolledText(main_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 9))
        self.log_widget.pack(fill=tk.BOTH, expand=True)
        
        self.initialize_paths()

    def initialize_paths(self):
        """Sets the user-defined default paths for the configuration entries."""
        self.windhawk_path_var.set(DEFAULT_WINDHAWK_ROOT)
        self.backup_path_var.set(DEFAULT_BACKUP_FOLDER)
        self.log("Info: Default configuration paths have been loaded. Please verify they are correct.", "info")

    def select_windhawk_path(self):
        path = filedialog.askdirectory(title="Select Windhawk Installation Directory", initialdir=self.windhawk_path_var.get())
        if path:
            self.windhawk_path_var.set(os.path.normpath(path))

    def select_backup_path(self):
        path = filedialog.askdirectory(title="Select Backup Destination Folder", initialdir=self.backup_path_var.get())
        if path:
            self.backup_path_var.set(os.path.normpath(path))
            
    def log(self, message, level="info"):
        """Writes a message to the log widget with appropriate color coding."""
        self.log_widget.config(state=tk.NORMAL)
        color_map = {"info": "blue", "error": "red", "success": "green", "warning": "darkorange"}
        tag_name = level
        self.log_widget.insert(tk.END, message + "\n", (tag_name,))
        self.log_widget.tag_config(tag_name, foreground=color_map.get(level, "black"))
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def run_backup(self):
        wh_path = self.windhawk_path_var.get()
        bk_path = self.backup_path_var.get()

        if not wh_path or not bk_path:
            messagebox.showwarning("Configuration Incomplete", "Please ensure both Windhawk root and backup destination paths are specified.")
            return

        self.log("\n--- Initializing Backup Operation... ---", "info")
        success, message = execute_backup_operation(wh_path, bk_path)
        if success:
            self.log(message, "success")
            messagebox.showinfo("Operation Succeeded", "The Windhawk backup operation completed successfully.")
        else:
            self.log(message, "error")
            messagebox.showerror("Operation Failed", "An error occurred during the backup operation. Please review the log for details.")

    def run_restore(self):
        wh_path = self.windhawk_path_var.get()
        bk_path = self.backup_path_var.get()

        if not wh_path:
             messagebox.showwarning("Configuration Incomplete", "Please specify the Windhawk root path before restoring.")
             return
        
        if not os.path.exists(bk_path):
            os.makedirs(bk_path)

        backup_file = filedialog.askopenfilename(
            title="Select a Backup Archive to Restore",
            initialdir=bk_path,
            filetypes=[("Zip Archives", "*.zip")]
        )
        
        if backup_file:
            self.log(f"\n--- Initializing Restore Operation from '{os.path.basename(backup_file)}'... ---", "info")
            success, message = execute_restore_operation(wh_path, backup_file)
            if success:
                self.log(message, "success")
                messagebox.showinfo("Operation Succeeded", "The Windhawk restore operation completed successfully.")
            else:
                self.log(message, "error")
                messagebox.showerror("Operation Failed", "An error occurred during the restore operation. Please review the log for details.")

# =============================================================================
#                          APPLICATION ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    if not is_admin():
        # Re-launch with admin rights if not already elevated
        run_as_admin()
        sys.exit()

    # If running as admin, proceed to launch the GUI
    root = tk.Tk()
    app = WindhawkManagerApp(root)
    root.mainloop()
