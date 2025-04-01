import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkFont
import os
import subprocess
import sys
import ctypes  # For admin check and elevation
import json
from datetime import datetime
import time

# --- Constants ---
APP_NAME = "LinkForge"
APP_AUTHOR = "JunctionApp" # For AppData path
WIN_WIDTH = 750
WIN_HEIGHT = 550
PAD_GENERAL = 15
PAD_SMALL = 7

# Font Sizes
FONT_SIZE_BASE = 11
FONT_SIZE_LARGE = 12
FONT_SIZE_MONO = 10

# Colors
COLOR_SUCCESS = "#4CAF50"
COLOR_ERROR = "#F44336"
COLOR_INFO = "#2196F3"
COLOR_WARN = "#FF9800"
COLOR_TOOLTIP_BG = "#FFFFE0"
COLOR_TOOLTIP_FG = "#000000"
COLOR_DISABLED_FG = "#9E9E9E"
COLOR_VALID = "#2E7D32"
COLOR_INVALID = "#C62828"

# Icons
ICON_INFO = "ℹ️"
ICON_BROWSE = "📁"
ICON_COPY = "📋"
ICON_HISTORY = "🕒"
ICON_VALID = "✅"
ICON_INVALID = "❌"
ICON_DELETE = "🗑️" # Keep for potential future use? Or remove if definitely not needed.
ICON_EDIT = "✏️"

# History File
HISTORY_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Local", APP_AUTHOR, APP_NAME)
HISTORY_FILE = os.path.join(HISTORY_DIR, "history.json")

# --- Tooltip Texts ---
TOOLTIP_SOURCE = "The EXISTING directory that the link will point TO."
TOOLTIP_PARENT = "The directory WHERE the new link folder will be CREATED."
TOOLTIP_NAME = "The NAME of the new link folder to be created."
TOOLTIP_COPY = "Copy the 'mklink' command below to the clipboard."
TOOLTIP_HISTORY = "View, manage, and validate previously created junctions."
TOOLTIP_CREATE_DISABLED = "Run as Administrator to enable creating links."
TOOLTIP_CREATE_ENABLED = "Click to create the junction link."
TOOLTIP_EDIT = "Load selected entry into main window for editing."
TOOLTIP_VIEW_FOLDER = "Open the selected link's location or its source target in File Explorer." # New
# Removed delete tooltips

# --- Helper Functions ---
# is_admin, relaunch_as_admin, ensure_dir_exists, load_history, save_history, check_junction_validity
# remain the same as the previous correct version.
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def relaunch_as_admin():
    try:
        script_path = os.path.abspath(sys.argv[0])
        params = f'"{script_path}"'
        ret = ctypes.windll.shell32.ShellExecuteW( None, "runas", sys.executable, params, None, 1)
        return ret > 32
    except Exception as e:
        print(f"Error attempting elevation: {e}")
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Elevation Error", f"Could not relaunch as admin:\n{e}", parent=None)
        root.destroy()
        return False

def ensure_dir_exists(dir_path):
    if not os.path.exists(dir_path):
        try: os.makedirs(dir_path)
        except OSError as e: print(f"Error creating dir {dir_path}: {e}"); return False
    return True

def load_history():
    if not os.path.exists(HISTORY_FILE): return []
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f: history = json.load(f)
        return history if isinstance(history, list) else []
    except Exception as e: print(f"Error loading history: {e}"); return []

def save_history(history_list):
    if not ensure_dir_exists(HISTORY_DIR): return False
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f: json.dump(history_list, f, indent=4)
        return True
    except Exception as e: print(f"Error saving history: {e}"); return False

def check_junction_validity(link_path, source_path):
    try:
        if not os.path.lexists(link_path): return (1, "Link Missing")
        is_link_type = False
        try:
            link_stat = os.lstat(link_path)
            if hasattr(link_stat, 'st_file_attributes'):
                FILE_ATTRIBUTE_REPARSE_POINT = 0x0400
                if link_stat.st_file_attributes & FILE_ATTRIBUTE_REPARSE_POINT: is_link_type = True
            elif os.path.islink(link_path): is_link_type = True
        except OSError: pass
        except AttributeError:
             if os.path.islink(link_path): is_link_type = True
        if not is_link_type: return (3, "Exists, Not Link/Junction")
        if not os.path.isdir(source_path): return (2, "Source Missing/Invalid")
        return (0, "Valid")
    except Exception as e:
        print(f"Error validating junction {link_path}: {e}")
        return (4, f"Validation Error ({type(e).__name__})")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # _MEIPASS not defined, running in normal Python environment
        base_path = os.path.abspath(".") # Or os.path.dirname(__file__)

    return os.path.join(base_path, relative_path)

# --- Main Application Class ---
class JunctionApp(tk.Tk):
    # __init__, setup_window, setup_styles, _configure_widget_styles, setup_variables,
    # _create_widgets (main part), _create_input_row, _add_info_icon, setup_bindings,
    # create_tooltip, _show_tooltip, _hide_tooltip, _browse_source, _browse_link_parent,
    # _copy_command, _update_status, _update_command_preview, _create_junction,
    # populate_fields_from_history, _open_history_window, on_closing
    # remain the same as the previous correct version, EXCEPT for _check_admin_status
    def __init__(self, running_as_admin):
        super().__init__()
        self.running_as_admin = running_as_admin
        print(f"JunctionApp initialized with running_as_admin = {self.running_as_admin}") # DEBUG
        self.history_data = load_history()
        self.history_window = None
        self.tooltip_window = None

        self.setup_window()
        self.setup_styles()
        self.setup_variables()
        self._create_widgets()
        self.setup_bindings()
        self._configure_widget_styles()

        initial_status = f"Ready. {ICON_INFO} for help."
        initial_color = COLOR_INFO
        if not self.running_as_admin:
             initial_status = f"Running without admin privileges. Create disabled."
             initial_color = COLOR_WARN
        self._update_status(initial_status, initial_color)
        self._check_admin_status()
        self._update_command_preview()

    def setup_window(self):
        self.title(f"{APP_NAME}{' (Admin)' if self.running_as_admin else ''}")
        try:
            # Use the helper function to find the icon
            icon_path = resource_path("icon.ico")
            self.iconbitmap(icon_path) # Use self.iconbitmap() for window icon
            print(f"DEBUG: Attempting to load icon from: {icon_path}") # Optional debug print
        except tk.TclError as e:
            print(f"Warning: Could not load window icon 'icon.ico': {e}", file=sys.stderr)
        except FileNotFoundError:
             print(f"Warning: Icon file 'icon.ico' not found at expected path: {icon_path}", file=sys.stderr)
        self.center_window(WIN_WIDTH, WIN_HEIGHT)
        self.minsize(WIN_WIDTH - 200, WIN_HEIGHT - 200)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_coord = int((screen_width / 2) - (width / 2))
        y_coord = int((screen_height / 2) - (height / 2))
        self.geometry(f"{width}x{height}+{x_coord}+{y_coord}")

    def setup_styles(self):
        self.style = ttk.Style()
        try:
             available = self.style.theme_names()
             if 'vista' in available: self.style.theme_use('vista')
             elif 'clam' in available: self.style.theme_use('clam')
             elif 'xpnative' in available: self.style.theme_use('xpnative')
             print(f"Using theme: {self.style.theme_use()}")
        except tk.TclError: print("Error setting default theme.")
        try: self.configure(bg=self.style.lookup('TFrame', 'background'))
        except tk.TclError: self.configure(bg='SystemButtonFace')

    def _configure_widget_styles(self):
        base_font_family = 'Segoe UI'
        mono_font_family = "Consolas" if "Consolas" in tkFont.families(self) else ("Courier New" if "Courier New" in tkFont.families(self) else "Courier")
        self.base_font = tkFont.Font(family=base_font_family, size=FONT_SIZE_BASE)
        self.large_font = tkFont.Font(family=base_font_family, size=FONT_SIZE_LARGE, weight='bold')
        self.mono_font = tkFont.Font(family=mono_font_family, size=FONT_SIZE_MONO)
        self.icon_font = tkFont.Font(family=base_font_family, size=FONT_SIZE_BASE + 2)
        self.tree_heading_font = tkFont.Font(family=base_font_family, size=FONT_SIZE_BASE, weight='bold')
        self.tooltip_font = tkFont.Font(family=base_font_family, size=FONT_SIZE_BASE - 1)

        self.style.configure('.', font=self.base_font)
        self.style.configure('TLabel', padding=PAD_SMALL)
        self.style.configure('TButton', padding=(PAD_SMALL * 1.5, PAD_SMALL), font=self.base_font)
        self.style.configure('Toolbutton.TButton', font=self.icon_font, padding=(PAD_SMALL // 2))
        self.style.configure('Accent.TButton', font=self.large_font, padding=(PAD_GENERAL, PAD_SMALL * 1.5))
        self.style.configure("InfoIcon.TLabel", foreground=COLOR_INFO, font=self.icon_font)
        self.style.configure('TLabelframe.Label', font=self.base_font, padding=(0, PAD_SMALL // 2))
        self.style.configure('Treeview', rowheight=int(FONT_SIZE_BASE * 2.2))
        self.style.configure('Treeview.Heading', font=self.tree_heading_font)

        self.style.map("Accent.TButton", foreground=[('active', 'white'), ('!disabled', self.style.lookup('TButton', 'foreground'))], background=[('active', COLOR_INFO), ('!disabled', self.style.lookup('TButton', 'background'))], relief=[('pressed', tk.SUNKEN), ('!pressed', tk.RAISED)])
        try: theme_active_bg = self.style.map('TButton', 'background')[1][1]; hover_color = theme_active_bg if theme_active_bg else "#e0e0e0"
        except Exception: hover_color = "#cccccc"
        self.style.map("TButton", background=[('active', hover_color), ('!disabled', self.style.lookup('TButton', 'background'))], relief=[('pressed', tk.SUNKEN), ('!pressed', tk.RAISED)])
        self.style.map("Toolbutton.TButton", background=[('active', hover_color), ('!disabled', self.style.lookup('Toolbutton.TButton', 'background'))], relief=[('pressed', tk.SUNKEN), ('!pressed', tk.RAISED)])

        try: label_bg = self.style.lookup('TLabel', 'background'); default_fg = self.style.lookup('TLabel', 'foreground')
        except tk.TclError: label_bg = 'SystemButtonFace'; default_fg = 'black'
        status_padding = (PAD_SMALL, PAD_SMALL // 2)
        self.style.configure("Status.TLabel", padding=status_padding, anchor=tk.W, font=self.base_font)
        self.style.configure("Success.Status.TLabel", foreground=COLOR_SUCCESS, background=label_bg)
        self.style.configure("Error.Status.TLabel", foreground=COLOR_ERROR, background=label_bg)
        self.style.configure("Info.Status.TLabel", foreground=COLOR_INFO, background=label_bg)
        self.style.configure("Warn.Status.TLabel", foreground=COLOR_WARN, background=label_bg)
        self.style.configure("Default.Status.TLabel", foreground=default_fg, background=label_bg)

        self.style.configure("Valid.Treeview", foreground=COLOR_VALID)
        self.style.configure("Invalid.Treeview", foreground=COLOR_INVALID)
        self.style.configure("Error.Treeview", foreground=COLOR_WARN)

        if hasattr(self, 'command_preview_text'): self.command_preview_text.config(font=self.mono_font)

    def setup_variables(self):
        self.source_dir_var = tk.StringVar()
        self.link_parent_dir_var = tk.StringVar()
        self.link_name_var = tk.StringVar()
        self.status_var = tk.StringVar()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=PAD_GENERAL); main_frame.pack(expand=True, fill=tk.BOTH); main_frame.columnconfigure(2, weight=1)
        row_index = 0
        self.source_entry = self._create_input_row(main_frame, row_index, "Source Directory (Target):", self.source_dir_var, self._browse_source, TOOLTIP_SOURCE); row_index += 1
        self.link_parent_entry = self._create_input_row(main_frame, row_index, "Link Parent Directory:", self.link_parent_dir_var, self._browse_link_parent, TOOLTIP_PARENT); row_index += 1
        ttk.Label(main_frame, text="Link Name (New Folder):").grid(row=row_index, column=0, sticky=tk.W, padx=(0, PAD_SMALL), pady=PAD_SMALL)
        self._add_info_icon(main_frame, row=row_index, column=1, text=TOOLTIP_NAME)
        self.link_name_entry = ttk.Entry(main_frame, textvariable=self.link_name_var, width=45); self.link_name_entry.grid(row=row_index, column=2, sticky=tk.W, padx=PAD_SMALL, pady=PAD_SMALL); row_index += 1
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(row=row_index, column=0, columnspan=4, sticky=tk.EW, pady=PAD_GENERAL); row_index += 1
        preview_frame = ttk.LabelFrame(main_frame, text="Command Preview", padding=PAD_SMALL); preview_frame.grid(row=row_index, column=0, columnspan=4, sticky=tk.EW, padx=0, pady=(PAD_SMALL, PAD_GENERAL)); preview_frame.columnconfigure(0, weight=1)
        preview_font = self.mono_font if hasattr(self, 'mono_font') else ('Consolas', FONT_SIZE_MONO)
        self.command_preview_text = tk.Text(preview_frame, height=2, wrap=tk.WORD, font=preview_font, relief=tk.FLAT, state=tk.DISABLED, padx=PAD_SMALL, pady=PAD_SMALL)
        try: lf_bg = self.style.lookup('TLabelframe', 'background'); self.command_preview_text.config(background=lf_bg)
        except tk.TclError: pass
        self.command_preview_text.grid(row=0, column=0, sticky=tk.EW, padx=(PAD_SMALL, 0), pady=PAD_SMALL)
        self.copy_btn = ttk.Button(preview_frame, text=ICON_COPY, command=self._copy_command, style='Toolbutton.TButton', width=3); self.copy_btn.grid(row=0, column=1, sticky=tk.NE, padx=PAD_SMALL, pady=PAD_SMALL); self.create_tooltip(self.copy_btn, TOOLTIP_COPY); row_index += 1
        action_frame = ttk.Frame(main_frame); action_frame.grid(row=row_index, column=0, columnspan=4, pady=(PAD_GENERAL, PAD_GENERAL * 1.5)); action_frame.columnconfigure(0, weight=1); action_frame.columnconfigure(2, weight=1)
        self.create_button = ttk.Button(action_frame, text="Create Junction Link", command=self._create_junction, style="Accent.TButton"); self.create_button.grid(row=0, column=1, padx=PAD_SMALL)
        history_button_text = f"View History {ICON_HISTORY} ({len(self.history_data)})"
        self.history_button = ttk.Button(action_frame, text=history_button_text, command=self._open_history_window); self.history_button.grid(row=0, column=2, sticky=tk.E, padx=(PAD_GENERAL, 0)); self.create_tooltip(self.history_button, TOOLTIP_HISTORY); row_index += 1
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, style="Default.Status.TLabel"); self.status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(PAD_SMALL, 0))

    def _create_input_row(self, parent, row, label_text, var, browse_cmd, tooltip_text):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, padx=(0, PAD_SMALL), pady=PAD_SMALL)
        self._add_info_icon(parent, row=row, column=1, text=tooltip_text)
        entry = ttk.Entry(parent, textvariable=var, width=60); entry.grid(row=row, column=2, sticky=tk.EW, padx=PAD_SMALL, pady=PAD_SMALL)
        browse_btn = ttk.Button(parent, text=ICON_BROWSE, command=browse_cmd, style='Toolbutton.TButton', width=3); browse_btn.grid(row=row, column=3, sticky=tk.E, padx=(PAD_SMALL, 0), pady=PAD_SMALL)
        self.create_tooltip(browse_btn, f"Browse for '{label_text.split(':')[0]}'")
        return entry

    def _add_info_icon(self, parent, row, column, text):
        info_label = ttk.Label(parent, text=ICON_INFO, style="InfoIcon.TLabel", cursor="question_arrow"); info_label.grid(row=row, column=column, sticky=tk.W, padx=(0, PAD_SMALL))
        info_label.bind("<Enter>", lambda event, t=text: self._show_tooltip(event, t)); info_label.bind("<Leave>", self._hide_tooltip); info_label.bind("<Button-1>", lambda event, t=text: self._show_tooltip(event, t, sticky=True))

    def setup_bindings(self):
        self.source_dir_var.trace_add("write", self._update_command_preview); self.link_parent_dir_var.trace_add("write", self._update_command_preview); self.link_name_var.trace_add("write", self._update_command_preview)
        self.bind_all("<Escape>", self._hide_tooltip)

    def create_tooltip(self, widget, text):
         widget.bind("<Enter>", lambda event, t=text: self._show_tooltip(event, t)); widget.bind("<Leave>", self._hide_tooltip); widget.bind("<Button-1>", lambda event, t=text: self._show_tooltip(event, t, sticky=True))

    def _show_tooltip(self, event, text, sticky=False):
        try:
            if hasattr(event.widget, 'cget') and event.widget.cget('state') == tk.DISABLED: return
        except tk.TclError: pass
        if self.tooltip_window and sticky: return
        self._hide_tooltip(); x = event.x_root + 20; y = event.y_root + 10
        self.tooltip_window = tk.Toplevel(self); self.tooltip_window.wm_overrideredirect(True); self.tooltip_window.wm_geometry(f"+{x}+{y}"); self.tooltip_window.wm_attributes("-topmost", True)
        tooltip_font = self.tooltip_font if hasattr(self, 'tooltip_font') else ('Segoe UI', FONT_SIZE_BASE - 1)
        label = tk.Label(self.tooltip_window, text=text, justify=tk.LEFT, background=COLOR_TOOLTIP_BG, foreground=COLOR_TOOLTIP_FG, relief=tk.SOLID, borderwidth=1, font=tooltip_font, wraplength=WIN_WIDTH * 0.5, padx=PAD_SMALL, pady=PAD_SMALL); label.pack(ipadx=1)

    def _hide_tooltip(self, event=None):
        if self.tooltip_window: self.tooltip_window.destroy(); self.tooltip_window = None

    def _browse_source(self):
        dir_path = filedialog.askdirectory(title="Select Source Directory", parent=self)
        if dir_path: self.source_dir_var.set(os.path.normpath(dir_path)); self._update_status(f"Selected source: {dir_path}", COLOR_INFO); self.source_entry.focus(); self.source_entry.xview_moveto(1.0)

    def _browse_link_parent(self):
        dir_path = filedialog.askdirectory(title="Select Parent Directory for Link", parent=self)
        if dir_path: self.link_parent_dir_var.set(os.path.normpath(dir_path)); self._update_status(f"Selected link parent: {dir_path}", COLOR_INFO); self.link_parent_entry.focus(); self.link_parent_entry.xview_moveto(1.0)

    def _copy_command(self):
        command_text = self.command_preview_text.get("1.0", tk.END).strip()
        if not command_text or "<" in command_text or ">" in command_text: self._update_status("Cannot copy incomplete command.", COLOR_WARN); return
        try:
            self.clipboard_clear(); self.clipboard_append(command_text); self._update_status("Command copied to clipboard!", COLOR_SUCCESS)
            original_text = self.copy_btn.cget("text"); self.copy_btn.config(text="OK!"); self.after(1500, lambda: self.copy_btn.config(text=original_text))
        except tk.TclError: self._update_status("Error accessing clipboard.", COLOR_ERROR)

    def _update_status(self, message, color=None):
        self.status_var.set(f" {message}")
        map_colors = { COLOR_SUCCESS: "Success.Status.TLabel", COLOR_ERROR: "Error.Status.TLabel", COLOR_INFO: "Info.Status.TLabel", COLOR_WARN: "Warn.Status.TLabel", }
        style = map_colors.get(color, "Default.Status.TLabel")
        try: self.status_bar.configure(style=style)
        except tk.TclError: self.status_bar.configure(style="Default.Status.TLabel")

    # --- Modified _check_admin_status ---
    def _check_admin_status(self):
         print(f"Running _check_admin_status. self.running_as_admin = {self.running_as_admin}") # DEBUG
         is_currently_admin = self.running_as_admin

         # Update Main Create Button
         if not is_currently_admin:
             self.create_button.config(state=tk.DISABLED)
             self.create_tooltip(self.create_button, TOOLTIP_CREATE_DISABLED)
         else:
             self.create_button.config(state=tk.NORMAL)
             self.create_tooltip(self.create_button, TOOLTIP_CREATE_ENABLED)

         # No longer need to manage history delete button state here
         # Its functionality is now tied to the View Folder button which doesn't need admin

    def _update_command_preview(self, *args):
        source = self.source_dir_var.get().strip(); link_parent = self.link_parent_dir_var.get().strip(); link_name = self.link_name_var.get().strip()
        display_source = f'"{source}"' if source else "<Source Path>"
        display_link_name = link_name if link_name else "<Link Name>"
        display_link_parent = link_parent if link_parent else "<Parent Dir>"
        if link_parent and link_name: full_link_path = os.path.normpath(os.path.join(link_parent, link_name)); display_full_link = f'"{full_link_path}"'
        else: mock_path = os.path.join(display_link_parent, display_link_name); display_full_link = f'"{mock_path}"'
        command = f'mklink /J {display_full_link} {display_source}'
        self.command_preview_text.config(state=tk.NORMAL); self.command_preview_text.delete("1.0", tk.END); self.command_preview_text.insert("1.0", command); self.command_preview_text.config(state=tk.DISABLED)
        if "<" in command or ">" in command: self.copy_btn.config(state=tk.DISABLED); self.copy_btn.bind("<Enter>", lambda e: self._show_tooltip(e, "Fill fields to enable copy."))
        else: self.copy_btn.config(state=tk.NORMAL); self.create_tooltip(self.copy_btn, TOOLTIP_COPY)

    def _create_junction(self):
        source_dir = self.source_dir_var.get().strip(); link_parent_dir = self.link_parent_dir_var.get().strip(); link_name = self.link_name_var.get().strip()
        if not all([source_dir, link_parent_dir, link_name]): self._update_status("Error: All fields required.", COLOR_ERROR); return
        if not os.path.isdir(source_dir): self._update_status(f"Error: Source not found: {source_dir}", COLOR_ERROR); return
        if not os.path.isdir(link_parent_dir): self._update_status(f"Error: Link parent not found: {link_parent_dir}", COLOR_ERROR); return
        invalid_chars = '<>:"/\\|?*';
        if any(c in invalid_chars for c in link_name) or link_name in (".", ".."): self._update_status(f"Error: Link name invalid.", COLOR_ERROR); return
        full_link_path = os.path.normpath(os.path.join(link_parent_dir, link_name))
        if os.path.lexists(full_link_path): self._update_status(f"Error: Path exists: {full_link_path}", COLOR_ERROR); messagebox.showerror("Creation Error", f"Path exists:\n{full_link_path}", parent=self); return
        if not self.running_as_admin: self._update_status("Error: Admin required.", COLOR_ERROR); messagebox.showerror("Permission Error", "Admin required.", parent=self); return
        cmd_string = f'mklink /J "{full_link_path}" "{source_dir}"'
        self._update_status(f"Executing...", COLOR_INFO); self.update_idletasks()
        try:
            result = subprocess.run(cmd_string, capture_output=True, text=True, check=False, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0 and os.path.lexists(full_link_path):
                success_msg = f"Success: '{link_name}' -> '{source_dir}'"
                self._update_status(success_msg, COLOR_SUCCESS)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.history_data.append({"source": source_dir, "link": full_link_path, "timestamp": timestamp})
                save_history(self.history_data)
                self.history_button.config(text=f"View History {ICON_HISTORY} ({len(self.history_data)})")
                if self.history_window and self.history_window.winfo_exists(): self.history_window.refresh_list()
                messagebox.showinfo("Success", success_msg, parent=self)
            else:
                error_details = result.stderr.strip() if result.stderr else result.stdout.strip()
                if not error_details: error_details = f"mklink failed (code {result.returncode}). Check paths/permissions or run as admin."
                elif "syntax error" in error_details.lower() or "syntax of the command is incorrect" in error_details.lower(): error_details = f"Syntax error (code {result.returncode}). Check paths.\n{error_details}"
                elif "already exists" in error_details.lower(): error_details = f"Path already exists (code {result.returncode}).\n{full_link_path}"
                elif "access is denied" in error_details.lower(): error_details = f"Access Denied (code {result.returncode}). Run as Admin?\n{error_details}"
                self._update_status(f"Error: {error_details}", COLOR_ERROR)
                messagebox.showerror("Command Error", f"Failed.\n\nCmd:\n{cmd_string}\n\nError:\n{error_details}", parent=self)
        except Exception as e: self._update_status(f"Error executing command: {e}", COLOR_ERROR); messagebox.showerror("Unexpected Error", f"Error during command execution:\n{e}", parent=self)

    def populate_fields_from_history(self, source, link):
        try:
            parent_dir = os.path.dirname(link); link_name = os.path.basename(link)
            self.source_dir_var.set(source); self.link_parent_dir_var.set(parent_dir); self.link_name_var.set(link_name)
            self._update_status(f"Populated fields from history: {link_name}", COLOR_INFO)
            self.lift(); self.focus_force(); self.link_name_entry.focus()
        except Exception as e: self._update_status(f"Error populating fields: {e}", COLOR_ERROR); messagebox.showerror("Error", f"Could not populate fields:\n{e}", parent=self)

    def _open_history_window(self):
        if self.history_window and self.history_window.winfo_exists(): self.history_window.lift(); self.history_window.focus()
        else: self.history_window = HistoryWindow(self, self.history_data)

    def on_closing(self):
        if self.history_window and self.history_window.winfo_exists(): self.history_window.destroy()
        self.destroy()


# --- History Window Class ---
class HistoryWindow(tk.Toplevel):
    def __init__(self, parent, history_data):
        super().__init__(parent)
        self.parent_app = parent; self.history_data = history_data
        self.title(f"{APP_NAME} - History"); self.geometry("800x500"); self.minsize(600, 300) # Adjusted width
        try:
            # Use the helper function to find the icon
            icon_path = resource_path("icon.ico")
            self.iconbitmap(icon_path) # Use self.iconbitmap() for window icon
            print(f"DEBUG: Attempting to load icon from: {icon_path}") # Optional debug print
        except tk.TclError as e:
            print(f"Warning: Could not load window icon 'icon.ico': {e}", file=sys.stderr)
        except FileNotFoundError:
             print(f"Warning: Icon file 'icon.ico' not found at expected path: {icon_path}", file=sys.stderr)
        parent_x=parent.winfo_x();parent_y=parent.winfo_y();parent_w=parent.winfo_width();parent_h=parent.winfo_height(); w=800;h=500;x=parent_x+(parent_w//2)-(w//2);y=parent_y+(parent_h//2)-(h//2); self.geometry(f"{w}x{h}+{x}+{y}")
        self.protocol("WM_DELETE_WINDOW", self.on_close); self.transient(parent); self.grab_set()

        self._create_view_options_menu() # Create menu before widgets that use it
        self._create_history_widgets()
        self.refresh_list()
        # No need to call _check_admin_status here as View Folder doesn't need admin

    # --- NEW: Create the pop-up menu ---
    def _create_view_options_menu(self):
        self.view_menu = tk.Menu(self, tearoff=0)
        self.view_menu.add_command(label="Open Link Location", command=lambda: self._open_explorer('link'))
        self.view_menu.add_command(label="Open Source Location", command=lambda: self._open_explorer('source'))
    # ---

    def _create_history_widgets(self):
        main_frame = ttk.Frame(self, padding=PAD_GENERAL); main_frame.pack(expand=True, fill=tk.BOTH)
        tree_frame = ttk.Frame(main_frame); tree_frame.pack(expand=True, fill=tk.BOTH, pady=(0, PAD_GENERAL)); tree_frame.rowconfigure(0, weight=1); tree_frame.columnconfigure(0, weight=1)
        columns = ("status", "link", "source", "created")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("status", text="Status", anchor=tk.CENTER); self.tree.heading("link", text="Junction Link Path"); self.tree.heading("source", text="Target Source Path"); self.tree.heading("created", text="Date Created", anchor=tk.W)
        # Adjusted Column Widths
        self.tree.column("status", width=110, stretch=tk.NO, anchor=tk.CENTER)
        self.tree.column("link", width=280, stretch=tk.YES)
        self.tree.column("source", width=280, stretch=tk.YES)
        self.tree.column("created", width=130, stretch=tk.NO, anchor=tk.W)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview); hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew'); vsb.grid(row=0, column=1, sticky='ns'); hsb.grid(row=1, column=0, sticky='ew')
        self.tree.tag_configure("Valid", foreground=COLOR_VALID); self.tree.tag_configure("Invalid", foreground=COLOR_INVALID); self.tree.tag_configure("Error", foreground=COLOR_WARN)

        # --- Modified Button Frame ---
        button_frame = ttk.Frame(main_frame); button_frame.pack(fill=tk.X)
        self.refresh_button = ttk.Button(button_frame, text="Refresh Status", command=self.refresh_list); self.refresh_button.pack(side=tk.LEFT, padx=(0, PAD_GENERAL))
        self.edit_button = ttk.Button(button_frame, text=f"Edit {ICON_EDIT}", command=self._edit_selected); self.edit_button.pack(side=tk.LEFT, padx=(0, PAD_GENERAL)); self.parent_app.create_tooltip(self.edit_button, TOOLTIP_EDIT)

        # --- New View Folder Button ---
        self.view_folder_button = ttk.Button(button_frame, text=f"View Folder {ICON_BROWSE}", command=self._show_view_options)
        self.view_folder_button.pack(side=tk.LEFT, padx=(0, PAD_GENERAL))
        self.parent_app.create_tooltip(self.view_folder_button, TOOLTIP_VIEW_FOLDER)
        # ---

        # --- Removed Delete Button(s) ---

        self.close_button = ttk.Button(button_frame, text="Close", command=self.on_close); self.close_button.pack(side=tk.RIGHT)
        # ---

    def refresh_list(self):
        # ... (refresh list logic remains the same) ...
        for item in self.tree.get_children(): self.tree.delete(item)
        if not self.history_data: self.tree.insert("", tk.END, values=("", "No history found.", "", "")); return
        sorted_history = sorted(self.history_data, key=lambda x: x.get("timestamp", ""), reverse=True)
        for entry in sorted_history:
            link = entry.get("link", "N/A"); source = entry.get("source", "N/A"); created = entry.get("timestamp", "N/A")
            if link == "N/A" or source == "N/A": status_icon = ICON_INVALID; status_text = "Data Error"; tag = "Error"
            else:
                status_code, status_text = check_junction_validity(link, source)
                if status_code == 0: status_icon = ICON_VALID; tag = "Valid"
                elif status_code in [1, 2, 3]: status_icon = ICON_INVALID; tag = "Invalid"
                else: status_icon = ICON_INVALID; tag = "Error"
            values = (f"{status_icon} {status_text}", link, source, created); iid = link
            try: self.tree.insert("", tk.END, iid=iid, values=values, tags=(tag,))
            except tk.TclError:
                try: self.tree.insert("", tk.END, values=values, tags=(tag,))
                except Exception as insert_e: print(f"Error inserting item even with default iid: {values}, {insert_e}")
        self._update_status(f"History refreshed. {len(sorted_history)} items.")


    def _edit_selected(self):
        # ... (edit selected logic remains the same) ...
        selected_iids = self.tree.selection()
        if not selected_iids: self._update_status("No item selected to edit.", COLOR_WARN); return
        if len(selected_iids) > 1: self._update_status("Please select only one item to edit.", COLOR_WARN); return
        selected_iid = selected_iids[0]
        try:
            item_data = self.tree.item(selected_iid); values = item_data.get('values')
            if not values or len(values) < 3: self._update_status("Could not retrieve data.", COLOR_ERROR); return
            link_path = values[1]; source_path = values[2]
            if link_path == "N/A" or source_path == "N/A" or not link_path or not source_path: self._update_status("Selected item has incomplete data.", COLOR_WARN); return
            self.parent_app.populate_fields_from_history(source_path, link_path)
            self.on_close()
        except Exception as e: self._update_status(f"Error during edit prep: {e}", COLOR_ERROR); messagebox.showerror("Error", f"Could not prepare edit:\n{e}", parent=self)

    # --- REMOVED _delete_selected_links method ---

    # --- NEW: Show View Options Menu ---
    def _show_view_options(self):
        selected_iids = self.tree.selection()
        if not selected_iids:
            self._update_status("No item selected to view.", COLOR_WARN)
            return
        if len(selected_iids) > 1:
            self._update_status("Please select only one item to view.", COLOR_WARN)
            return

        # Get button position to post menu nearby
        try:
            x = self.view_folder_button.winfo_rootx()
            y = self.view_folder_button.winfo_rooty() + self.view_folder_button.winfo_height()
            self.view_menu.post(x, y)
        except tk.TclError:
             # Fallback if button info isn't ready
             self.view_menu.post(self.winfo_rootx() + 50, self.winfo_rooty() + 100)
    # ---

    # --- NEW: Open Explorer Method ---
    def _open_explorer(self, location_type):
        """Opens File Explorer selecting the selected item's link path or source path."""
        selected_iids = self.tree.selection()
        if not selected_iids:
            self._update_status("No item selected to view.", COLOR_WARN)
            return

        selected_iid = selected_iids[0]
        path_to_select = None
        path_description = ""

        try:
            item_data = self.tree.item(selected_iid)
            values = item_data.get('values')
            if not values or len(values) < 3:
                self._update_status("Could not retrieve path data for selection.", COLOR_ERROR); return

            link_path = values[1]; source_path = values[2]

            if location_type == 'link':
                path_to_select = link_path; path_description = "Link Folder"
                if path_to_select == "N/A" or not path_to_select: self._update_status("Selected item has no valid link path.", COLOR_WARN); return
                # Check existence using lexists (doesn't follow link)
                if not os.path.lexists(path_to_select): messagebox.showwarning("Path Not Found", f"The link path does not seem to exist:\n{path_to_select}", parent=self); return
            elif location_type == 'source':
                path_to_select = source_path; path_description = "Source Folder"
                if path_to_select == "N/A" or not path_to_select: self._update_status("Selected item has no valid source path.", COLOR_WARN); return
                # Check existence using isdir (follows link, confirms target is directory)
                if not os.path.isdir(path_to_select): messagebox.showwarning("Path Not Found", f"The source path does not seem to exist or is not a directory:\n{path_to_select}", parent=self); return
            else:
                self._update_status("Invalid location type.", COLOR_ERROR); return

            # --- Prepare and execute the explorer command using /select ---
            # Normalize the path first for consistency
            norm_path_to_select = os.path.normpath(path_to_select)

            os.startfile(norm_path_to_select)


            # # Ensure single backslashes for the explorer argument string
            # path_for_explorer_arg = norm_path_to_select.replace('\\\\', '\\')

            # # Construct the argument string for explorer
            # explorer_arg = f'/select,"{path_for_explorer_arg}"'

            # # Create the command list for Popen
            # explorer_command = ['explorer', explorer_arg]

            # self._update_status(f"Opening {path_description} location in Explorer...", COLOR_INFO)
            # print(f"Executing command list: {explorer_command}") # Debugging

            # # Launch File Explorer asynchronously using the list
            # subprocess.Popen(explorer_command)

        except FileNotFoundError: messagebox.showerror("Error", "Could not find 'explorer.exe'.", parent=self); self._update_status("Error: explorer.exe not found.", COLOR_ERROR)
        except OSError as e: messagebox.showerror("OS Error", f"Could not open File Explorer (OS Error):\n{e}", parent=self); self._update_status(f"Error opening explorer (OSError): {e}", COLOR_ERROR)
        except Exception as e: messagebox.showerror("Error", f"An unexpected error occurred opening File Explorer:\n{e}", parent=self); self._update_status(f"Error opening explorer: {e}", COLOR_ERROR)

    def _update_status(self, message, color=None):
        print(f"History Status: {message}")
        self.parent_app._update_status(message, color)

    def on_close(self):
        self.parent_app.history_window = None
        self.grab_release(); self.destroy()

# --- Main Execution ---
if __name__ == "__main__":
    if os.name != 'nt':
        root = tk.Tk(); root.withdraw(); messagebox.showerror("Compatibility Error", "Requires Windows."); sys.exit(1)

    print(f"Initial check: is_admin() = {is_admin()}") # DEBUG
    HAS_ADMIN = is_admin()

    if not HAS_ADMIN:
        root = tk.Tk(); root.withdraw()
        message = ("Admin privileges needed to create links.\n\n" # Removed delete mention
                   "Without them, you can only view/copy the command.\n\n"
                   "Relaunch as Administrator?")
        if messagebox.askyesno("Admin Required", message, icon='warning'):
            print("Attempting relaunch as admin...")
            if relaunch_as_admin(): print("Relaunch success. Exiting old process."); sys.exit(0)
            else: print("Relaunch failed/cancelled."); messagebox.showinfo("Relaunch Failed", "Could not relaunch as admin.\nContinuing with limited features.", parent=None)
        else: print("Continuing without admin."); messagebox.showinfo("Limited Features", "Running without admin rights.", parent=None)
        root.destroy()

    if not ensure_dir_exists(HISTORY_DIR):
        root = tk.Tk(); root.withdraw(); messagebox.showwarning("Startup Warning", f"Could not access history folder:\n{HISTORY_DIR}\nHistory may not work."); root.destroy()

    current_admin_status = is_admin()
    print(f"Status before creating JunctionApp: is_admin() = {current_admin_status}") # DEBUG
    app = JunctionApp(running_as_admin=current_admin_status)
    app.mainloop()
