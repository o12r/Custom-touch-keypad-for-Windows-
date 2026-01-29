
import tkinter as tk
from tkinter import filedialog, colorchooser
import keyboard
import json
import win32gui
import win32con
import math
import winreg
import ctypes
import time
import threading
from collections import defaultdict

from ctypes import windll, wintypes
# Patch for missing ULONG_PTR in some Python versions
if not hasattr(wintypes, 'ULONG_PTR'):
    import sys
    if sys.maxsize > 2**32:
        wintypes.ULONG_PTR = ctypes.c_uint64
    else:
        wintypes.ULONG_PTR = ctypes.c_uint32

# --- Windows Touch API integration ---
WM_TOUCH = 0x0240
TOUCHEVENTF_MOVE = 0x0001
TOUCHEVENTF_DOWN = 0x0002
TOUCHEVENTF_UP = 0x0004
TOUCHEVENTF_INRANGE = 0x0008
TOUCHEVENTF_PRIMARY = 0x0010
TOUCHEVENTF_NOCOALESCE = 0x0020
TOUCHEVENTF_PEN = 0x0040
TOUCHEVENTF_PALM = 0x0080

class TOUCHINPUT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
        ("hSource", wintypes.HANDLE),
        ("dwID", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("dwMask", wintypes.DWORD),
        ("dwTime", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
        ("cxContact", wintypes.DWORD),
        ("cyContact", wintypes.DWORD),
    ]

user32 = ctypes.windll.user32
RegisterTouchWindow = user32.RegisterTouchWindow
RegisterTouchWindow.restype = wintypes.BOOL
RegisterTouchWindow.argtypes = [wintypes.HWND, wintypes.ULONG]
GetTouchInputInfo = user32.GetTouchInputInfo
GetTouchInputInfo.restype = wintypes.BOOL
GetTouchInputInfo.argtypes = [wintypes.HANDLE, wintypes.UINT, ctypes.POINTER(TOUCHINPUT), wintypes.INT]
CloseTouchInputHandle = user32.CloseTouchInputHandle
CloseTouchInputHandle.restype = wintypes.BOOL
CloseTouchInputHandle.argtypes = [wintypes.HANDLE]


# Set process priority to HIGH for maximum responsiveness
try:
    ctypes.windll.kernel32.SetPriorityClass(ctypes.windll.kernel32.GetCurrentProcess(), 0x00000080)  # HIGH_PRIORITY_CLASS
except:
    pass

# --- DirectInput Configuration for Games ---
# Games require 'Scan Codes', not just virtual key presses.
# This structure defines the C input structures needed for SendInput.

SendInput = ctypes.windll.user32.SendInput

# C struct definitions
class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]

# Common DirectInput Scan Codes
SCAN_CODES = {
    'esc': 0x01, '1': 0x02, '2': 0x03, '3': 0x04, '4': 0x05, '5': 0x06, '6': 0x07, '7': 0x08, '8': 0x09, '9': 0x0A, '0': 0x0B,
    'q': 0x10, 'w': 0x11, 'e': 0x12, 'r': 0x13, 't': 0x14, 'y': 0x15, 'u': 0x16, 'i': 0x17, 'o': 0x18, 'p': 0x19,
    'a': 0x1E, 's': 0x1F, 'd': 0x20, 'f': 0x21, 'g': 0x22, 'h': 0x23, 'j': 0x24, 'k': 0x25, 'l': 0x26,
    'z': 0x2C, 'x': 0x2D, 'c': 0x2E, 'v': 0x2F, 'b': 0x30, 'n': 0x31, 'm': 0x32,
    'space': 0x39, 'enter': 0x1C, 'shift': 0x2A, 'ctrl': 0x1D, 'alt': 0x38,
    'up': 0xC8, 'left': 0xCB, 'right': 0xCD, 'down': 0xD0,
    'f1': 0x3B, 'f2': 0x3C, 'f3': 0x3D, 'f4': 0x3E, 'f5': 0x3F, 'f6': 0x40, 
    'f7': 0x41, 'f8': 0x42, 'f9': 0x43, 'f10': 0x44, 'f11': 0x57, 'f12': 0x58
}

def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    # 0x0008 represents KEYEVENTF_SCANCODE
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    # 0x0008 | 0x0002 represents KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

# --- Standard App Config ---
DEFAULT_ROWS = 4
DEFAULT_COLS = 5
BASE_UNIT = 10 
REPEAT_INTERVAL = 50   
DRAG_THRESHOLD = 5
# Optimization constants
EVENT_THROTTLE_MS = 2  # Tighter throttle for faster motion response
KEY_PRESS_DELAY = 12   # Faster key release for snappier feel
RENDER_BATCH_INTERVAL = 15  # Batch visual updates     

def get_system_accent():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\DWM")
        value, _ = winreg.QueryValueEx(key, "AccentColor")
        a = (value >> 24) & 0xFF
        b = (value >> 16) & 0xFF
        g = (value >> 8) & 0xFF
        r = value & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception: return "#0078d7"

def get_system_mode():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "Light" if value == 1 else "Dark"
    except Exception: return "Light"

THEMES = {
    "Modern Light": { "bg": "#f9f9f9", "fg": "#000000", "grid_bg": "#d0d0d0", "btn_bg": "#ffffff", "btn_fg": "#000000", "btn_hover": "#0078d7", "btn_active": "#e0e0e0", "header_bg": "#f9f9f9", "header_fg": "#000000", "accent": "#0078d7", "relief": "flat", "border": 0, "gap": 1, "separator": "#e0e0e0", "slider_bg": "#f9f9f9", "slider_trough": "#e0e0e0", "slider_active": "#0078d7" },
    "Modern Dark": { "bg": "#202020", "fg": "#ffffff", "grid_bg": "#3a3a3a", "btn_bg": "#2d2d2d", "btn_fg": "#ffffff", "btn_hover": "#007acc", "btn_active": "#404040", "header_bg": "#202020", "header_fg": "#aaaaaa", "accent": "#007acc", "relief": "flat", "border": 0, "gap": 1, "separator": "#3a3a3a", "slider_bg": "#202020", "slider_trough": "#404040", "slider_active": "#007acc" },
    "Retro Gray": { "bg": "#c0c0c0", "fg": "#000000", "grid_bg": "#c0c0c0", "btn_bg": "#c0c0c0", "btn_fg": "#000000", "btn_hover": "#000080", "btn_active": "#aaaaaa", "header_bg": "#c0c0c0", "header_fg": "#000000", "accent": "#000080", "relief": "raised", "border": 2, "gap": 0, "separator": "#808080", "slider_bg": "#c0c0c0", "slider_trough": "#808080", "slider_active": "#000080" },
}

class ControlPanel(tk.Toplevel):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.geometry("260x420+100+100")
        
        self.expanded = True
        self.drag_data = {"x": 0, "y": 0, "start_x": 0, "start_y": 0, "moved": False}
        
        self.frame_full = tk.Frame(self)
        
        self.drag_bar = tk.Label(self.frame_full, text="::: Controls :::", cursor="fleur", pady=8, font=("Segoe UI", 9, "bold"))
        self.drag_bar.pack(fill="x", pady=(0, 0))
        self.drag_bar.bind("<ButtonPress-1>", self.start_move)
        self.drag_bar.bind("<B1-Motion>", self.do_move)

        self.btn_container = tk.Frame(self.frame_full)
        self.btn_container.pack(fill="both", expand=True, padx=10, pady=5)
        self.tool_buttons = [] 
        self.separators = []
        self.btn_lock = None
        
        self.create_buttons()

        self.frame_mini = tk.Frame(self, bg="red")
        
        self.btn_unlock = tk.Label(self.frame_mini, text="🔓", bg="#ff4444", fg="white", font=("Segoe UI Emoji", 14))
        self.btn_unlock.pack(fill="both", expand=True)
        self.btn_unlock.bind("<ButtonPress-1>", self.start_unlock_drag)
        self.btn_unlock.bind("<B1-Motion>", self.do_move)
        self.btn_unlock.bind("<ButtonRelease-1>", self.end_unlock_drag)

        self.set_mode("design")

    def create_buttons(self):
        def add_sep():
            f = tk.Frame(self.btn_container, height=1)
            f.pack(fill="x", pady=8)
            self.separators.append(f)

        r0 = tk.Frame(self.btn_container)
        r0.pack(fill="x", pady=(5,0))
        self.mk_btn(r0, "Exit ❌", self.app.quit_app, "#cc0000")

        add_sep()

        r1 = tk.Frame(self.btn_container)
        r1.pack(fill="x")
        self.btn_lock = self.mk_btn(r1, "Lock 🔒", self.app.toggle_mode, self.app.current_theme['accent'])
        self.mk_btn(r1, "Save 💾", self.app.save_layout)
        self.mk_btn(r1, "Load 📂", self.app.load_layout)
        
        add_sep()

        r2 = tk.Frame(self.btn_container)
        r2.pack(fill="x")
        self.mk_btn(r2, "+ Row", self.app.add_row_at_selection)
        self.mk_btn(r2, "- Row", self.app.del_row_at_selection)
        
        r3 = tk.Frame(self.btn_container)
        r3.pack(fill="x", pady=4)
        self.mk_btn(r3, "+ Col", self.app.add_col_at_selection)
        self.mk_btn(r3, "- Col", self.app.del_col_at_selection)

        add_sep()

        r4 = tk.Frame(self.btn_container)
        r4.pack(fill="x")
        self.mk_btn(r4, "Merge", self.app.merge_cells)
        self.mk_btn(r4, "Split", self.app.unmerge_cells)

        add_sep()

        r5 = tk.Frame(self.btn_container)
        r5.pack(fill="x")
        self.mk_btn(r5, "Theme 🎨", self.app.open_theme_menu)
        self.mk_btn(r5, "Color 🌈", self.app.pick_accent_color)

        self.btn_slide = tk.Button(self.btn_container, text="Mode: TYPE ⌨️", command=self.toggle_rapid)
        self.btn_slide.pack(fill="x", pady=6)
        self.btn_slide.configure(relief="flat", borderwidth=0, highlightthickness=0)
        self.tool_buttons.append(self.btn_slide)

        tk.Label(self.btn_container, text="Opacity", font=("Segoe UI", 8)).pack(anchor="w")
        self.opacity_slider = tk.Scale(self.btn_container, from_=0.2, to=1.0, resolution=0.05, 
                                       orient="horizontal", command=self.app.update_opacity, showvalue=0,
                                       relief="flat", borderwidth=0, highlightthickness=0)
        self.opacity_slider.set(1.0)
        self.opacity_slider.pack(fill="x")
        
        add_sep()
        
        tk.Label(self.btn_container, text="Visual Feedback", font=("Segoe UI", 8)).pack(anchor="w")
        self.feedback_btn = tk.Button(self.btn_container, text="Feedback: ON", command=self.toggle_visual_feedback, 
                                      font=("Segoe UI", 9))
        self.feedback_btn.pack(fill="x", pady=4)
        self.feedback_btn.configure(relief="flat", borderwidth=0, highlightthickness=0)
        self.tool_buttons.append(self.feedback_btn)
        
        add_sep()
        
        tk.Label(self.btn_container, text="Debounce (ms)", font=("Segoe UI", 8)).pack(anchor="w")
        self.debounce_slider = tk.Scale(self.btn_container, from_=0, to=100, resolution=1, 
                                        orient="horizontal", command=self.set_debounce_time,
                                        showvalue=1, relief="flat", borderwidth=0, highlightthickness=0)
        self.debounce_slider.set(0)
        self.debounce_slider.pack(fill="x")

    def toggle_rapid(self):
        self.app.rapid_mode = not self.app.rapid_mode
        txt = "Mode: SLIDE 〰️" if self.app.rapid_mode else "Mode: TYPE ⌨️"
        self.btn_slide.configure(text=txt)
        self.app.update_input_bindings()
    
    def toggle_visual_feedback(self):
        self.app.visual_feedback_enabled = not self.app.visual_feedback_enabled
        txt = "Feedback: ON" if self.app.visual_feedback_enabled else "Feedback: OFF"
        self.feedback_btn.configure(text=txt)
        # Update bindings to apply/remove hover effects
        self.app.update_input_bindings()
    
    def set_debounce_time(self, value):
        """Update global debounce threshold."""
        self.app.key_press_threshold = int(float(value))

    def mk_btn(self, parent, text, cmd, special_color=None):
        b = tk.Button(parent, text=text, command=cmd, font=("Segoe UI", 9))
        b.pack(side="left", fill="x", expand=True, padx=2)
        b.configure(relief="flat", borderwidth=0, highlightthickness=0)
        if special_color: b.special_color = special_color
        b.bind("<Enter>", lambda e: self.on_hover(b, True))
        b.bind("<Leave>", lambda e: self.on_hover(b, False))
        self.tool_buttons.append(b)
        return b

    def on_hover(self, btn, hovering):
        t = self.app.current_theme
        if hasattr(btn, 'special_color'): return 
        bg = t['btn_hover'] if hovering else t['btn_bg']
        fg = "#ffffff" if hovering else t['btn_fg']
        btn.configure(bg=bg, fg=fg)

    def start_move(self, event):
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y

    def do_move(self, event):
        x = self.winfo_x() + (event.x - self.drag_data["x"])
        y = self.winfo_y() + (event.y - self.drag_data["y"])
        self.geometry(f"+{x}+{y}")
        self.drag_data["moved"] = True

    def start_unlock_drag(self, event):
        self.start_move(event)
        self.drag_data["moved"] = False
        self.drag_data["start_x"] = event.x_root
        self.drag_data["start_y"] = event.y_root

    def end_unlock_drag(self, event):
        dist = math.hypot(event.x_root - self.drag_data["start_x"], event.y_root - self.drag_data["start_y"])
        if dist < 5: self.app.toggle_mode()

    def set_mode(self, mode):
        self.expanded = (mode == "design")
        current_x = self.winfo_x()
        current_y = self.winfo_y()
        if self.expanded:
            self.frame_mini.pack_forget()
            self.frame_full.pack(fill="both", expand=True)
            self.geometry(f"220x440+{current_x}+{current_y}")
        else:
            self.frame_full.pack_forget()
            self.frame_mini.pack(fill="both", expand=True)
            self.geometry(f"40x40+{current_x}+{current_y}")

    def update_theme(self, t):
        self.configure(bg=t['bg'])
        self.frame_full.configure(bg=t['bg'])
        self.btn_container.configure(bg=t['bg'])
        self.drag_bar.configure(bg=t['accent'], fg="#ffffff")
        for s in self.separators: s.configure(bg=t['separator'])
        for b in self.tool_buttons:
            relief = t['relief']
            bd = t['border']
            if b == self.btn_lock:
                b.special_color = t['accent']
                b.configure(bg=t['accent'], fg="white", relief=relief, bd=bd)
            elif hasattr(b, 'special_color'):
                b.configure(bg=b.special_color, fg="white", relief=relief, bd=bd)
            else:
                b.configure(bg=t['btn_bg'], fg=t['btn_fg'], relief=relief, bd=bd)
        for w in self.btn_container.winfo_children():
            if isinstance(w, tk.Label): w.configure(bg=t['bg'], fg=t['fg'])
        self.opacity_slider.configure(bg=t['slider_bg'], troughcolor=t['slider_trough'], activebackground=t['slider_active'], fg=t['fg'])

class VirtualKeyboardApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Grid Controller")
        self.overrideredirect(True)
        self.resizable(False, False)

        self.mode = "design"
        self.design_geometry = None
        self.grid_data = {}
        self.row_sizes = [40] * DEFAULT_ROWS
        self.col_sizes = [60] * DEFAULT_COLS
        self.selected_cells = set()
        self.selected_rows_indices = set()
        self.selected_cols_indices = set()
        self.rapid_mode = False
        self.current_theme = THEMES["Modern Light"]
        self.active_entry = None
        self.button_refs = {}
        self.drag_data = {"x": 0, "y": 0}

        # Visual feedback toggle
        self.visual_feedback_enabled = True

        # Multi-touch tracking: touch_id -> {"widget": w, "key": k, "active": True, "time": t}
        self.active_touches = {}
        self.next_touch_id = 0

        # Event throttling
        self.last_motion_event = 0
        self.motion_pending = False
        self.pending_render = False
        self.render_batch_job = None

        # High-speed input tracking
        self.active_keys = {}  # key -> timestamp of last press
        self.pressed_buttons = {}  # button widget id -> (key, time, is_pressed)
        self.button_states = {}  # button widget id -> {"feedback_job": id, "orig_color": color}
        self.button_pressed_state = {}  # Track if button is physically pressed (not released yet)
        self.last_button_hit = None  # Track last button to distinguish new presses
        self.key_press_threshold = 0  # ms between same-button presses (0 = no debounce, user configurable)
        self.input_lock = threading.Lock()  # Thread-safe key tracking

        self.mouse_pressed = False
        self.active_key = None
        self.repeat_job = None
        self.drag_start_pos = None
        self.paint_value = None
        self.interaction_type = None

        for r in range(DEFAULT_ROWS):
            for c in range(DEFAULT_COLS):
                self.grid_data[(r, c)] = {'key': '', 'span_r': 1, 'span_c': 1}

        self.main_frame = tk.Frame(self, bg="gray")
        self.main_frame.pack(fill="both", expand=True)

        self.title_bar = tk.Label(self.main_frame, text="::: Grid Editor :::", cursor="fleur", pady=4, font=("Segoe UI", 8))
        self.title_bar.pack(fill="x", side="top")
        self.title_bar.bind("<ButtonPress-1>", self.start_window_move)
        self.title_bar.bind("<B1-Motion>", self.do_window_move)

        self.grid_frame = tk.Frame(self.main_frame)
        self.grid_frame.pack(fill="both", expand=True)

        self.panel = ControlPanel(self, self)
        self.bind_all("<Button-1>", self.global_click_handler)

        self.apply_theme("Modern Light")
        self.refresh_grid()
        self.fit_window_to_content()
        self.center_window()

        self.update_idletasks()
        x, y = self.winfo_x(), self.winfo_y()
        self.panel.geometry(f"+{x+300}+{y}")

        # --- Register for Windows Touch events ---
        self.after(100, self._register_touch_window)
        self._orig_wndproc = None
        self._touch_id_map = {}  # Windows touch ID -> our touch_id
        self._touch_down_widgets = {}  # touch_id -> widget
        self._setup_touch_wndproc()

    def _register_touch_window(self):
        hwnd = self.winfo_id()
        RegisterTouchWindow(hwnd, 0)


    def _setup_touch_wndproc(self):
        # Subclass the window proc to intercept WM_TOUCH
        import sys
        if sys.platform != "win32":
            return
        hwnd = self.winfo_id()
        GWL_WNDPROC = -4
        WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_long, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        user32 = ctypes.windll.user32
        # Set argtypes/restype for pointer safety
        user32.GetWindowLongPtrW.restype = ctypes.c_void_p
        user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.SetWindowLongPtrW.restype = ctypes.c_void_p
        user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        user32.DefWindowProcW.restype = ctypes.c_long
        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]

        old_wndproc = user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC)


        def py_wndproc(hWnd, msg, wParam, lParam):
            if msg == WM_TOUCH:
                self._handle_wm_touch(wParam, lParam)
                # Explicitly call CloseTouchInputHandle and return TRUE to suppress feedback
                return 1  # TRUE: message handled
            return user32.DefWindowProcW(hWnd, msg, wParam, lParam)

        self._orig_wndproc = WNDPROC(py_wndproc)
        user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, ctypes.cast(self._orig_wndproc, ctypes.c_void_p))

    def _handle_wm_touch(self, wParam, lParam):
        cInputs = wParam & 0xffff
        hTouchInput = lParam
        inputs = (TOUCHINPUT * cInputs)()
        if not GetTouchInputInfo(hTouchInput, cInputs, inputs, ctypes.sizeof(TOUCHINPUT)):
            return
        if not hasattr(self, '_touch_repeat_jobs'):
            self._touch_repeat_jobs = {}
        for ti in inputs:
            x = ti.x // 100  # Touch coordinates are in 1/100 of a pixel
            y = ti.y // 100
            touch_id = ti.dwID
            flags = ti.dwFlags
            # Map Windows touch ID to our own
            if flags & TOUCHEVENTF_DOWN:
                widget = self.get_touch_at_position(x, y)
                if widget:
                    self._touch_id_map[touch_id] = self.register_touch(widget)
                    self._touch_down_widgets[self._touch_id_map[touch_id]] = widget
                    # Immediately send key signal for this finger
                    self.check_input_hit(widget, force=True)
                    # Start repeat loop for this finger
                    self._start_touch_repeat(self._touch_id_map[touch_id], widget)
            elif flags & TOUCHEVENTF_UP:
                if touch_id in self._touch_id_map:
                    our_touch_id = self._touch_id_map[touch_id]
                    self._stop_touch_repeat(our_touch_id)
                    self.unregister_touch(our_touch_id)
                    self._touch_down_widgets.pop(our_touch_id, None)
                    self._touch_id_map.pop(touch_id, None)
            elif flags & TOUCHEVENTF_MOVE:
                # On move, check if finger is still on the same widget, if so, send signal
                widget = self.get_touch_at_position(x, y)
                if touch_id in self._touch_id_map:
                    our_touch_id = self._touch_id_map[touch_id]
                    prev_widget = self._touch_down_widgets.get(our_touch_id)
                    if widget and widget == prev_widget:
                        self.check_input_hit(widget, force=True)
                    elif widget and widget != prev_widget:
                        # Finger moved to a new button: stop old repeat, start new
                        self._stop_touch_repeat(our_touch_id)
                        self._touch_down_widgets[our_touch_id] = widget
                        self.check_input_hit(widget, force=True)
                        self._start_touch_repeat(our_touch_id, widget)
        CloseTouchInputHandle(lParam)

    def _start_touch_repeat(self, touch_id, widget):
        # Schedule next repeat for this finger
        def repeat():
            if touch_id in self._touch_down_widgets:
                self.check_input_hit(widget, force=True)
                self._touch_repeat_jobs[touch_id] = self.after(REPEAT_INTERVAL, repeat)
        self._touch_repeat_jobs[touch_id] = self.after(REPEAT_INTERVAL, repeat)

    def _stop_touch_repeat(self, touch_id):
        if hasattr(self, '_touch_repeat_jobs') and touch_id in self._touch_repeat_jobs:
            self.after_cancel(self._touch_repeat_jobs[touch_id])
            del self._touch_repeat_jobs[touch_id]

    def _start_touch_repeat(self, touch_id, widget):
        # Immediately send key and start repeat
        self._touch_repeat_now(touch_id, widget)
        # Schedule next repeat
        if not hasattr(self, '_touch_repeat_jobs'):
            self._touch_repeat_jobs = {}
        def repeat():
            if touch_id in self._touch_down_widgets:
                self._touch_repeat_now(touch_id, widget)
                self._touch_repeat_jobs[touch_id] = self.after(REPEAT_INTERVAL, repeat)
        self._touch_repeat_jobs[touch_id] = self.after(REPEAT_INTERVAL, repeat)

    def _stop_touch_repeat(self, touch_id):
        if hasattr(self, '_touch_repeat_jobs') and touch_id in self._touch_repeat_jobs:
            self.after_cancel(self._touch_repeat_jobs[touch_id])
            del self._touch_repeat_jobs[touch_id]

    def _touch_repeat_now(self, touch_id, widget):
        # Send key signal for this widget
        self.check_input_hit(widget, force=True)

    def start_window_move(self, event):
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
    def do_window_move(self, event):
        x = self.winfo_x() + (event.x - self.drag_data["x"])
        y = self.winfo_y() + (event.y - self.drag_data["y"])
        self.geometry(f"+{x}+{y}")

    def quit_app(self): self.destroy()

    def safe_commit_entry(self):
        if self.active_entry:
            try:
                if self.active_entry.winfo_exists(): self.active_entry.event_generate("<Return>")
                else: self.active_entry = None
            except tk.TclError: self.active_entry = None

    def global_click_handler(self, event):
        if self.active_entry:
            if event.widget != self.active_entry:
                self.safe_commit_entry()

    def center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        ws, hs = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{int((ws/2)-(w/2))}+{int((hs/2)-(h/2))}")

    def fit_window_to_content(self):
        self.update_idletasks()
        grid_w = sum(self.col_sizes)
        grid_h = sum(self.row_sizes)
        title_h = 30 
        if self.mode == "play": self.geometry(f"{grid_w}x{grid_h}")
        else: self.geometry(f"{grid_w+30}x{grid_h+20+title_h}")

    def refresh_grid(self):
        self.button_refs = {}
        for w in self.grid_frame.winfo_children(): w.destroy()
        t = self.current_theme
        rows, cols = len(self.row_sizes), len(self.col_sizes)
        offset = 1 if self.mode == "design" else 0
        
        if self.mode == "design":
            self.grid_frame.columnconfigure(0, minsize=30, weight=0) 
            self.grid_frame.rowconfigure(0, minsize=20, weight=0)    
            tk.Label(self.grid_frame, bg=t['bg']).grid(row=0, column=0, sticky="nsew")

            for c in range(cols):
                val = self.col_sizes[c]//BASE_UNIT
                bg = t['accent'] if c in self.selected_cols_indices else t['header_bg']
                fg = "#ffffff" if c in self.selected_cols_indices else t['header_fg']
                lbl = tk.Label(self.grid_frame, text=str(val), bg=bg, fg=fg, relief="flat", bd=0)
                lbl.grid(row=0, column=c+offset, sticky="nsew", padx=1, pady=1)
                lbl.bind("<ButtonPress-1>", lambda e, c=c, v=val: self.on_header_press(e, 'col', c, v))
                lbl.bind("<B1-Motion>", lambda e: self.on_header_drag(e, 'col'))
                lbl.bind("<ButtonRelease-1>", lambda e, c=c, l=lbl, v=val: self.on_header_release(e, 'col', c, l, str(v)))
                lbl.bind("<ButtonPress-3>", lambda e, c=c: self.on_header_select_start(e, 'col', c))
                lbl.bind("<B3-Motion>", lambda e: self.on_header_select_drag(e, 'col'))

            for r in range(rows):
                val = self.row_sizes[r]//BASE_UNIT
                bg = t['accent'] if r in self.selected_rows_indices else t['header_bg']
                fg = "#ffffff" if r in self.selected_rows_indices else t['header_fg']
                lbl = tk.Label(self.grid_frame, text=str(val), bg=bg, fg=fg, relief="flat", bd=0)
                lbl.grid(row=r+offset, column=0, sticky="nsew", padx=1, pady=1)
                lbl.bind("<ButtonPress-1>", lambda e, r=r, v=val: self.on_header_press(e, 'row', r, v))
                lbl.bind("<B1-Motion>", lambda e: self.on_header_drag(e, 'row'))
                lbl.bind("<ButtonRelease-1>", lambda e, r=r, l=lbl, v=val: self.on_header_release(e, 'row', r, l, str(v)))
                lbl.bind("<ButtonPress-3>", lambda e, r=r: self.on_header_select_start(e, 'row', r))
                lbl.bind("<B3-Motion>", lambda e: self.on_header_select_drag(e, 'row'))

        for r in range(rows): self.grid_frame.rowconfigure(r + offset, minsize=self.row_sizes[r], weight=0)
        for c in range(cols): self.grid_frame.columnconfigure(c + offset, minsize=self.col_sizes[c], weight=0)

        skip = set()
        for r in range(rows):
            for c in range(cols):
                if (r, c) in skip: continue
                cell = self.grid_data.get((r,c), {})
                span_r, span_c = cell.get('span_r', 1), cell.get('span_c', 1)
                
                if span_r > 1 or span_c > 1:
                    for sr in range(span_r):
                        for sc in range(span_c):
                            if sr==0 and sc==0: continue
                            skip.add((r+sr, c+sc))

                is_sel = (r, c) in self.selected_cells
                bg = t['accent'] if is_sel else t['btn_bg']
                fg = "#ffffff" if is_sel else t['btn_fg']
                
                if r >= rows or c >= cols: continue

                btn = tk.Button(self.grid_frame, text=cell.get('key',''), bg=bg, fg=fg,
                                relief=t['relief'], bd=t['border'], highlightthickness=0,
                                activebackground=t.get('btn_active', bg), activeforeground=fg)
                btn.grid_pos = (r, c) 
                btn.meta_key = cell.get('key','')
                self.button_refs[(r,c)] = btn 
                
                px, py = t['gap'], t['gap']
                btn.grid(row=r+offset, column=c+offset, rowspan=span_r, columnspan=span_c, sticky="nsew", padx=(0, px), pady=(0, py))
                
                if self.mode == "design":
                    btn.bind("<Enter>", lambda e, b=btn: self.on_btn_hover(b, True))
                    btn.bind("<Leave>", lambda e, b=btn: self.on_btn_hover(b, False))
                    btn.bind("<ButtonPress-1>", lambda e, r=r, c=c: self.on_cell_press(e, r, c))
                    btn.bind("<B1-Motion>", self.on_cell_paint_drag)
                    btn.bind("<ButtonRelease-1>", lambda e, r=r, c=c, b=btn: self.on_cell_release(e, r, c, b))
                    
                    btn.bind("<ButtonPress-3>", lambda e, r=r, c=c: self.start_drag_select(r, c))
                    btn.bind("<B3-Motion>", lambda e: self.do_drag_select(e))
                    btn.bind("<ButtonRelease-3>", lambda e: self.end_drag_select())
        
        self.update_input_bindings()

    def on_header_press(self, event, axis, idx, val):
        self.drag_start_pos = (event.x_root, event.y_root)
        self.paint_value = val 
        self.interaction_type = 'edit_wait'

    def on_header_drag(self, event, axis):
        if self.interaction_type == 'edit_wait':
            dist = math.hypot(event.x_root - self.drag_start_pos[0], event.y_root - self.drag_start_pos[1])
            if dist > DRAG_THRESHOLD:
                self.interaction_type = 'paint'
        if self.interaction_type == 'paint':
            x, y = event.x_root, event.y_root
            widget = self.winfo_containing(x, y)
            if widget and widget.master == self.grid_frame:
                info = widget.grid_info()
                if 'column' in info and 'row' in info:
                    c, r = int(info['column']), int(info['row'])
                    offset = 1 
                    target_idx = -1
                    if axis == 'col' and r == 0 and c >= offset: target_idx = c - offset
                    elif axis == 'row' and c == 0 and r >= offset: target_idx = r - offset
                    if target_idx != -1:
                        try:
                            new_size = int(self.paint_value * BASE_UNIT)
                            changed = False
                            if axis == 'col' and self.col_sizes[target_idx] != new_size:
                                self.col_sizes[target_idx] = new_size; changed = True
                            elif axis == 'row' and self.row_sizes[target_idx] != new_size:
                                self.row_sizes[target_idx] = new_size; changed = True
                            if changed:
                                widget.configure(text=str(self.paint_value))
                                self.fit_window_to_content()
                                self.after(10, self.fit_window_to_content)
                        except: pass

    def on_header_release(self, event, axis, idx, widget, current_val):
        if self.interaction_type == 'edit_wait':
            ctx = {'type': axis, 'idx': idx}
            self.start_inline_edit(widget, current_val, lambda v: self.finish_header_resize(axis, idx, v), ctx)
        if self.interaction_type == 'paint':
            self.refresh_grid(); self.fit_window_to_content()
        self.interaction_type = None

    def on_header_select_start(self, event, axis, idx):
        if axis == 'col':
            if idx not in self.selected_cols_indices: self.selected_cols_indices = {idx}
            self.selected_rows_indices.clear()
        else:
            if idx not in self.selected_rows_indices: self.selected_rows_indices = {idx}
            self.selected_cols_indices.clear()
        self.drag_start_pos = (event.x_root, event.y_root) 
        self.interaction_type = 'select'
        self.refresh_grid()

    def on_header_select_drag(self, event, axis):
        if self.interaction_type == 'select':
            x, y = event.x_root, event.y_root
            widget = self.winfo_containing(x, y)
            if widget and widget.master == self.grid_frame:
                info = widget.grid_info()
                offset = 1
                if 'column' in info:
                    if axis == 'col' and int(info['row']) == 0:
                        c = int(info['column']) - offset
                        if 0 <= c < len(self.col_sizes):
                            self.selected_cols_indices.add(c); self.refresh_grid()
                    elif axis == 'row' and int(info['column']) == 0:
                        r = int(info['row']) - offset
                        if 0 <= r < len(self.row_sizes):
                            self.selected_rows_indices.add(r); self.refresh_grid()

    def finish_header_resize(self, axis, idx, val):
        try:
            new_size = int(float(val) * BASE_UNIT)
            indices = self.selected_cols_indices if axis == 'col' else self.selected_rows_indices
            target_list = self.col_sizes if axis == 'col' else self.row_sizes
            if idx in indices:
                for i in indices:
                    if 0 <= i < len(target_list): target_list[i] = new_size
            else: target_list[idx] = new_size
            self.refresh_grid(); self.fit_window_to_content()
        except: pass

    def on_cell_press(self, event, r, c):
        self.drag_start_pos = (event.x_root, event.y_root)
        self.paint_value = self.grid_data.get((r,c), {}).get('key', '')
        self.interaction_type = 'edit_wait'

    def on_cell_paint_drag(self, event):
        if self.interaction_type == 'edit_wait':
            dist = math.hypot(event.x_root - self.drag_start_pos[0], event.y_root - self.drag_start_pos[1])
            if dist > DRAG_THRESHOLD: self.interaction_type = 'paint'
        if self.interaction_type == 'paint':
            x, y = event.x_root, event.y_root
            widget = self.winfo_containing(x, y)
            if widget and isinstance(widget, tk.Button) and hasattr(widget, 'grid_pos'):
                r, c = widget.grid_pos
                if self.grid_data[(r,c)]['key'] != self.paint_value:
                    self.grid_data[(r,c)]['key'] = self.paint_value
                    widget.configure(text=self.paint_value)
                    widget.meta_key = self.paint_value

    def on_cell_release(self, event, r, c, widget):
        if self.interaction_type == 'edit_wait':
            ctx = {'type': 'cell', 'r': r, 'c': c}
            self.start_inline_edit(widget, self.grid_data[(r,c)].get('key', ''), lambda v: self.finish_key_edit(r, c, v), ctx)
        self.interaction_type = None

    def start_inline_edit(self, widget, initial_text, callback, context=None):
        if self.active_entry: self.safe_commit_entry()
        if not widget.winfo_exists(): return
        t = self.current_theme
        entry = tk.Entry(widget.master, justify='center', relief="flat", bd=1)
        entry.config(highlightbackground=t['accent'], highlightthickness=1)
        entry.insert(0, initial_text)
        entry.select_range(0, tk.END)
        entry.focus_set()
        x, y = widget.winfo_x(), widget.winfo_y()
        w, h = widget.winfo_width(), widget.winfo_height()
        entry.place(x=x, y=y, width=w, height=h)
        def commit(event=None):
            if not self.active_entry: return
            val = entry.get()
            self.active_entry = None
            entry.destroy()
            callback(val)
            return "break"
        def commit_and_next(event=None):
            commit()
            if context: self.goto_next_editable(context)
            return "break"
        def cancel(event=None):
            if not self.active_entry: return
            self.active_entry = None
            entry.destroy()
        entry.bind("<Return>", commit)
        entry.bind("<Tab>", commit_and_next)
        entry.bind("<Escape>", cancel)
        self.active_entry = entry

    def goto_next_editable(self, ctx):
        next_widget = None
        if ctx['type'] == 'col':
            idx = ctx['idx'] + 1
            if idx < len(self.col_sizes):
                for w in self.grid_frame.winfo_children():
                    info = w.grid_info()
                    if int(info['row']) == 0 and int(info['column']) == idx + 1:
                        next_widget = w; break
                if next_widget:
                    val = str(self.col_sizes[idx]//BASE_UNIT)
                    self.start_inline_edit(next_widget, val, lambda v: self.finish_header_resize('col', idx, v), {'type':'col', 'idx':idx})
        elif ctx['type'] == 'row':
            idx = ctx['idx'] + 1
            if idx < len(self.row_sizes):
                for w in self.grid_frame.winfo_children():
                    info = w.grid_info()
                    if int(info['column']) == 0 and int(info['row']) == idx + 1:
                        next_widget = w; break
                if next_widget:
                    val = str(self.row_sizes[idx]//BASE_UNIT)
                    self.start_inline_edit(next_widget, val, lambda v: self.finish_header_resize('row', idx, v), {'type':'row', 'idx':idx})
        elif ctx['type'] == 'cell':
            r, c = ctx['r'], ctx['c']
            next_c = c + 1; next_r = r
            if next_c >= len(self.col_sizes):
                next_c = 0; next_r += 1
            if next_r < len(self.row_sizes):
                if (next_r, next_c) in self.button_refs:
                    btn = self.button_refs[(next_r, next_c)]
                    val = self.grid_data[(next_r, next_c)].get('key','')
                    self.start_inline_edit(btn, val, lambda v: self.finish_key_edit(next_r, next_c, v), {'type':'cell', 'r':next_r, 'c':next_c})

    def on_btn_hover(self, btn, hovering):
        r, c = btn.grid_pos
        if (r,c) in self.selected_cells: return
        # Skip hover effect if visual feedback is disabled
        if not self.visual_feedback_enabled:
            return
        t = self.current_theme
        bg = t['btn_hover'] if hovering else t['btn_bg']
        fg = "#ffffff" if hovering else t['btn_fg']
        btn.configure(bg=bg, fg=fg)

    def update_input_bindings(self):
        if self.mode != "play": return
        rapid = self.rapid_mode
        for btn in self.button_refs.values():
            btn.unbind("<ButtonPress-1>")
            btn.unbind("<ButtonRelease-1>")
            btn.unbind("<B1-Motion>")
            btn.unbind("<Enter>")
            btn.unbind("<Leave>")
            if not rapid:
                # Only add hover effects if visual feedback is enabled
                if self.visual_feedback_enabled:
                    btn.bind("<Enter>", lambda e, b=btn: self.on_btn_hover(b, True))
                    btn.bind("<Leave>", lambda e, b=btn: self.on_btn_hover(b, False))
            if rapid:
                btn.bind("<ButtonPress-1>", self.on_press)
                btn.bind("<ButtonRelease-1>", self.on_release)
                btn.bind("<B1-Motion>", self.on_motion)
                btn.configure(command=lambda: None) 
            else:
                btn.configure(command=lambda k=btn.meta_key, b=btn: self.play_key_pulse(k, b))

    def update_visuals(self):
        t = self.current_theme
        for (r,c), widget in self.button_refs.items():
            try:
                if not widget.winfo_exists(): continue
                is_sel = (r, c) in self.selected_cells
                target_bg = t['accent'] if is_sel else t['btn_bg']
                target_fg = "#ffffff" if is_sel else t['btn_fg']
                if widget.cget('bg') != target_bg and widget.cget('bg') != "#aaaaaa":
                    widget.configure(bg=target_bg, fg=target_fg, relief=t['relief'], bd=t['border'])
            except: pass

    def on_press(self, event):
        w = event.widget
        if not isinstance(w, tk.Button) or not hasattr(w, "meta_key"): 
            return
        self.mouse_pressed = True
        self.last_button_hit = id(w)  # Track this press start
        self.check_input_hit(w, force=True)  # Force register on initial press

    def on_release(self, event):
        self.mouse_pressed = False
        self.last_button_hit = None  # Reset button tracking on release
        if self.repeat_job:
            self.after_cancel(self.repeat_job)
            self.repeat_job = None
        self.active_key = None
        # Clear tracking for this widget release
        try:
            widget_id = id(event.widget)
            with self.input_lock:
                # Mark button as released (allows next press to register)
                self.button_pressed_state[widget_id] = False
                
                self.pressed_buttons.pop(widget_id, None)
                # Cancel pending feedback job if exists
                if widget_id in self.button_states:
                    old_job = self.button_states[widget_id].get("feedback_job")
                    if old_job:
                        try:
                            self.after_cancel(old_job)
                        except:
                            pass
                    # Force reset to original color on release
                    orig_color = self.button_states[widget_id].get("orig_color")
                    try:
                        if event.widget.winfo_exists() and orig_color:
                            event.widget.config(bg=orig_color)
                    except:
                        pass
                    del self.button_states[widget_id]
        except:
            pass

    def on_motion(self, event):
        if not self.mouse_pressed:
            return
        
        current_time = time.time() * 1000
        # Use tight throttle for faster response
        if current_time - self.last_motion_event < EVENT_THROTTLE_MS:
            if not self.motion_pending:
                self.motion_pending = True
                self.after(1, lambda: self._process_motion(event))  # Minimal delay
            return
        
        self.last_motion_event = current_time
        self._process_motion(event)
    
    def _process_motion(self, event):
        """Process motion with minimal overhead - direct widget lookup."""
        self.motion_pending = False
        # Fast widget lookup at pointer position
        x, y = event.x_root, event.y_root
        w = self.winfo_containing(x, y)
        
        if not w:
            return
        
        # Traverse up to find button efficiently
        widget = None
        current = w
        while current and current != self.grid_frame:
            if isinstance(current, tk.Button) and hasattr(current, "meta_key"):
                widget = current
                break
            current = current.master
        
        if widget and hasattr(widget, "meta_key"):
            # Check if we switched buttons
            widget_id = id(widget)
            if widget_id != self.last_button_hit:
                self.last_button_hit = widget_id
                self.check_input_hit(widget, force=True)
    
    def _get_widgets_at_pointer(self):
        """Get widget(s) at current pointer position(s)."""
        x, y = self.winfo_pointerx(), self.winfo_pointery()
        w = self.winfo_containing(x, y)
        if not w:
            return []
        
        # Traverse up to find button
        result = []
        temp = w
        while temp and not (isinstance(temp, tk.Button) and hasattr(temp, "meta_key")):
            temp = temp.master
            if temp == self:
                break
        
        if isinstance(temp, tk.Button) and hasattr(temp, "meta_key"):
            result.append(temp)
        
        return result

    def check_input_hit(self, widget, force=False):
        key = widget.meta_key
        if not key:
            return
        
        widget_id = id(widget)
        current_time = time.time() * 1000  # milliseconds
        
        with self.input_lock:
            # Get button state - track if THIS button is currently pressed
            was_pressed = self.button_pressed_state.get(widget_id, False)
            
            # If not currently pressed, always allow new press (different fingers on different buttons)
            if not was_pressed:
                self.button_pressed_state[widget_id] = True
                self.pressed_buttons[widget_id] = (key, current_time, True)
                self.active_keys[key] = current_time
                
                # Cancel any pending feedback for this button
                if widget_id in self.button_states:
                    old_job = self.button_states[widget_id].get("feedback_job")
                    if old_job:
                        try:
                            self.after_cancel(old_job)
                        except:
                            pass
            else:
                # Button already pressed - check debounce for repeated presses on SAME button
                if not force:
                    last_press_time = self.pressed_buttons.get(widget_id, (key, 0, False))[1]
                    time_since_last = current_time - last_press_time
                    
                    # Only skip if debounce is enabled AND threshold not met
                    if self.key_press_threshold > 0 and time_since_last < self.key_press_threshold:
                        return  # Debounce active - ignore repeat press
                
                # Update timestamp for this repeated press
                self.pressed_buttons[widget_id] = (key, current_time, True)
                self.active_keys[key] = current_time
        
        # Send key immediately on press
        self.send_key(key)
        
        # Visual feedback ONLY if enabled (completely skip if disabled)
        if not self.visual_feedback_enabled:
            return  # Early exit - no visual feedback at all
        
        try:
            if not widget.winfo_exists():
                return
            orig = widget.cget("bg")
            widget.config(bg="#aaaaaa")
            # Schedule reset
            feedback_job = self.after(50, lambda: self._reset_widget_color_tracked(widget, orig, widget_id))
            with self.input_lock:
                self.button_states[widget_id] = {"feedback_job": feedback_job, "orig_color": orig}
        except tk.TclError:
            pass
    
    def _reset_widget_color_tracked(self, widget, orig_color, widget_id):
        """Reset widget color and clear state tracking."""
        try:
            if widget.winfo_exists():
                widget.config(bg=orig_color)
        except tk.TclError:
            pass
        finally:
            with self.input_lock:
                if widget_id in self.button_states:
                    del self.button_states[widget_id]
    
    def _reset_widget_color(self, widget, orig_color):
        """Safely reset widget color with existence check."""
        try:
            if widget.winfo_exists():
                widget.config(bg=orig_color)
        except tk.TclError:
            pass

    def repeat_loop(self, key):
        """Deprecated - keys are now sent on press instead of repeat loop."""
        pass

    def send_key(self, key):
        # DirectInput Handling - SEND ON PRESS IMMEDIATELY
        key_lower = key.lower().strip()
        if key_lower in SCAN_CODES:
            code = SCAN_CODES[key_lower]
            PressKey(code)
            # Release asynchronously to not block input thread
            threading.Thread(target=self._release_key_async, args=(code,), daemon=True).start()
        else:
            # Fallback for keys not in scan list
            try: keyboard.write(key)
            except: pass
    
    def _release_key_async(self, code):
        """Release key asynchronously with minimal delay."""
        time.sleep(KEY_PRESS_DELAY / 1000.0)
        try:
            ReleaseKey(code)
        except:
            pass

    def play_key_pulse(self, key_str, btn):
        if not key_str: return
        orig = btn.cget('bg')
        btn.configure(bg="#aaaaaa")
        self.after(100, lambda: btn.configure(bg=orig))
        self.send_key(key_str)

    def finish_key_edit(self, r, c, val):
        self.grid_data[(r,c)]['key'] = val
        if (r,c) in self.button_refs:
            btn = self.button_refs[(r,c)]
            if btn.winfo_exists():
                btn.configure(text=val)
                btn.meta_key = val 

    def finish_col_resize(self, c, val): pass
    def finish_row_resize(self, r, val): pass

    def start_drag_select(self, r, c):
        if (r,c) in self.selected_cells:
            self.drag_start_state = "deselect"
            self.selected_cells.remove((r,c))
        else:
            self.drag_start_state = "select"
            self.selected_cells.add((r,c))
        self.update_visuals()

    def do_drag_select(self, event):
        x, y = self.grid_frame.winfo_pointerx(), self.grid_frame.winfo_pointery()
        widget = self.grid_frame.winfo_containing(x, y)
        if widget and isinstance(widget, tk.Button) and hasattr(widget, 'grid_pos'):
            r, c = widget.grid_pos
            if self.drag_start_state == "select": 
                if (r,c) not in self.selected_cells:
                    self.selected_cells.add((r,c)); self.update_visuals()
            elif self.drag_start_state == "deselect":
                if (r,c) in self.selected_cells:
                    self.selected_cells.remove((r,c)); self.update_visuals()

    def end_drag_select(self): self.drag_start_state = None

    def toggle_mode(self):
        if self.mode == "design": self.enter_play_mode()
        else: self.enter_design_mode()

    def enter_play_mode(self):
        self.safe_commit_entry()
        self.design_geometry = self.geometry()
        self.update_idletasks()
        rx, ry = self.winfo_x(), self.winfo_y()
        y_offset = self.title_bar.winfo_height()
        self.mode = "play"
        self.panel.set_mode("play")
        self.title_bar.pack_forget() 
        self.attributes("-topmost", True)
        self.apply_click_through_style()
        self.refresh_grid()
        self.fit_window_to_content()
        self.geometry(f"+{rx}+{ry + y_offset}")

    def enter_design_mode(self):
        self.mode = "design"
        self.panel.set_mode("design")
        self.attributes("-topmost", False)
        hwnd = win32gui.GetParent(self.winfo_id()) or self.winfo_id()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        style = style & ~win32con.WS_EX_NOACTIVATE & ~win32con.WS_EX_TOPMOST
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
        # Unpack and repack to ensure titlebar is visible
        self.title_bar.pack_forget()
        self.title_bar.pack(fill="x", side="top", before=self.grid_frame)
        self.refresh_grid()
        self.fit_window_to_content()
        if self.design_geometry: self.geometry(self.design_geometry)

    def apply_click_through_style(self):
        hwnd = win32gui.GetParent(self.winfo_id()) or self.winfo_id()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        style = style | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)

    def open_theme_menu(self):
        m = tk.Menu(self, tearoff=0)
        for n in THEMES: m.add_command(label=n, command=lambda name=n: self.apply_theme(name))
        m.add_separator()
        m.add_command(label="Follow System", command=lambda: self.apply_theme("Follow System"))
        m.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    def pick_accent_color(self):
        color = colorchooser.askcolor(title="Choose Accent Color")[1]
        if color:
            self.current_theme['accent'] = color
            self.current_theme['btn_hover'] = color 
            self.current_theme['slider_active'] = color
            self.apply_theme_to_ui()

    def apply_theme(self, name):
        if name == "Follow System":
            mode = get_system_mode() 
            accent = get_system_accent()
            base_theme = THEMES["Modern Dark"] if mode == "Dark" else THEMES["Modern Light"]
            self.current_theme = base_theme.copy()
            self.current_theme["accent"] = accent
            self.current_theme["btn_hover"] = accent
            self.current_theme["slider_active"] = accent
        else:
            self.current_theme = THEMES[name].copy()
        self.apply_theme_to_ui()

    def apply_theme_to_ui(self):
        t = self.current_theme
        self.configure(bg=t['bg'])
        self.main_frame.configure(bg=t['bg'])
        self.grid_frame.configure(bg=t['grid_bg']) 
        self.title_bar.configure(bg=t['accent'], fg="#ffffff")
        self.panel.update_theme(t)
        self.refresh_grid()

    def update_opacity(self, val):
        v = float(val)
        self.attributes("-alpha", v)
        self.panel.attributes("-alpha", v)

    def get_context_indices(self):
        if not self.selected_cells: return len(self.row_sizes)-1, len(self.col_sizes)-1
        rows = [r for r,c in self.selected_cells]; cols = [c for r,c in self.selected_cells]
        return max(rows), max(cols)

    def add_row_at_selection(self):
        r_idx, _ = self.get_context_indices()
        new_row_idx = r_idx + 1
        self.row_sizes.insert(new_row_idx, 40)
        new_data = {}
        for (r, c), v in self.grid_data.items():
            if r >= new_row_idx: new_data[(r + 1, c)] = v
            else: new_data[(r, c)] = v
        for c in range(len(self.col_sizes)):
            new_data[(new_row_idx, c)] = {'key': '', 'span_r': 1, 'span_c': 1}
        self.grid_data = new_data
        self.refresh_grid(); self.fit_window_to_content()

    def add_col_at_selection(self):
        _, c_idx = self.get_context_indices()
        new_col_idx = c_idx + 1
        self.col_sizes.insert(new_col_idx, 60)
        new_data = {}
        for (r, c), v in self.grid_data.items():
            if c >= new_col_idx: new_data[(r, c + 1)] = v
            else: new_data[(r, c)] = v
        for r in range(len(self.row_sizes)):
            new_data[(r, new_col_idx)] = {'key': '', 'span_r': 1, 'span_c': 1}
        self.grid_data = new_data
        self.refresh_grid(); self.fit_window_to_content()

    def del_row_at_selection(self):
        if len(self.row_sizes) <= 1: return
        r_idx, _ = self.get_context_indices()
        self.row_sizes.pop(r_idx)
        new_data = {}
        for (r, c), v in self.grid_data.items():
            span_r = v.get('span_r', 1)
            if r == r_idx: pass
            elif r > r_idx: new_data[(r - 1, c)] = v
            else:
                if r + span_r - 1 >= r_idx: v['span_r'] = max(1, span_r - 1)
                new_data[(r, c)] = v
        self.grid_data = new_data
        self.selected_cells.clear(); self.refresh_grid(); self.fit_window_to_content()

    def del_col_at_selection(self):
        if len(self.col_sizes) <= 1: return
        _, c_idx = self.get_context_indices()
        self.col_sizes.pop(c_idx)
        new_data = {}
        for (r, c), v in self.grid_data.items():
            span_c = v.get('span_c', 1)
            if c == c_idx: pass
            elif c > c_idx: new_data[(r, c - 1)] = v
            else:
                if c + span_c - 1 >= c_idx: v['span_c'] = max(1, span_c - 1)
                new_data[(r, c)] = v
        self.grid_data = new_data
        self.selected_cells.clear(); self.refresh_grid(); self.fit_window_to_content()

    def merge_cells(self):
        if not self.selected_cells: return
        min_r, min_c = float('inf'), float('inf')
        max_r, max_c = float('-inf'), float('-inf')
        for r, c in self.selected_cells:
            cell = self.grid_data.get((r,c), {})
            sr = cell.get('span_r', 1); sc = cell.get('span_c', 1)
            min_r, min_c = min(min_r, r), min(min_c, c)
            max_r, max_c = max(max_r, r + sr - 1), max(max_c, c + sc - 1)
        min_r, min_c = int(min_r), int(min_c)
        max_r, max_c = int(max_r), int(max_c)
        base = (min_r, min_c)
        self.grid_data[base].update({'span_r': max_r - min_r + 1, 'span_c': max_c - min_c + 1})
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                if (r,c) != base: self.grid_data[(r,c)].update({'span_r': 1, 'span_c': 1})
        self.selected_cells.clear(); self.refresh_grid()

    def unmerge_cells(self):
        for r, c in self.selected_cells: self.grid_data[(r,c)].update({'span_r': 1, 'span_c': 1})
        self.selected_cells.clear(); self.refresh_grid()

    # ===== Multi-Touch Optimization Methods =====
    def register_touch(self, widget):
        """Register a new touch (for multi-finger support)."""
        touch_id = self.next_touch_id
        self.next_touch_id += 1
        key = widget.meta_key if hasattr(widget, 'meta_key') else ''
        self.active_touches[touch_id] = {
            'widget': widget,
            'key': key,
            'active': True,
            'time': time.time()
        }
        return touch_id
    
    def unregister_touch(self, touch_id):
        """Unregister a touch."""
        if touch_id in self.active_touches:
            del self.active_touches[touch_id]
    
    def get_touch_at_position(self, x_root, y_root):
        """Find widget at given position efficiently."""
        w = self.winfo_containing(x_root, y_root)
        if not w: return None
        while w and not (isinstance(w, tk.Button) and hasattr(w, "meta_key")):
            w = w.master
            if w == self: break
        return w if isinstance(w, tk.Button) and hasattr(w, "meta_key") else None
    
    def batch_render_update(self):
        """Batch visual updates to reduce latency."""
        if self.render_batch_job:
            self.after_cancel(self.render_batch_job)
        self.pending_render = True
        self.render_batch_job = self.after(RENDER_BATCH_INTERVAL, self._flush_render)
    
    def _flush_render(self):
        """Execute pending render updates."""
        if self.pending_render:
            self.update_visuals()
            self.pending_render = False
        self.render_batch_job = None

    def save_layout(self):
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if f: 
            with open(f, 'w') as o: json.dump({"row_sizes": self.row_sizes, "col_sizes": self.col_sizes, "cells": {str(k): v for k, v in self.grid_data.items()}}, o)
    
    def load_layout(self):
        f = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if f: 
            with open(f, 'r') as i:
                d = json.load(i)
                self.row_sizes, self.col_sizes = d["row_sizes"], d["col_sizes"]
                self.grid_data = {eval(k): v for k, v in d["cells"].items()}
                self.refresh_grid(); self.fit_window_to_content()

if __name__ == "__main__":
    app = VirtualKeyboardApp()
    app.mainloop()