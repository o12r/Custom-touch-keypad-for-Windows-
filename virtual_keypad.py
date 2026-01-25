import tkinter as tk
from tkinter import filedialog, colorchooser
import keyboard
import json
import win32gui
import win32con
import math
import winreg

# --- Configuration ---
DEFAULT_ROWS = 4
DEFAULT_COLS = 5
BASE_UNIT = 10 
REPEAT_INTERVAL = 50   # Speed of rapid fire (ms)
DRAG_THRESHOLD = 5     # Pixels to move before counting as a drag

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
    except Exception:
        return "#0078d7"

def get_system_mode():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "Light" if value == 1 else "Dark"
    except Exception:
        return "Light"

THEMES = {
    "Modern Light": {
        "bg": "#f9f9f9", "fg": "#000000", 
        "grid_bg": "#d0d0d0", 
        "btn_bg": "#ffffff", "btn_fg": "#000000", "btn_hover": "#0078d7", "btn_active": "#e0e0e0",
        "header_bg": "#f9f9f9", "header_fg": "#000000",
        "accent": "#0078d7", "relief": "flat", "border": 0, "gap": 1,
        "separator": "#e0e0e0",
        "slider_bg": "#f9f9f9", "slider_trough": "#e0e0e0", "slider_active": "#0078d7"
    },
    "Modern Dark": {
        "bg": "#202020", "fg": "#ffffff", 
        "grid_bg": "#3a3a3a", 
        "btn_bg": "#2d2d2d", "btn_fg": "#ffffff", "btn_hover": "#007acc", "btn_active": "#404040",
        "header_bg": "#202020", "header_fg": "#aaaaaa", 
        "accent": "#007acc", "relief": "flat", "border": 0, "gap": 1,
        "separator": "#3a3a3a",
        "slider_bg": "#202020", "slider_trough": "#404040", "slider_active": "#007acc"
    },
    "Retro Gray": {
        "bg": "#c0c0c0", "fg": "#000000", 
        "grid_bg": "#c0c0c0",
        "btn_bg": "#c0c0c0", "btn_fg": "#000000", "btn_hover": "#000080", "btn_active": "#aaaaaa",
        "header_bg": "#c0c0c0", "header_fg": "#000000", 
        "accent": "#000080", "relief": "raised", "border": 2, "gap": 0,
        "separator": "#808080",
        "slider_bg": "#c0c0c0", "slider_trough": "#808080", "slider_active": "#000080"
    },
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
        
        # --- UI Structure ---
        self.frame_full = tk.Frame(self)
        
        # Custom Title Bar
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

    def toggle_rapid(self):
        self.app.rapid_mode = not self.app.rapid_mode
        txt = "Mode: SLIDE 〰️" if self.app.rapid_mode else "Mode: TYPE ⌨️"
        self.btn_slide.configure(text=txt)
        self.app.update_input_bindings()

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
        
        # --- State ---
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
        
        # Interaction State
        self.mouse_pressed = False
        self.active_key = None
        self.repeat_job = None
        self.drag_start_pos = None 
        self.paint_value = None 
        self.interaction_type = None

        # Initialize Data
        for r in range(DEFAULT_ROWS):
            for c in range(DEFAULT_COLS):
                self.grid_data[(r, c)] = {'key': '', 'span_r': 1, 'span_c': 1}

        # --- Main Window Layout ---
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

    # --- Window Dragging ---
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
        
        # Design Headers
        if self.mode == "design":
            self.grid_frame.columnconfigure(0, minsize=30, weight=0) 
            self.grid_frame.rowconfigure(0, minsize=20, weight=0)    
            tk.Label(self.grid_frame, bg=t['bg']).grid(row=0, column=0, sticky="nsew")

            # Column Headers
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

            # Row Headers
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

        # Draw Cells
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

    # --- Header Interaction Logic ---
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
                                self.col_sizes[target_idx] = new_size
                                changed = True
                            elif axis == 'row' and self.row_sizes[target_idx] != new_size:
                                self.row_sizes[target_idx] = new_size
                                changed = True
                            
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
            self.refresh_grid()
            self.fit_window_to_content()
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
                            self.selected_cols_indices.add(c)
                            self.refresh_grid()
                    elif axis == 'row' and int(info['column']) == 0:
                        r = int(info['row']) - offset
                        if 0 <= r < len(self.row_sizes):
                            self.selected_rows_indices.add(r)
                            self.refresh_grid()

    def finish_header_resize(self, axis, idx, val):
        try:
            new_size = int(float(val) * BASE_UNIT)
            indices = self.selected_cols_indices if axis == 'col' else self.selected_rows_indices
            target_list = self.col_sizes if axis == 'col' else self.row_sizes
            if idx in indices:
                for i in indices:
                    if 0 <= i < len(target_list): target_list[i] = new_size
            else:
                target_list[idx] = new_size
            self.refresh_grid()
            self.fit_window_to_content()
        except: pass

    # --- Cell Paint Interaction ---
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
            if context:
                self.goto_next_editable(context)
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
            next_c = c + 1
            next_r = r
            if next_c >= len(self.col_sizes):
                next_c = 0
                next_r += 1
            
            if next_r < len(self.row_sizes):
                if (next_r, next_c) in self.button_refs:
                    btn = self.button_refs[(next_r, next_c)]
                    val = self.grid_data[(next_r, next_c)].get('key','')
                    self.start_inline_edit(btn, val, lambda v: self.finish_key_edit(next_r, next_c, v), {'type':'cell', 'r':next_r, 'c':next_c})

    def on_btn_hover(self, btn, hovering):
        r, c = btn.grid_pos
        if (r,c) in self.selected_cells: return
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
        if not isinstance(w, tk.Button) or not hasattr(w, "meta_key"): return
        self.mouse_pressed = True
        self.check_input_hit(w)

    def on_release(self, event):
        self.mouse_pressed = False
        if self.repeat_job:
            self.after_cancel(self.repeat_job)
            self.repeat_job = None
        self.active_key = None

    def on_motion(self, event):
        if self.mouse_pressed:
            w = self.winfo_containing(event.x_root, event.y_root)
            if not w: return 
            while w and not (isinstance(w, tk.Button) and hasattr(w, "meta_key")):
                w = w.master
                if w == self: break
            if isinstance(w, tk.Button) and hasattr(w, "meta_key"): self.check_input_hit(w)

    def check_input_hit(self, widget):
        key = widget.meta_key
        if key and key != self.active_key:
            if self.repeat_job: self.after_cancel(self.repeat_job)
            self.active_key = key
            orig = widget.cget("bg")
            widget.config(bg="#aaaaaa")
            self.after(100, lambda: widget.config(bg=orig))
            self.send_key(key)
            self.repeat_job = self.after(REPEAT_INTERVAL, lambda: self.repeat_loop(key))

    def repeat_loop(self, key):
        if self.mouse_pressed and self.active_key == key:
            self.send_key(key)
            self.repeat_job = self.after(REPEAT_INTERVAL, lambda: self.repeat_loop(key))

    def send_key(self, key):
        try: keyboard.write(key)
        except: pass

    def play_key_pulse(self, key_str, btn):
        if not key_str: return
        orig = btn.cget('bg')
        btn.configure(bg="#aaaaaa")
        self.after(100, lambda: btn.configure(bg=orig))
        try: keyboard.send(key_str)
        except: pass

    def finish_key_edit(self, r, c, val):
        self.grid_data[(r,c)]['key'] = val
        if (r,c) in self.button_refs:
            btn = self.button_refs[(r,c)]
            if btn.winfo_exists():
                btn.configure(text=val)
                btn.meta_key = val 

    def finish_col_resize(self, c, val):
        try:
            self.col_sizes[c] = int(float(val) * BASE_UNIT)
            self.refresh_grid(); self.fit_window_to_content()
        except: pass
    def finish_row_resize(self, r, val):
        try:
            self.row_sizes[r] = int(float(val) * BASE_UNIT)
            self.refresh_grid(); self.fit_window_to_content()
        except: pass

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
                    self.selected_cells.add((r,c))
                    self.update_visuals()
            elif self.drag_start_state == "deselect":
                if (r,c) in self.selected_cells:
                    self.selected_cells.remove((r,c))
                    self.update_visuals()

    def end_drag_select(self): self.drag_start_state = None
    
    def toggle_row_select(self, r): pass 
    def toggle_col_select(self, c): pass 

    def toggle_mode(self):
        if self.mode == "design": self.enter_play_mode()
        else: self.enter_design_mode()

    def enter_play_mode(self):
        self.safe_commit_entry()
        self.design_geometry = self.geometry()
        self.update_idletasks()
        rx, ry = self.winfo_x(), self.winfo_y()
        
        # Calculate offset to keep grid stationary
        y_offset = self.title_bar.winfo_height()
        
        self.mode = "play"
        self.panel.set_mode("play")
        self.title_bar.pack_forget() 
        self.attributes("-topmost", True)
        self.apply_click_through_style()
        self.refresh_grid()
        self.fit_window_to_content()
        
        # Apply offset to geometry
        self.geometry(f"+{rx}+{ry + y_offset}")

    def enter_design_mode(self):
        self.mode = "design"
        self.panel.set_mode("design")
        self.title_bar.pack(fill="x", side="top") 
        self.attributes("-topmost", False)
        hwnd = win32gui.GetParent(self.winfo_id()) or self.winfo_id()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        style = style & ~win32con.WS_EX_NOACTIVATE & ~win32con.WS_EX_TOPMOST
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
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

    # --- ROW/COL MANIPULATION ---
    def get_context_indices(self):
        if not self.selected_cells:
            return len(self.row_sizes)-1, len(self.col_sizes)-1
        rows = [r for r,c in self.selected_cells]
        cols = [c for r,c in self.selected_cells]
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
        self.selected_cells.clear()
        self.refresh_grid(); self.fit_window_to_content()

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
        self.selected_cells.clear()
        self.refresh_grid(); self.fit_window_to_content()

    def merge_cells(self):
        if not self.selected_cells: return
        min_r, min_c = float('inf'), float('inf')
        max_r, max_c = float('-inf'), float('-inf')
        for r, c in self.selected_cells:
            cell = self.grid_data.get((r,c), {})
            sr = cell.get('span_r', 1)
            sc = cell.get('span_c', 1)
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