import json
import os
import tkinter as tk
from tkinter import messagebox, simpledialog
from pywinauto import Desktop, Application
import pyperclip
import re
import datetime
import pyautogui
import winsound
import ctypes
import ctypes.wintypes
import win32gui
import win32con
import threading
import time
try:
    import hid
except ImportError:
    hid = None

try:
    import keyboard
except ImportError:
    keyboard = None

import subprocess
import sys

import csv
import io
import win32api

# Optional Features Dependencies
pil_lib = False
try:
    from PIL import Image, ImageGrab, ImageTk
    pil_lib = True
except ImportError: pass

np_lib = False
try:
    import numpy as np
    np_lib = True
except ImportError: pass

cv2_lib = False
try:
    import cv2
    cv2_lib = True
except ImportError:
    cv2_lib = False

# OCR Engine Import
rapid_ocr_lib = None
try:
    from rapidocr_onnxruntime import RapidOCR
    rapid_ocr_lib = RapidOCR
except ImportError:
    pass

def log_trace(msg):
    try:
        with open("mosaic_setup_trace.txt", "a") as f:
            f.write(f"{datetime.datetime.now()}: {msg}\n")
    except: pass

def play_beep(frequency, duration_ms, volume=0.04):
    """Play a beep with volume control (0.0 to 1.0).
    Falls back to winsound.Beep if numpy is unavailable.
    """
    import wave
    import struct
    
    if not np_lib:
        # Fallback: no volume control
        winsound.Beep(frequency, duration_ms)
        return
    
    sample_rate = 44100
    n_samples = int(sample_rate * duration_ms / 1000)
    t = np.linspace(0, duration_ms / 1000, n_samples, False)
    
    # Generate sine wave at controlled amplitude
    amplitude = int(32767 * volume)
    samples = (amplitude * np.sin(2 * np.pi * frequency * t)).astype(np.int16)
    
    # Create in-memory WAV
    buf = io.BytesIO()
    with wave.open(buf, 'wb') as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(sample_rate)
        wav.writeframes(samples.tobytes())
    
    buf.seek(0)
    winsound.PlaySound(buf.read(), winsound.SND_MEMORY)

def kill_other_instances():
    """Nuclear option: Kill all OTHER python processes running MosaicTools."""
    try:
        print("Checking for duplicate instances...")
        current_pid = os.getpid()
        # Use tasklist instead of wmic (deprecated/flakey)
        # We check common python executables
        exes = ['python.exe', 'py.exe', 'pythonw.exe', 'pyw.exe']
        
        # Get all tasks in CSV format including command line? 
        # tasklist doesn't give commandstack easily without /V.
        # But wmic is the only standard way to get commandline.
        # Let's try wmic again but with VERY strict parsing and broad safety.
        
        cmd = 'wmic process where "name like \'%python%\' or name like \'%py%\'" get commandline,processid /format:csv'
        
        # CREATE_NO_WINDOW = 0x08000000
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, startupinfo=si)
        out, err = proc.communicate(timeout=5)
        
        lines = out.decode('utf-8', errors='ignore').strip().splitlines()
        
        # Parse CSV output from wmic
        # Node,CommandLine,ProcessId
        header_found = False
        for line in lines:
            if not line.strip(): continue
            if "CommandLine" in line and "ProcessId" in line:
                header_found = True
                continue
                
            parts = line.split(',')
            if len(parts) < 2: continue
            
            # Last element is PID, rest is command line (handles commas in path)
            pid_str = parts[-1].strip()
            cmd_line = ",".join(parts[:-1]) # Rejoin command line parts
            
            if not pid_str.isdigit(): continue
            
            pid = int(pid_str)
            
            if pid == current_pid:
                continue
                
            if "MosaicTools" in cmd_line:
                print(f"Found duplicate PID {pid}. Terminating...")
                subprocess.call(['taskkill', '/F', '/PID', str(pid)], 
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=si)
                
    except Exception as e:
        print(f"Startup cleanup failed (non-critical): {e}")

# Nuance / Dictaphone Vendor IDs
VENDOR_IDS = [0x0554, 0x0558]

# Button Definitions
BUTTON_DEFINITIONS = {
    "Left Button": [0x00, 0x80, 0x00],
    "Right Button": [0x00, 0x00, 0x02],
    "T Button": [0x00, 0x01, 0x00],
    "Record Button": [0x00, 0x04, 0x00],
    "Skip Back": [0x00, 0x02, 0x00],
    "Skip Forward": [0x00, 0x08, 0x00],
    "Rewind": [0x00, 0x10, 0x00],
    "Fast Forward": [0x00, 0x20, 0x00],
    "Stop/Play": [0x00, 0x40, 0x00],
    "Checkmark": [0x00, 0x00, 0x01]
}

# Available Actions
ACTION_NONE = "None"
ACTION_BEEP = "System Beep"
ACTION_GET_PRIOR = "Get Prior"
ACTION_SCRAPE = "Critical Findings"
ACTION_DEBUG = "Debug Scrape"
ACTION_SHOW_REPORT = "Show Report"
ACTION_CAPTURE_SERIES = "Capture Series/Image"
ACTION_TOGGLE_RECORD = "Start/Stop Recording"
ACTION_PROCESS_REPORT = "Process Report"
ACTION_SIGN_REPORT = "Sign Report"

AVAILABLE_ACTIONS = [
    ACTION_NONE,
    ACTION_BEEP,
    ACTION_GET_PRIOR,
    ACTION_SCRAPE,
    ACTION_DEBUG,
    ACTION_SHOW_REPORT,
    ACTION_CAPTURE_SERIES,
    ACTION_TOGGLE_RECORD,
    ACTION_PROCESS_REPORT,
    ACTION_SIGN_REPORT
]

DEFAULT_MAPPINGS = {
    "Left Button": ACTION_GET_PRIOR,
    "Right Button": ACTION_SCRAPE,
    "T Button": ACTION_DEBUG,
    "Record Button": ACTION_BEEP,
    "Skip Back": ACTION_NONE,
    "Skip Forward": ACTION_NONE,
    "Rewind": ACTION_NONE,
    "Fast Forward": ACTION_NONE,
    "Stop/Play": ACTION_NONE,
    "Checkmark": ACTION_NONE
}

# Windows Messages for AHK Communication
WM_TRIGGER_SCRAPE = 0x0401
WM_TRIGGER_DEBUG = 0x0402
WM_TRIGGER_BEEP = 0x0403
WM_TRIGGER_SHOW_REPORT = 0x0404
WM_TRIGGER_CAPTURE_SERIES = 0x0405
WM_TRIGGER_GET_PRIOR = 0x0406
WM_TRIGGER_TOGGLE_RECORD = 0x0407
WM_TRIGGER_PROCESS_REPORT = 0x0408
WM_TRIGGER_SIGN_REPORT = 0x0409
# =============================================================================
# FLOATING TOOLBAR LOGIC
# =============================================================================
class FloatingToolbarWindow:
    def __init__(self, master, settings, save_callback):
        self.master = master
        self.settings = settings
        self.save_callback = save_callback
        
        self.window = tk.Toplevel(master)
        self.window.withdraw() # Hide immediately
        self.window.title("Floating Buttons")
        self.window.config(bg='black')
        self.window.attributes('-topmost', True)
        self.window.overrideredirect(True)
        
        # Geometry
        width = 120
        height = 175
        x = settings.get("floating_toolbar_x", 100)
        y = settings.get("floating_toolbar_y", 620)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.deiconify() # Show at correct position
        
        # Save on drag end
        self.window.bind("<ButtonRelease-1>", self.on_release)
        
        # UI
        self.setup_ui()
        
    def setup_ui(self):
        self.frame = tk.Frame(self.window, bg='black')
        self.frame.pack(expand=True, fill='both')
        
        # Grid
        for i in range(4): self.frame.grid_rowconfigure(i, weight=1 if i>0 else 0)
        for i in range(2): self.frame.grid_columnconfigure(i, weight=1)
        
        # Drag Bar
        drag_bar = tk.Frame(self.frame, bg='black', height=15)
        drag_bar.grid(row=0, column=0, columnspan=2, sticky='ew', padx=2, pady=2)
        drag_bar.grid_propagate(False)
        
        lbl = tk.Label(drag_bar, text="⋯", font=("Segoe UI", 10), bg='black', fg='#666')
        lbl.pack(side=tk.RIGHT, padx=2)
        
        drag_bar.bind("<Button-1>", self.start_drag)
        drag_bar.bind("<B1-Motion>", self.on_drag)
        lbl.bind("<Button-1>", self.start_drag)
        lbl.bind("<B1-Motion>", self.on_drag)
        
        # Buttons
        options = {'bg':'black', 'fg':'#CCCCCC', 'font':('Segoe UI Symbol', 16, 'bold'), 
                   'relief':'flat', 'activebackground':'#333', 'activeforeground':'#CCC'}
        
        # Rows 1-2
        specs = [
            [("↕", ('ctrl', 'v')), ("↔", ('ctrl', 'h'))],
            [("↺", (',')), ("↻", ('.'))]
        ]
        
        for r, row in enumerate(specs):
            for c, (text, keys) in enumerate(row):
                f = tk.Frame(self.frame, bg='#CCCCCC')
                f.grid(row=r+1, column=c, padx=1, pady=1, sticky='nsew')
                btn = tk.Button(f, text=text, command=lambda k=keys: self.send_iv_keys(k), **options)
                btn.pack(expand=True, fill='both', padx=1, pady=1)
                
        # Zoom Out (Row 3)
        f = tk.Frame(self.frame, bg='#CCCCCC')
        f.grid(row=3, column=0, columnspan=2, padx=1, pady=1, sticky='nsew')
        z_opt = options.copy()
        z_opt['font'] = ('Segoe UI', 11, 'bold')
        btn = tk.Button(f, text="Zoom Out", command=lambda: self.send_iv_keys('-'), **z_opt)
        btn.pack(expand=True, fill='both', padx=1, pady=1)
        
        self._drag_start_x = 0
        self._drag_start_y = 0

    def start_drag(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def on_drag(self, event):
        x = self.window.winfo_x() + event.x - self._drag_start_x
        y = self.window.winfo_y() + event.y - self._drag_start_y
        self.window.geometry(f"+{x}+{y}")
        
    def on_release(self, event):
        self.settings["floating_toolbar_x"] = self.window.winfo_x()
        self.settings["floating_toolbar_y"] = self.window.winfo_y()
        self.save_callback()

    def send_iv_keys(self, keys):
        """Send keys to InteleViewer."""
        import win32gui, win32con
        def find_iv(hwnd, res):
            if win32gui.IsWindowVisible(hwnd) and "inteleviewer" in win32gui.GetWindowText(hwnd).lower():
                res.append(hwnd)
        
        wins = []
        win32gui.EnumWindows(find_iv, wins)
        if wins:
            iv = wins[0]
            try:
                win32gui.ShowWindow(iv, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(iv)
                time.sleep(0.05)
                if isinstance(keys, tuple): pyautogui.hotkey(*keys)
                else: pyautogui.press(keys)
            except Exception as e:
                 print(f"IV Send Error: {e}")

    def destroy(self):
        self.window.destroy()

# =============================================================================
# OCR LOGIC (From ReadSeriesImage)
# =============================================================================
# OCR LOGIC (From ReadSeriesImage)
# =============================================================================
ocr_import_error = None
try:
    from rapidocr_onnxruntime import RapidOCR
    rapid_ocr_lib = True
except ImportError as e:
    rapid_ocr_lib = False
    ocr_import_error = f"ImportError: {e}"
    try:
        log_trace(f"OCR Import Error: {e}")
    except: pass
except Exception as e:
    rapid_ocr_lib = False
    ocr_import_error = f"Import Crash: {e}"
    try:
        log_trace(f"OCR Import Crash: {e}")
    except: pass

class OCREngineWrapper:
    def __init__(self):
        self.rapid = None
        self.init_error = ocr_import_error
        
        if rapid_ocr_lib:
            try:
                self.rapid = RapidOCR()
                # log_trace("RapidOCR Initialized Successfully")
            except Exception as e:
                self.init_error = f"Init Failed: {e}"
                log_trace(f"RapidOCR Init Failed: {e}")
        else:
            if not self.init_error:
                self.init_error = "Library not imported."
            # log_trace("RapidOCR library not available.")

    def run_rapid(self, img_np):
        if not self.rapid: 
            log_trace(f"OCR Error: Engine not initialized. Reason: {self.init_error}")
            return None
        try:
            # log_trace(f"OCR Input Shape: {img_np.shape if hasattr(img_np, 'shape') else 'Invalid'}")
            result, _ = self.rapid(img_np)
            if result:
                return "\n".join([line[1] for line in result])
            return ""
        except Exception as e:
            log_trace(f"OCR Engine Crash: {e}")
            return None

    def find_yellow_target(self):
        """Find the yellow box target on screen."""
        if not cv2_lib: 
            print("CV2 not installed, yellow box detection disabled.")
            return None
        try:
            full_screenshot = ImageGrab.grab(all_screens=True)
            img_bgr = cv2.cvtColor(np.array(full_screenshot), cv2.COLOR_RGB2BGR)
            hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
            
            # Yellow mask
            lower = np.array([13, 150, 150])
            upper = np.array([33, 255, 255])
            mask = cv2.inRange(hsv, lower, upper)
            
            contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            best_rect = None
            max_area = 0
            for cnt in contours:
                x, y, cw, ch = cv2.boundingRect(cnt)
                if cw > 400 and ch > 400: # Match ReadSeriesImage threshold
                    area = cw * ch
                    if area > max_area:
                        max_area = area
                        best_rect = (x, y)
            
            if best_rect:
                bx, by = best_rect
                user32 = ctypes.windll.user32
                v_screen_x = user32.GetSystemMetrics(76)
                v_screen_y = user32.GetSystemMetrics(77)
                return v_screen_x + bx, v_screen_y + by
        except Exception as e:
            log_trace(f"Yellow box find error: {e}")
        return None

def extract_series_image_numbers(text):
    """Robust extraction logic for Series/Image (Synced with ReadSeriesImage.py).
    
    Logic Priorities:
    1. Same line match: "Se: 3 Im: 15"
    2. Adjacent line match: Line i-1 is treated as Series if line i is Image (Positional Trust).
       - Weak regex used on line i-1 if strict 'Se:' fails.
    3. Closest Series before Image (Fallback).
    """
    if not text: return None
    
    # Clean Typos
    text = re.sub(r"^(0e|5e|8e|Be|S8|S6|Se|Se:)", "Se:", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"^(Irm|1rn|1m|Lm|1n|In|ln|Im|Im:)", "Im:", text, flags=re.MULTILINE | re.IGNORECASE)
    
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    series_candidates = [] # list of (line_index, value)
    image_candidates = []  # list of (line_index, value)

    # Pass 1: Gather all strict match candidates
    for idx, line in enumerate(lines):
        # Priority 1: Combined pattern on same line
        combined = re.search(r"Se:\s*[^0-9]*(\d+).*?Im:\s*[^0-9]*(\d+)", line, re.IGNORECASE)
        if combined:
            return f"(series {combined.group(1)}, image {combined.group(2)})"

        # Strict Series Check
        m_se = re.search(r"Se:\s*[^0-9]*(\d+)", line, re.IGNORECASE)
        if m_se:
            series_candidates.append((idx, m_se.group(1)))
            
        # Image Check
        m_im = re.search(r"Im:\s*[^0-9]*(\d+)", line, re.IGNORECASE)
        if m_im:
            image_candidates.append((idx, m_im.group(1)))
        else:
            # Fallback X/Y format
            m_fall = re.search(r"\b(\d+)\s*/\s*\d+", line)
            if m_fall:
                image_candidates.append((idx, m_fall.group(1)))

    # Selection Logic
    final_series = None
    final_image = None
    
    if image_candidates:
        final_image_idx, final_image = image_candidates[0] # Focus on first detected image for now
        
        # Check specific line immediately above (index - 1)
        target_se_idx = final_image_idx - 1
        
        # 1. Is there a strict Series candidate at this exact line?
        matches = [val for idx, val in series_candidates if idx == target_se_idx]
        if matches:
            return f"(series {matches[0]}, image {final_image})"
            
        # 2. Positional Trust: If line exists, isn't 'Accession', and has a number -> Use it.
        # This handles garbled headers like "3e: 5" or just "# 5"
        if target_se_idx >= 0 and target_se_idx < len(lines):
            prev_line = lines[target_se_idx]
            
            # Guard: ensure checking line isn't Accession/Acc
            is_accession = re.search(r"(Acc|Accession)", prev_line, re.IGNORECASE)
            
            if not is_accession:
                # Weak extract: Look for ANY digit sequence
                # We skip "Se:" part since we already failed strict match, just look for number
                m_weak = re.search(r"[:#\s]+(\d+)", prev_line)
                if not m_weak:
                    # Try raw number at start or end
                    m_weak = re.search(r"\b(\d+)\b", prev_line)
                
                if m_weak:
                    return f"(series {m_weak.group(1)}, image {final_image})"

        # Priority 3: No adjacent pair, use closest preceding Series
        best_se = None
        for se_idx, se_val in series_candidates:
            if se_idx < final_image_idx:
                best_se = se_val
            else:
                break 
        final_series = best_se
        
    else:
        if series_candidates:
            final_series = series_candidates[-1][1]

    if final_series or final_image:
        return f"(series {final_series or '?'}, image {final_image or '?'})"
        
    return None

# =============================================================================
# MOSAIC REPORT LOGIC
# =============================================================================
def find_mosaic_report_window():
    """Find Mosaic Info Hub or Mosaic Reporting window."""
    desktop = Desktop(backend="uia")
    try:
        all_windows = desktop.windows(visible_only=True)
        for window in all_windows:
            try:
                window_text = window.window_text().lower()
                if "rvu counter" in window_text or "test" in window_text: continue
                
                if ("mosaic" in window_text and "info hub" in window_text) or \
                   ("mosaic" in window_text and "reporting" in window_text):
                    return window
            except: continue
    except: pass
    return None

def get_report_text_content(window):
    """Exhaustive search for report element text."""
    try:
        candidates = []
        keywords = ['TECHNIQUE', 'CLINICAL HISTORY', 'FINDINGS', 'IMPRESSION']
        
        for elem in window.descendants():
            try:
                cname = elem.element_info.class_name or ""
                if 'ProseMirror' in cname: # Tiptap editor
                    text = elem.get_win32_window_text() if hasattr(elem, 'get_win32_window_text') else (elem.window_text() or "")
                    if not text: continue
                    
                    score = sum(1 for kw in keywords if kw in text.upper())
                    candidates.append({'text': text, 'score': score})
            except: continue
            
        if not candidates: return None
        candidates.sort(key=lambda x: x['score'], reverse=True)
        return candidates[0]['text']
    except Exception as e:
        print(f"Report scan error: {e}")
        return None

def format_report_text(text):
    """Format the report text for the popup."""
    if not text: return "No report text found."
    text = text.replace('\ufffc', ' ')
    
    # Headers logic
    header_pattern = r'(?m)^(\s*[A-Z][A-Z0-9\s,/\.-]+:\s*)$'
    parts = re.split(header_pattern, text)
    formatted = []
    
    for part in parts:
        if not part.strip(): continue
        if re.match(header_pattern, part):
            formatted.append(f"\n\n{part.strip()}\n")
        else:
            cleaned = re.sub(r'\s+', ' ', part).strip()
            # List formatting: "1. Foo" -> "\n1. Foo"
            cleaned = re.sub(r'(\s)(\d+\.\s)', r'\n\2', cleaned)
            if cleaned: formatted.append(cleaned)
            
    text = "".join(formatted).strip()
    text = re.sub(r'\n{3,}', '\n\n', text)
    # Tighten headers
    text = re.sub(r'(:\n)\n+([A-Z])', r'\1\2', text)
    return text

# ============================================================================
# FLOATING INDICATOR WINDOW - Dictation Status Pill
# ============================================================================
class FloatingIndicatorWindow:
    """A small floating pill showing dictation state (mic icon)."""
    
    def __init__(self, parent, settings, save_callback):
        self.parent = parent
        self.settings = settings
        self.save_callback = save_callback
        
        self.win = tk.Toplevel(parent)
        self.win.title("Dictation Status")
        self.win.overrideredirect(True)
        self.win.attributes('-topmost', True)
        
        # Load Position
        x = self.settings.get("indicator_x", 100)
        y = self.settings.get("indicator_y", 550)
        self.win.geometry(f"50x50+{x}+{y}")
        
        # Visuals
        self.bg_off = "#444444"
        self.bg_on = "#CC0000"
        self.fg_color = "white"
        
        self.frame = tk.Frame(self.win, bg=self.bg_off, relief='raised', bd=2)
        self.frame.pack(fill='both', expand=True)
        
        self.icon_lbl = tk.Label(self.frame, text="🎙", font=("Segoe UI Symbol", 20), 
                                 bg=self.bg_off, fg=self.fg_color)
        self.icon_lbl.pack(expand=True, fill='both')
        
        # Drag Logic
        self.win.bind("<Button-1>", self._start_drag)
        self.win.bind("<B1-Motion>", self._on_drag)
        self.win.bind("<ButtonRelease-1>", self._on_release)
        self.icon_lbl.bind("<Button-1>", self._start_drag)
        self.icon_lbl.bind("<B1-Motion>", self._on_drag)
        self.icon_lbl.bind("<ButtonRelease-1>", self._on_release)
        
        self._drag_x = 0
        self._drag_y = 0

    def set_state(self, is_recording):
        """Update color based on state."""
        color = self.bg_on if is_recording else self.bg_off
        self.frame.config(bg=color)
        self.icon_lbl.config(bg=color)

    def _start_drag(self, event):
        self._drag_x = event.x
        self._drag_y = event.y

    def _on_drag(self, event):
        x = self.win.winfo_x() + event.x - self._drag_x
        y = self.win.winfo_y() + event.y - self._drag_y
        self.win.geometry(f"+{x}+{y}")

    def _on_release(self, event):
        self.settings["indicator_x"] = self.win.winfo_x()
        self.settings["indicator_y"] = self.win.winfo_y()
        self.save_callback()

    def destroy(self):
        self.win.destroy()

# ============================================================================
# FLOATING BUTTONS WINDOW - Dynamic InteleViewer Controls
# ============================================================================

class FloatingButtonsWindow:
    """A draggable floating toolbar with user-configurable buttons for InteleViewer."""
    
    def __init__(self, parent, settings, save_callback):
        """
        Args:
            parent: tk.Tk or tk.Toplevel parent
            settings: dict containing both 'floating_buttons' config and window position settings
            save_callback: function to call to persist settings
        """
        self.parent = parent
        self.settings = settings
        self.config = settings.get("floating_buttons", {})
        self.save_callback = save_callback
        
        self.win = tk.Toplevel(parent)
        self.win.title("InteleViewer Controls")
        self.win.config(bg='black')
        self.win.attributes('-topmost', True)
        self.win.overrideredirect(True)
        
        # Calculate size based on config
        columns = self.config.get("columns", 2)
        buttons = self.config.get("buttons", [])
        
        # Estimate window size
        btn_size = 55
        rows = 0
        col_idx = 0
        for btn in buttons:
            if btn.get("type") == "wide":
                if col_idx > 0: rows += 1  # finish current row
                rows += 1
                col_idx = 0
            else:
                col_idx += 1
                if col_idx >= columns:
                    rows += 1
                    col_idx = 0
        if col_idx > 0: rows += 1
        
        win_width = columns * btn_size + 10
        win_height = rows * btn_size + 20  # drag bar
        
        # Load Position
        x = self.settings.get("floating_toolbar_x", 100)
        y = self.settings.get("floating_toolbar_y", 100)
        
        self.win.geometry(f"{win_width}x{win_height}+{x}+{y}")
        
        # Frame
        self.frame = tk.Frame(self.win, bg='black')
        self.frame.pack(expand=True, fill='both')
        
        # Drag bar
        drag_bar = tk.Frame(self.frame, bg='black', height=15)
        drag_bar.pack(fill='x', padx=2, pady=2)
        drag_bar.pack_propagate(False)
        
        drag_indicator = tk.Label(drag_bar, text="⋯", font=("Segoe UI", 10), bg='black', fg='#666666')
        drag_indicator.pack(side=tk.RIGHT, padx=2)
        
        drag_bar.bind("<Button-1>", self._start_drag)
        drag_bar.bind("<B1-Motion>", self._on_drag)
        drag_bar.bind("<ButtonRelease-1>", self._on_release)
        
        drag_indicator.bind("<Button-1>", self._start_drag)
        drag_indicator.bind("<B1-Motion>", self._on_drag)
        drag_indicator.bind("<ButtonRelease-1>", self._on_release)
        
        # Button grid frame
        self.btn_frame = tk.Frame(self.frame, bg='black')
        self.btn_frame.pack(expand=True, fill='both', padx=2, pady=2)
        
        self._render_buttons()
        
        # Drag state
        self._drag_x = 0
        self._drag_y = 0
        
        # Context menu
        self.menu = tk.Menu(self.win, tearoff=0)
        self.menu.add_command(label="Close", command=self.destroy)
        self.win.bind("<Button-3>", lambda e: self.menu.tk_popup(e.x_root, e.y_root))
    
    def _render_buttons(self):
        """Render buttons from config."""
        for w in self.btn_frame.winfo_children():
            w.destroy()
        
        columns = self.config.get("columns", 2)
        buttons = self.config.get("buttons", [])
        
        off_white = '#CCCCCC'
        btn_font_icon = ('Segoe UI Symbol', 16, 'bold')
        btn_font_text = ('Segoe UI', 10, 'bold')
        
        row = 0
        col = 0
        
        for i, btn_cfg in enumerate(buttons):
            btn_type = btn_cfg.get("type", "square")
            icon = btn_cfg.get("icon", "")
            label = btn_cfg.get("label", "")
            keystroke = btn_cfg.get("keystroke", "")
            
            # Border frame
            border_color = off_white
            
            if btn_type == "wide":
                # Wide button spans all columns
                if col > 0:
                    row += 1
                    col = 0
                
                border_frame = tk.Frame(self.btn_frame, bg=border_color)
                border_frame.grid(row=row, column=0, columnspan=columns, sticky='nsew', padx=1, pady=1)
                
                btn = tk.Button(border_frame, text=label, command=lambda k=keystroke: self._send_key(k),
                               bg='black', fg=off_white, font=btn_font_text, borderwidth=0,
                               activebackground='#333', activeforeground='white', takefocus=False)
                btn.pack(expand=True, fill='both', padx=1, pady=1)
                btn.bind("<Enter>", lambda e, b=btn: b.config(fg='white'))
                btn.bind("<Leave>", lambda e, b=btn: b.config(fg=off_white))
                
                row += 1
                col = 0
            else:
                # Square button
                border_frame = tk.Frame(self.btn_frame, bg=border_color)
                border_frame.grid(row=row, column=col, sticky='nsew', padx=1, pady=1)
                
                display = icon if icon else label[:3]
                btn = tk.Button(border_frame, text=display, command=lambda k=keystroke: self._send_key(k),
                               bg='black', fg=off_white, font=btn_font_icon, borderwidth=0,
                               activebackground='#333', activeforeground='white', takefocus=False)
                btn.pack(expand=True, fill='both', padx=1, pady=1)
                btn.bind("<Enter>", lambda e, b=btn: b.config(fg='white'))
                btn.bind("<Leave>", lambda e, b=btn: b.config(fg=off_white))
                
                col += 1
                if col >= columns:
                    col = 0
                    row += 1
        
        # Configure grid weights
        for c in range(columns):
            self.btn_frame.grid_columnconfigure(c, weight=1)
        for r in range(row + 1):
            self.btn_frame.grid_rowconfigure(r, weight=1)
    
    def _send_key(self, keystroke):
        """Send keystroke to InteleViewer.
        
        Fix: Logic must explicitly EXCLUDE this tool's own windows ("Mosaic", "Controls")
        otherwise it finds itself and sends keys to nowhere.
        """
        if not keystroke: return
        
        try:
            # Find InteleViewer window
            hwnd = None
            def enum_cb(h, results):
                if win32gui.IsWindowVisible(h):
                    title = win32gui.GetWindowText(h).lower()
                    # Must contain 'inteleviewer' but NOT 'mosaic' or 'controls'
                    if "inteleviewer" in title:
                        if "mosaic" not in title and "controls" not in title:
                            results.append(h)
            
            windows = []
            win32gui.EnumWindows(enum_cb, windows)
            if windows:
                hwnd = windows[0]
            
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                
                # Parse and send keystroke
                keys = keystroke.lower().split('+')
                if len(keys) == 1:
                    pyautogui.press(keys[0])
                else:
                    pyautogui.hotkey(*keys)
                
                log_trace(f"Sent to InteleViewer: {keystroke}")
            else:
                log_trace("InteleViewer not found")
        except Exception as e:
            log_trace(f"Key send error: {e}")
    
    def _start_drag(self, event):
        self._drag_x = event.x
        self._drag_y = event.y
    
    def _on_drag(self, event):
        x = self.win.winfo_x() + event.x - self._drag_x
        y = self.win.winfo_y() + event.y - self._drag_y
        self.win.geometry(f"+{x}+{y}")
        
    def _on_release(self, event):
        """Save position on drag end."""
        self.settings["floating_toolbar_x"] = self.win.winfo_x()
        self.settings["floating_toolbar_y"] = self.win.winfo_y()
        self.save_callback()
    
    def destroy(self):
        self.win.destroy()
    
    def refresh(self, new_config):
        """Update with new config and re-render."""
        self.config = new_config
        self._render_buttons()

class MosaicToolsApp:
    def __init__(self):
        log_trace("App Init Started")
        
        # Initialize basic state
        self.running = True
        self.dictation_active = False 
        self.is_user_active = False 
        self.active_toasts = []
        
        # Initialize Root early to avoid AttributeError in threads/listeners
        self.root = tk.Tk()
        self.root.withdraw() # Hide main window if we only want floating toolbar
        
        # Load Settings First
        log_trace("Loading Settings...")
        self.settings = self.load_settings()
        self.doctor_name = self.settings.get("doctor_name", "Richter")
        self.start_beep_enabled = self.settings.get("start_beep_enabled", True)
        self.stop_beep_enabled = self.settings.get("stop_beep_enabled", True)
        self.start_beep_volume = self.settings.get("start_beep_volume", 0.04)
        self.stop_beep_volume = self.settings.get("stop_beep_volume", 0.04)
        self.dictation_pause_ms = self.settings.get("dictation_pause_ms", 1000)
        self.floating_toolbar_enabled = self.settings.get("floating_toolbar_enabled", False)
        self.indicator_enabled = self.settings.get("indicator_enabled", False)
        self.auto_stop_dictation = self.settings.get("auto_stop_dictation", False)
        self.dead_man_switch = self.settings.get("dead_man_switch", False)
        self.ptt_busy = False  # Mutex for PTT retry logic
        
        # Default to 'v' per user request, but allow override
        self.iv_report_hotkey = self.settings.get("iv_report_hotkey", "v")
        
        self.sync_lock = threading.Lock()
        self.sync_lock = threading.Lock()
        self.sync_check_token = 0 # Debounce token
        
        # Action Cancellation Token
        self.action_token = 0

        
        # Load Action Mappings (Unified)
        log_trace("Loading Mappings...")
        self.action_mappings = self.settings.get("action_mappings", {})
        
        # --- Migration Logic: Convert old 'button_mappings' to new 'action_mappings' ---
        legacy_mappings = self.settings.get("button_mappings", {})
        if legacy_mappings and not self.action_mappings:
            log_trace("Migrating legacy button mappings...")
            self.action_mappings = {}
            for btn, action in legacy_mappings.items():
                if action == ACTION_NONE: continue
                
                # Normalize old action names
                if action == "Get Prior (Alt+Shift+F3)": action = ACTION_GET_PRIOR
                elif action == "Scrape Clario": action = ACTION_SCRAPE
                
                if action not in self.action_mappings:
                    self.action_mappings[action] = {"hotkey": "", "mic_button": btn}
                else:
                    self.action_mappings[action]["mic_button"] = btn
            
            # Ensure defaults for unmapped actions
            for action in AVAILABLE_ACTIONS:
                if action == ACTION_NONE: continue
                if action not in self.action_mappings:
                    self.action_mappings[action] = {"hotkey": "", "mic_button": ""}
            
            self.settings["action_mappings"] = self.action_mappings
            # Keep legacy key for specific backward compat if needed, or remove?
            # self.settings.pop("button_mappings", None) 
        
        # If no mappings exist at all (fresh install), populate defaults
        if not self.action_mappings:
            self.action_mappings = {}
            # Reverse Default Mappings
            for btn, action in DEFAULT_MAPPINGS.items():
                if action == ACTION_NONE: continue
                if action not in self.action_mappings:
                    self.action_mappings[action] = {"hotkey": "", "mic_button": btn}
            
            # Add missing actions
            for action in AVAILABLE_ACTIONS:
                if action == ACTION_NONE: continue
                if action not in self.action_mappings:
                    self.action_mappings[action] = {"hotkey": "", "mic_button": ""}

        # --- Load Floating Buttons Config ---
        DEFAULT_FLOATING_BUTTONS = {
            "columns": 2,
            "buttons": [
                {"type": "square", "icon": "↕", "label": "", "keystroke": "ctrl+v"},
                {"type": "square", "icon": "↔", "label": "", "keystroke": "ctrl+h"},
                {"type": "square", "icon": "↺", "label": "", "keystroke": ","},
                {"type": "square", "icon": "↻", "label": "", "keystroke": "."},
                {"type": "wide", "icon": "", "label": "Zoom Out", "keystroke": "-"}
            ]
        }
        self.floating_buttons_config = self.settings.get("floating_buttons", DEFAULT_FLOATING_BUTTONS)
        if "floating_buttons" not in self.settings:
            self.settings["floating_buttons"] = self.floating_buttons_config

        # Initialize Helpers
        log_trace("Initializing Helpers...")
        self.ocr_engine = OCREngineWrapper()
        
        # Enforce Single Instance
        # log_trace("Checking for other instances...")
        # kill_other_instances() # Disabled due to crash reporting
        
        # Start background hardware listener
        log_trace("Starting HID listener thread...")
        self.thread = threading.Thread(target=self.background_listener, daemon=True)
        self.thread.start()

        # Start Keyboard Hotkey Listener
        if keyboard:
            try:
                log_trace("Initializing Keyboard Hotkeys...")
                self.setup_hotkeys()
            except Exception as e:
                log_trace(f"Hotkey Init Error: {e}")
        
        # --- UI Initialization ---
        log_trace("Creating Root Window...")
        # self.root already created at the top of __init__ to avoid race conditions
        self.root.title("Mosaic Tools")
        self.root.config(bg='black')
        self.root.attributes('-topmost', True)
        self.root.overrideredirect(True)
        
        # Load Window Position
        pos_x = self.settings.get("window_x", 100)
        pos_y = self.settings.get("window_y", 100)
        self.root.geometry(f"160x40+{pos_x}+{pos_y}")
        self.root.deiconify() # Show at correct position

        
        log_trace("Setting up UI components...")
        self.setup_ui()
        
        # Floating Toolbar Init
        self.toolbar_window = None
        if self.floating_toolbar_enabled:
            log_trace("Opening Floating Toolbar...")
            self.root.after(500, self.toggle_floating_toolbar, True)

        # Floating Indicator Init
        self.indicator_window = None
        if self.indicator_enabled:
            self.root.after(500, self.toggle_indicator, True)

        # Set up Windows Message Listener
        log_trace("Starting Message Listener thread...")
        self.msg_thread = threading.Thread(target=self.setup_message_listener, daemon=True)
        self.msg_thread.start()
        
        log_trace("App Init Complete. Entering MainLoop.")
        self.status_toast(f"Mosaic Tools Started ({self.doctor_name})", 2000)
        self.root.mainloop()

    def setup_ui(self):
        """Build the floating widget UI."""
        self.frame = tk.Frame(self.root, bg='#333333', relief='flat', bd=1)
        self.frame.pack(expand=True, fill='both')
        
        self.inner_frame = tk.Frame(self.frame, bg='black')
        self.inner_frame.pack(expand=True, fill='both', padx=1, pady=1)
        
        # Drag Handle (dots)
        self.drag_handle = tk.Label(self.inner_frame, text="⋮", font=("Segoe UI", 14), 
                                    bg='black', fg='#555555', cursor='fleur', width=2)
        self.drag_handle.pack(side=tk.LEFT, fill='y')
        
        self.drag_handle.bind("<Button-1>", self.start_drag)
        self.drag_handle.bind("<B1-Motion>", self.on_drag)
        self.drag_handle.bind("<ButtonRelease-1>", self.save_window_position)
        
        # Title Label (Clickable -> Settings)
        self.title_label = tk.Label(self.inner_frame, text="Mosaic Tools", font=("Segoe UI", 10, "bold"), 
                                    bg='black', fg='#CCCCCC', cursor='hand2')
        self.title_label.pack(side=tk.LEFT, expand=True, fill='both', padx=(0, 5))
        
        self.title_label.bind("<Button-1>", self.open_settings)
        self.title_label.bind("<Enter>", lambda e: self.title_label.config(fg='white'))
        self.title_label.bind("<Leave>", lambda e: self.title_label.config(fg='#CCCCCC'))

        self.create_context_menu()

        # Initialize drag state
        self._drag_start_x = 0
        self._drag_start_y = 0

    def toggle_floating_toolbar(self, show):
        if show:
            if not self.toolbar_window:
                self.toolbar_window = FloatingButtonsWindow(self.root, self.settings, self.save_settings_file)
        else:
            if self.toolbar_window:
                self.toolbar_window.destroy()
                self.toolbar_window = None

    def toggle_indicator(self, show):
        if show:
            if not self.indicator_window:
                self.indicator_window = FloatingIndicatorWindow(self.root, self.settings, self.save_settings_file)
                self.update_indicator_state()
        else:
            if self.indicator_window:
                self.indicator_window.destroy()
                self.indicator_window = None

    def update_indicator_state(self):
        """Update the visual state of the indicator if it exists."""
        if self.indicator_window:
            self.indicator_window.set_state(self.dictation_active)

    def create_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Reload", command=self.reload_app)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Exit", command=self.quit_app)
        
        # Bind right-click
        self.root.bind("<Button-3>", self.show_context_menu)
        self.frame.bind("<Button-3>", self.show_context_menu)
        self.inner_frame.bind("<Button-3>", self.show_context_menu)
        self.title_label.bind("<Button-3>", self.show_context_menu)
        self.drag_handle.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def quit_app(self):
        self.running = False
        self.root.quit()
        sys.exit()

    def reload_app(self):
        self.running = False
        self.root.destroy()
        # Relaunch
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def start_drag(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def on_drag(self, event):
        x = self.root.winfo_x() + event.x - self._drag_start_x
        y = self.root.winfo_y() + event.y - self._drag_start_y
        self.root.geometry(f"+{x}+{y}")
        
    def save_window_position(self, event=None):
        self.settings["window_x"] = self.root.winfo_x()
        self.settings["window_y"] = self.root.winfo_y()
        self.save_settings_file()

    def save_settings_file(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        settings_path = os.path.join(script_dir, "MosaicToolsSettings.json")
        try:
            with open(settings_path, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def open_settings(self, event=None):
        """Open the unified Settings Window."""
        win = tk.Toplevel(self.root)
        win.title("Mosaic Tools Settings")
        win.attributes("-topmost", True)
        
        # Center regarding screen
        x = self.root.winfo_x()
        y = self.root.winfo_y() + 50
        win.geometry(f"650x780+{x}+{y}") # Taller for the grid
        win.deiconify() 
        
        from tkinter import ttk
        notebook = ttk.Notebook(win)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        # --- TAB 1: General & Controls ---
        tab_main = tk.Frame(notebook)
        notebook.add(tab_main, text="Configuration")
        
        # 1. General Section
        gen_frame = tk.LabelFrame(tab_main, text="General Options", font=("Segoe UI", 10, "bold"), padx=15, pady=10)
        gen_frame.pack(fill='x', padx=10, pady=(10, 5))
        
        # Doctor Name
        name_frame = tk.Frame(gen_frame)
        name_frame.pack(fill='x', pady=2)
        tk.Label(name_frame, text="Doctor Name:").pack(side=tk.LEFT)
        name_var = tk.StringVar(value=self.doctor_name)
        tk.Entry(name_frame, textvariable=name_var).pack(side=tk.LEFT, padx=10, fill='x', expand=True)

        # Beep Settings - Inline Volume Controls
        start_beep_var = tk.BooleanVar(value=self.start_beep_enabled)
        stop_beep_var = tk.BooleanVar(value=self.stop_beep_enabled)
        start_vol_var = tk.IntVar(value=int(self.settings.get("start_beep_volume", 0.04) * 100))
        stop_vol_var = tk.IntVar(value=int(self.settings.get("stop_beep_volume", 0.04) * 100))
        
        # Start Beep Row (Checkbox + Volume Entry)
        start_row = tk.Frame(gen_frame)
        start_row.pack(fill='x', pady=2)
        tk.Checkbutton(start_row, text="Play Start Beep", variable=start_beep_var).pack(side=tk.LEFT)
        tk.Label(start_row, text="Vol:", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(10, 2))
        start_vol_entry = tk.Entry(start_row, textvariable=start_vol_var, width=3, font=("Segoe UI", 8))
        start_vol_entry.pack(side=tk.LEFT)
        tk.Label(start_row, text="%", font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # Custom Pause (Indented under Start Beep)
        pause_frame = tk.Frame(gen_frame)
        pause_frame.pack(fill='x', pady=2, padx=20) # Indented
        tk.Label(pause_frame, text="Dictation Start Delay (ms):").pack(side=tk.LEFT)
        pause_var = tk.StringVar(value=str(self.dictation_pause_ms))
        pause_entry = tk.Entry(pause_frame, textvariable=pause_var, width=6)
        pause_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(pause_frame, text="(Sound plays this long after keystroke)", 
                 font=("Segoe UI", 8, "italic"), fg="#666").pack(side=tk.LEFT, padx=5)
        
        # Toggle state of delay entry based on start beep
        def toggle_pause_entry(*args):
            state = 'normal' if start_beep_var.get() else 'disabled'
            pause_entry.config(state=state)
            start_vol_entry.config(state=state)
            
        start_beep_var.trace_add("write", toggle_pause_entry)
        toggle_pause_entry() # Init state

        # Stop Beep Row (Checkbox + Volume Entry)
        stop_row = tk.Frame(gen_frame)
        stop_row.pack(fill='x', pady=2)
        tk.Checkbutton(stop_row, text="Play Stop Beep", variable=stop_beep_var).pack(side=tk.LEFT)
        tk.Label(stop_row, text="Vol:", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(10, 2))
        stop_vol_entry = tk.Entry(stop_row, textvariable=stop_vol_var, width=3, font=("Segoe UI", 8))
        stop_vol_entry.pack(side=tk.LEFT)
        tk.Label(stop_row, text="%", font=("Segoe UI", 8)).pack(side=tk.LEFT)
        
        def toggle_stop_vol(*args):
            state = 'normal' if stop_beep_var.get() else 'disabled'
            stop_vol_entry.config(state=state)
        stop_beep_var.trace_add("write", toggle_stop_vol)
        toggle_stop_vol()
        
        indicator_var = tk.BooleanVar(value=self.indicator_enabled)
        tk.Checkbutton(gen_frame, text="Show Floating Dictation Indicator (Red/Grey Mic Pill)", variable=indicator_var).pack(anchor='w')
        
        auto_stop_var = tk.BooleanVar(value=self.auto_stop_dictation)
        tk.Checkbutton(gen_frame, text="Auto-stop Dictation on 'Process Report'", variable=auto_stop_var).pack(anchor='w')
        
        dead_man_var = tk.BooleanVar(value=self.dead_man_switch)
        tk.Checkbutton(gen_frame, text="Use 'Push-to-Talk' for Dictate Button (Dead Man's Switch)", variable=dead_man_var).pack(anchor='w')
        
        toolbar_var = tk.BooleanVar(value=self.floating_toolbar_enabled)
        tk.Checkbutton(gen_frame, text="Show Floating Toolbar (InteleViewer Controls)", variable=toolbar_var).pack(anchor='w')

        # Report Hotkey Config
        rep_frame = tk.Frame(gen_frame)
        rep_frame.pack(fill='x', pady=5)
        
        tk.Label(rep_frame, text="InteleViewer 'Show Report' Keystroke:").pack(side=tk.LEFT)
        
        rep_hk_var = tk.StringVar(value=self.iv_report_hotkey)
        rep_entry = tk.Entry(rep_frame, textvariable=rep_hk_var, width=15, state='readonly')
        rep_entry.pack(side=tk.LEFT, padx=5)
        
        rep_rec_btn = tk.Button(rep_frame, text="Rec", width=4, font=("Segoe UI", 8))
        rep_rec_btn.pack(side=tk.LEFT, padx=(2,0))
        
        def start_rep_rec():
            rep_rec_btn.config(text="...", bg="yellow")
            rep_hk_var.set("Recording...")
            win.update()
            threading.Thread(target=lambda: self.record_hotkey_thread(rep_hk_var, rep_rec_btn, win), daemon=True).start()
        
        rep_rec_btn.config(command=start_rep_rec)
        
        # Clear btn
        tk.Button(rep_frame, text="x", width=2, command=lambda: rep_hk_var.set(""), font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(2,0))

        tk.Label(rep_frame, text="(Default: 'v', used by Get Prior)", font=("Segoe UI", 8, "italic"), fg="#666").pack(side=tk.LEFT, padx=5)
        
        # Save handler
        def save_rep_hk(*args):
             val = rep_hk_var.get()
             if val == "Recording...": return
             self.settings["iv_report_hotkey"] = val
             self.iv_report_hotkey = val
             self.save_settings_file()
        rep_hk_var.trace_add("write", save_rep_hk)

        # 2. Control Map Section
        map_frame = tk.LabelFrame(tab_main, text="Control Map (Triggers)", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        map_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        tk.Label(map_frame, text="Assign actions to Keyboard Shortcuts AND/OR PowerMic Buttons.", 
                 font=("Segoe UI", 9, "italic"), fg="#555").pack(anchor='w', pady=(0, 10))

        # Grid Header
        grid_frame = tk.Frame(map_frame)
        grid_frame.pack(fill='both', expand=True)
        
        tk.Label(grid_frame, text="Action Function", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        tk.Label(grid_frame, text="Keyboard Shortcut", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        tk.Label(grid_frame, text="PowerMic Button", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky='w', padx=5, pady=5)
        tk.Label(grid_frame, text="Hardcoded Action", font=("Segoe UI", 9, "bold"), fg="#444").grid(row=0, column=3, sticky='w', padx=5, pady=5)

        # Dynamic Rows
        # Structure to hold Tk vars: action_vars[ActionName] = {'hotkey': Var, 'mic': Var}
        action_ui_vars = {}
        
        # Get PowerMic button list for dropdown
        mic_buttons = [""] + list(BUTTON_DEFINITIONS.keys()) # Add empty option
        
        row_idx = 1
        for action in AVAILABLE_ACTIONS:
            if action == ACTION_NONE: continue
            
            # Get current config or default
            config = self.action_mappings.get(action, {"hotkey": "", "mic_button": ""})
            
            # Label
            tk.Label(grid_frame, text=action).grid(row=row_idx, column=0, sticky='w', padx=5, pady=2)
            
            # Hotkey Frame (Entry + Record Button)
            hk_frame = tk.Frame(grid_frame)
            hk_frame.grid(row=row_idx, column=1, sticky='ew', padx=5, pady=2)
            
            hk_var = tk.StringVar(value=config.get("hotkey", ""))
            hk_entry = tk.Entry(hk_frame, textvariable=hk_var, width=15, state='readonly') # Readonly to prevent typing junk
            hk_entry.pack(side=tk.LEFT, fill='x', expand=True)
            
            rec_btn = tk.Button(hk_frame, text="Rec", width=4, font=("Segoe UI", 8))
            rec_btn.pack(side=tk.LEFT, padx=(2,0))
            
            # Hotkey Record Logic
            def start_rec(v=hk_var, b=rec_btn):
                b.config(text="...", bg="yellow")
                v.set("Recording...")
                win.update()
                # Run in thread to not freeze UI
                threading.Thread(target=lambda: self.record_hotkey_thread(v, b, win), daemon=True).start()
                
            rec_btn.config(command=start_rec)
            
            # Clear button (small x)
            def clear_hk(v=hk_var): v.set("")
            tk.Button(hk_frame, text="x", width=2, command=clear_hk, font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(2,0))

            # Mic Button Dropdown
            mic_var = tk.StringVar(value=config.get("mic_button", ""))
            mic_cb = ttk.Combobox(grid_frame, textvariable=mic_var, values=mic_buttons, state="readonly", width=18)
            mic_cb.grid(row=row_idx, column=2, sticky='ew', padx=5, pady=2)
            
            # Hardcoded Action Column
            hardcoded_text = ""
            if action == ACTION_TOGGLE_RECORD:
                hardcoded_text = "Record Button"
            elif action == ACTION_PROCESS_REPORT:
                hardcoded_text = "Skip Back Button"
            elif action == ACTION_SIGN_REPORT:
                hardcoded_text = "Checkmark Button"
            
            if hardcoded_text:
                tk.Label(grid_frame, text=hardcoded_text, font=("Segoe UI", 9, "italic"), fg="#555").grid(row=row_idx, column=3, sticky='w', padx=5, pady=2)

            action_ui_vars[action] = {'hotkey': hk_var, 'mic': mic_var}
            row_idx += 1

        tk.Label(grid_frame, text="Consider using combination of both CTRL and ALT, as well as function keys (F1-F12)", 
                 font=("Segoe UI", 8, "italic"), fg="#666").grid(row=row_idx, column=0, columnspan=3, sticky='w', padx=5, pady=(5, 0))

        # --- TAB 2: Button Studio ---
        tab_studio = tk.Frame(notebook)
        notebook.add(tab_studio, text="Button Studio")
        
        # Icon Library for dropdown
        ICON_LIBRARY = [
            "", "↑", "↓", "←", "→", "↕", "↔", "↺", "↻", "⟲", "⟳",
            "+", "−", "⊕", "⊖", "☰", "◎", "▶", "⏸", "⏹", "⚙", "☀", "★",
            "✚", "⎚", "▦", "◐", "◑", "⇑", "⇓", "⇐", "⇒"
        ]
        
        studio_main = tk.Frame(tab_studio, padx=10, pady=10)
        studio_main.pack(fill='both', expand=True)
        
        # --- Working Copy of buttons config (DEEP COPY to avoid mutations) ---
        import copy
        studio_buttons = copy.deepcopy(self.floating_buttons_config.get("buttons", []))
        studio_columns = tk.IntVar(value=self.floating_buttons_config.get("columns", 2))
        selected_btn_idx = tk.IntVar(value=0)
        
        # --- Top Row: Columns selector ---
        top_row = tk.Frame(studio_main)
        top_row.pack(fill='x', pady=(0, 10))
        tk.Label(top_row, text="Columns:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        col_spin = tk.Spinbox(top_row, from_=1, to=3, width=3, textvariable=studio_columns, state='readonly')
        col_spin.pack(side=tk.LEFT, padx=5)
        tk.Label(top_row, text="(max 9 buttons)", font=("Segoe UI", 8, "italic"), fg="#666").pack(side=tk.LEFT, padx=10)
        
        # --- Left: Live Preview + Button List ---
        left_frame = tk.Frame(studio_main)
        left_frame.pack(side=tk.LEFT, fill='both', expand=True)
        
        # Live Preview
        preview_lbl = tk.Label(left_frame, text="Live Preview", font=("Segoe UI", 9, "bold"))
        preview_lbl.pack(anchor='w')
        
        preview_frame = tk.Frame(left_frame, bg='#333', relief='sunken', bd=1)
        preview_frame.pack(fill='both', expand=True, pady=5)
        
        def render_preview():
            for w in preview_frame.winfo_children():
                w.destroy()
            cols = studio_columns.get()
            row = 0
            col = 0
            for i, btn_cfg in enumerate(studio_buttons):
                btn_type = btn_cfg.get("type", "square")
                icon = btn_cfg.get("icon", "")
                label = btn_cfg.get("label", "")
                is_selected = (i == selected_btn_idx.get())
                
                border_color = '#00ff00' if is_selected else '#888'
                
                if btn_type == "wide":
                    if col > 0:
                        row += 1
                        col = 0
                    bf = tk.Frame(preview_frame, bg=border_color)
                    bf.grid(row=row, column=0, columnspan=cols, sticky='nsew', padx=2, pady=2)
                    b = tk.Label(bf, text=label or "Wide", bg='black', fg='#ccc', font=("Segoe UI", 9))
                    b.pack(expand=True, fill='both', padx=1, pady=1)
                    b.bind("<Button-1>", lambda e, idx=i: select_button(idx))
                    row += 1
                    col = 0
                else:
                    bf = tk.Frame(preview_frame, bg=border_color)
                    bf.grid(row=row, column=col, sticky='nsew', padx=2, pady=2)
                    display = icon if icon else (label[:2] if label else "?")
                    b = tk.Label(bf, text=display, bg='black', fg='#ccc', font=("Segoe UI Symbol", 12))
                    b.pack(expand=True, fill='both', padx=1, pady=1)
                    b.bind("<Button-1>", lambda e, idx=i: select_button(idx))
                    col += 1
                    if col >= cols:
                        col = 0
                        row += 1
            for c in range(cols):
                preview_frame.grid_columnconfigure(c, weight=1)
            for r in range(row + 1):
                preview_frame.grid_rowconfigure(r, weight=1)
        
        def select_button(idx):
            selected_btn_idx.set(idx)
            render_preview()
            update_editor()
        
        # Button List (strip)
        list_lbl = tk.Label(left_frame, text="Button List", font=("Segoe UI", 9, "bold"))
        list_lbl.pack(anchor='w', pady=(10, 0))
        
        list_frame = tk.Frame(left_frame)
        list_frame.pack(fill='x', pady=5)
        
        def render_list():
            for w in list_frame.winfo_children():
                w.destroy()
            for i, btn_cfg in enumerate(studio_buttons):
                icon = btn_cfg.get("icon", "")
                label = btn_cfg.get("label", "")
                display = icon if icon else (label[:5] if label else "?")
                is_sel = (i == selected_btn_idx.get())
                bg = '#444' if is_sel else '#222'
                fg = 'white' if is_sel else '#ccc'
                lbl = tk.Label(list_frame, text=f"{i+1}:{display}", bg=bg, fg=fg, padx=5, pady=2, 
                              relief='ridge', cursor='hand2')
                lbl.pack(side=tk.LEFT, padx=1)
                lbl.bind("<Button-1>", lambda e, idx=i: select_button(idx))
        
        # Add/Delete/Reorder buttons
        ctrl_frame = tk.Frame(left_frame)
        ctrl_frame.pack(fill='x', pady=5)
        
        def add_button():
            if len(studio_buttons) >= 9:
                messagebox.showwarning("Limit", "Maximum 9 buttons allowed.")
                return
            studio_buttons.append({"type": "square", "icon": "?", "label": "", "keystroke": ""})
            selected_btn_idx.set(len(studio_buttons) - 1)
            render_preview()
            render_list()
            update_editor()
        
        def delete_button():
            if not studio_buttons: return
            idx = selected_btn_idx.get()
            if 0 <= idx < len(studio_buttons):
                del studio_buttons[idx]
                selected_btn_idx.set(max(0, idx - 1))
                render_preview()
                render_list()
                update_editor()
        
        def move_up():
            idx = selected_btn_idx.get()
            if idx > 0:
                studio_buttons[idx], studio_buttons[idx-1] = studio_buttons[idx-1], studio_buttons[idx]
                selected_btn_idx.set(idx - 1)
                render_preview()
                render_list()
        
        def move_down():
            idx = selected_btn_idx.get()
            if idx < len(studio_buttons) - 1:
                studio_buttons[idx], studio_buttons[idx+1] = studio_buttons[idx+1], studio_buttons[idx]
                selected_btn_idx.set(idx + 1)
                render_preview()
                render_list()
        
        def reset_to_defaults():
            """Reset to factory default button layout."""
            nonlocal studio_buttons
            default_btns = [
                {"type": "square", "icon": "↕", "label": "", "keystroke": "ctrl+v"},
                {"type": "square", "icon": "↔", "label": "", "keystroke": "ctrl+h"},
                {"type": "square", "icon": "↺", "label": "", "keystroke": ","},
                {"type": "square", "icon": "↻", "label": "", "keystroke": "."},
                {"type": "wide", "icon": "", "label": "Zoom Out", "keystroke": "-"}
            ]
            studio_buttons.clear()
            studio_buttons.extend(copy.deepcopy(default_btns))
            studio_columns.set(2)
            selected_btn_idx.set(0)
            render_preview()
            render_list()
            update_editor()
        
        tk.Button(ctrl_frame, text="+Add", command=add_button, width=6).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl_frame, text="-Del", command=delete_button, width=6).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl_frame, text="▲", command=move_up, width=3).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl_frame, text="▼", command=move_down, width=3).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl_frame, text="Reset", command=reset_to_defaults, width=6, fg='#c00').pack(side=tk.LEFT, padx=(10, 2))
        
        # --- Right: Button Editor ---
        right_frame = tk.LabelFrame(studio_main, text="Button Editor", font=("Segoe UI", 9, "bold"), padx=10, pady=10)
        right_frame.pack(side=tk.RIGHT, fill='y', padx=(10, 0))
        
        editor_type_var = tk.StringVar(value="square")
        editor_icon_var = tk.StringVar(value="")
        editor_label_var = tk.StringVar(value="")
        editor_key_var = tk.StringVar(value="")
        
        tk.Label(right_frame, text="Type:").grid(row=0, column=0, sticky='w', pady=2)
        type_frame = tk.Frame(right_frame)
        type_frame.grid(row=0, column=1, sticky='w')
        tk.Radiobutton(type_frame, text="Square", variable=editor_type_var, value="square").pack(side=tk.LEFT)
        tk.Radiobutton(type_frame, text="Wide", variable=editor_type_var, value="wide").pack(side=tk.LEFT)
        
        tk.Label(right_frame, text="Icon:").grid(row=1, column=0, sticky='w', pady=2)
        icon_cb = ttk.Combobox(right_frame, textvariable=editor_icon_var, values=ICON_LIBRARY, width=5, state='readonly')
        icon_cb.grid(row=1, column=1, sticky='w')
        
        tk.Label(right_frame, text="Label:").grid(row=2, column=0, sticky='w', pady=2)
        label_entry = tk.Entry(right_frame, textvariable=editor_label_var, width=12)
        label_entry.grid(row=2, column=1, sticky='w')
        
        tk.Label(right_frame, text="Keystroke:").grid(row=3, column=0, sticky='w', pady=2)
        key_frame = tk.Frame(right_frame)
        key_frame.grid(row=3, column=1, sticky='w')
        key_entry = tk.Entry(key_frame, textvariable=editor_key_var, width=10, state='readonly')
        key_entry.pack(side=tk.LEFT)
        key_rec_btn = tk.Button(key_frame, text="Rec", width=4, font=("Segoe UI", 8))
        key_rec_btn.pack(side=tk.LEFT, padx=2)
        
        def record_studio_key():
            key_rec_btn.config(text="...", bg="yellow")
            editor_key_var.set("Recording...")
            win.update()
            threading.Thread(target=lambda: self.record_hotkey_thread(editor_key_var, key_rec_btn, win), daemon=True).start()
        key_rec_btn.config(command=record_studio_key)
        
        def apply_editor():
            idx = selected_btn_idx.get()
            if 0 <= idx < len(studio_buttons):
                studio_buttons[idx]["type"] = editor_type_var.get()
                studio_buttons[idx]["icon"] = editor_icon_var.get()
                studio_buttons[idx]["label"] = editor_label_var.get()
                studio_buttons[idx]["keystroke"] = editor_key_var.get()
                render_preview()
                render_list()
        
        tk.Label(right_frame, text="Use saved shortcuts for Inteleviewer -\nUtilities | User Preferences | User Configuration", 
                 font=("Segoe UI", 8), fg="#888", justify=tk.LEFT).grid(row=4, column=0, columnspan=2, sticky='w', pady=(2,5))

        tk.Button(right_frame, text="Apply", command=apply_editor, width=8).grid(row=5, column=0, columnspan=2, pady=10)
        
        def update_editor():
            idx = selected_btn_idx.get()
            if 0 <= idx < len(studio_buttons):
                btn_cfg = studio_buttons[idx]
                editor_type_var.set(btn_cfg.get("type", "square"))
                editor_icon_var.set(btn_cfg.get("icon", ""))
                editor_label_var.set(btn_cfg.get("label", ""))
                editor_key_var.set(btn_cfg.get("keystroke", ""))
        
        # Wire column change to re-render
        studio_columns.trace_add("write", lambda *args: render_preview())
        
        # Initial render
        render_preview()
        render_list()
        update_editor()
        
        # Store studio state for save function
        def get_studio_config():
            return {"columns": studio_columns.get(), "buttons": studio_buttons}
        
        # --- TAB 3: Advanced / AHK ---
        tab_adv = tk.Frame(notebook)
        notebook.add(tab_adv, text="Advanced / AHK")
        
        adv_frame = tk.Frame(tab_adv, padx=15, pady=15)
        adv_frame.pack(fill='both', expand=True)
        
        lbl_adv = tk.Label(adv_frame, text="AutoHotkey & Message Integration", font=("Segoe UI", 11, "bold"))
        lbl_adv.pack(anchor='w', pady=(0, 10))
        
        ahk_guide_text = (
            "MOSAIC TOOLS - ADVANCED INTEGRATION\n\n"
            "This application listens for standard HID events (PowerMic) and Global Keyboard Hooks. "
            "However, you can also trigger functions from external scripts (like AutoHotkey) "
            "using Windows Messages (PostMessage).\n\n"
            "Target Window Class: MosaicToolsMessageWindow\n"
            "Target Window Title: Mosaic Tools Service\n\n"
            "Message IDs (wParam=0, lParam=0):\n"
            "• 0x0401 (1025) : Trigger Critical Findings\n"
            "• 0x0402 (1026) : Trigger Debug Window\n"
            "• 0x0403 (1027) : Trigger System Beep / Toggle State\n"
            "• 0x0404 (1028) : Trigger Show Report\n"
            "• 0x0405 (1029) : Trigger Capture Series/Image\n"
            "• 0x0406 (1030) : Trigger Get Prior\n"
            "• 0x0407 (1031) : Trigger Start/Stop Recording\n"
            "• 0x0408 (1032) : Trigger Process Report\n"
            "• 0x0409 (1033) : Trigger Sign Report\n\n"
            "AHK Example Script:\n"
            "--------------------------------------------------\n"
            "DetectHiddenWindows, On\n"
            "; Trigger 'Get Prior' with Ctrl+Alt+P\n"
            "^!p::\n"
            "  PostMessage, 0x0406, 0, 0, , Mosaic Tools Service\n"
            "return\n"
            "--------------------------------------------------\n"
        )
        
        txt_adv = tk.Text(adv_frame, wrap=tk.WORD, font=("Consolas", 9), bg="#f4f4f4", relief="flat")
        txt_adv.insert("1.0", ahk_guide_text)
        txt_adv.config(state="disabled")
        txt_adv.pack(fill='both', expand=True)

        # --- Footer Buttons ---
        btn_frame = tk.Frame(win, pady=10)
        btn_frame.pack(side=tk.BOTTOM, fill='x')
        
        def save(event=None):
            # General
            new_name = name_var.get().strip()
            if new_name:
                self.doctor_name = new_name
                self.settings["doctor_name"] = new_name
            
            self.start_beep_enabled = start_beep_var.get()
            self.settings["start_beep_enabled"] = self.start_beep_enabled
            
            self.stop_beep_enabled = stop_beep_var.get()
            self.settings["stop_beep_enabled"] = self.stop_beep_enabled
            
            try:
                self.start_beep_volume = max(1, min(100, int(start_vol_var.get()))) / 100.0
                self.settings["start_beep_volume"] = self.start_beep_volume
            except: pass
            try:
                self.stop_beep_volume = max(1, min(100, int(stop_vol_var.get()))) / 100.0
                self.settings["stop_beep_volume"] = self.stop_beep_volume
            except: pass

            try:
                ms = int(pause_var.get())
                self.dictation_pause_ms = ms
                self.settings["dictation_pause_ms"] = ms
            except: pass
            
            new_toolbar = toolbar_var.get()
            if new_toolbar != self.floating_toolbar_enabled:
                self.floating_toolbar_enabled = new_toolbar
                self.settings["floating_toolbar_enabled"] = new_toolbar
                self.toggle_floating_toolbar(new_toolbar)

            new_indicator = indicator_var.get()
            if new_indicator != self.indicator_enabled:
                self.indicator_enabled = new_indicator
                self.settings["indicator_enabled"] = new_indicator
            new_indicator = indicator_var.get()
            if new_indicator != self.indicator_enabled:
                self.indicator_enabled = new_indicator
                self.settings["indicator_enabled"] = new_indicator
                self.toggle_indicator(new_indicator)
            
            self.auto_stop_dictation = auto_stop_var.get()
            self.settings["auto_stop_dictation"] = self.auto_stop_dictation

            self.dead_man_switch = dead_man_var.get()
            self.settings["dead_man_switch"] = self.dead_man_switch

            # Mappings
            # Traverse UI vars and update Action Mappings
            for action, vars_dict in action_ui_vars.items():
                hk = vars_dict['hotkey'].get()
                mic = vars_dict['mic'].get()
                
                # Update schema
                if action not in self.action_mappings:
                    self.action_mappings[action] = {}
                
                self.action_mappings[action]["hotkey"] = hk
                self.action_mappings[action]["mic_button"] = mic

            self.settings["action_mappings"] = self.action_mappings
            
            # Button Studio: Save config
            new_floating_config = get_studio_config()
            self.floating_buttons_config = new_floating_config
            self.settings["floating_buttons"] = new_floating_config
            
            # Refresh floating toolbar if visible
            if self.toolbar_window:
                self.toolbar_window.refresh(new_floating_config)
            
            # Persist
            self.save_settings_file()
            
            # Apply Changes Immediately (Reload Hotkeys)
            self.setup_hotkeys()
            
            self.status_toast("Settings Saved & Applied", 1500)
            win.destroy()

        tk.Button(btn_frame, text="Save & Apply", command=save, width=15, bg='#dddddd', font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=20)
        tk.Button(btn_frame, text="Cancel", command=win.destroy, width=10).pack(side=tk.RIGHT, padx=5)

    def normalize_hotkey(self, key_string):
        """Clean up raw hotkey strings from the keyboard library."""
        if not key_string: return ""
        
        # Split by +
        parts = key_string.split('+')
        clean_parts = []
        
        seen = set()
        
        # Priority mapping for canonical ordering
        order = {'ctrl': 0, 'shift': 1, 'alt': 2, 'windows': 3}
        
        modifiers = []
        keys = []
        
        for p in parts:
            p = p.lower().strip()
            # Filter out common artifacts
            if p in ['n', 'unknown', 'decimal', '#']: 
                continue # Skip garbage scan codes often left by NumLock/Shift
            
            # Unify names
            if 'control' in p: p = 'ctrl'
            if 'win' in p and p != 'windows': p = 'windows'
            
            if p in seen: continue
            seen.add(p)
            
            if p in ['ctrl', 'shift', 'alt', 'windows', 'left windows', 'right windows']:
                modifiers.append(p)
            else:
                keys.append(p)
        
        # Sort modifiers
        modifiers.sort(key=lambda k: order.get(k, 10))
        
        # Reconstruct: Modifiers + Keys
        final_parts = modifiers + keys
        return '+'.join(final_parts)

    def record_hotkey_thread(self, var, btn, win):
        """Helper to record a single hotkey manually to avoid ghost keys."""
        if not keyboard:
            self.root.after(0, lambda: messagebox.showerror("Error", "Keyboard library not loaded."))
            self.root.after(0, lambda: btn.config(text="Rec", bg="SystemButtonFace"))
            return

        try:
            # Short sleep to clear initial click events
            time.sleep(0.3)
            
            # 0. Flush Buffer (Private queue access to ensure clean state)
            # The keyboard listener runs in background and queues events. 
            # We must clear old events (like the releasing of the Record button)
            try:
                while hasattr(keyboard, '_queue') and not keyboard._queue.empty():
                    keyboard._queue.get_nowait()
            except: pass

            # 1. Initialize Modifiers State
            # (Check what is currently held down so we don't miss held Ctrl/Shift)
            current_modifiers = set()
            possible_mods = ['ctrl', 'shift', 'alt', 'windows', 'left windows', 'right windows']
            for m in possible_mods:
                if keyboard.is_pressed(m):
                    current_modifiers.add(m.replace('left ', '').replace('right ', ''))

            # 2. Event Loop
            raw_key_name = None
            
            # Loop until we get a non-modifier DOWN event
            while True:
                # suppress=True ensures we catch it and don't type it into apps
                event = keyboard.read_event(suppress=True)
                
                name = event.name.lower()
                is_mod = any(m in name for m in ['ctrl', 'shift', 'alt', 'win', 'cmd'])
                
                # Normalize modifier name
                clean_mod_name = name
                if 'control' in name: clean_mod_name = 'ctrl'
                if 'win' in name and name != 'windows': clean_mod_name = 'windows'
                
                # Track Modifiers
                if is_mod:
                    if event.event_type == 'down':
                        current_modifiers.add(clean_mod_name)
                    elif event.event_type == 'up':
                        if clean_mod_name in current_modifiers:
                            current_modifiers.remove(clean_mod_name)
                    continue
                
                # Valid Non-Modifier Key Pressed?
                if event.event_type == 'down':
                     # Check for garbage scan codes
                    if name in ['unknown', 'decimal', '#', 'n']:
                        continue
                        
                    raw_key_name = name
                    break # Captured!

            # 3. Construct Hotkey String
            # Sort modifiers: Ctrl -> Shift -> Alt -> Win
            order = {'ctrl': 0, 'shift': 1, 'alt': 2, 'windows': 3}
            sorted_mods = sorted(list(current_modifiers), key=lambda k: order.get(k, 10))
            
            parts = sorted_mods + [raw_key_name]
            final_hotkey = '+'.join(parts)
            
            log_trace(f"Recorded Manually: '{final_hotkey}'")
            self.root.after(0, lambda: var.set(final_hotkey))
            
        except Exception as e:
            print(f"Rec Error: {e}")
        finally:
            self.root.after(0, lambda: btn.config(text="Rec", bg="SystemButtonFace"))

    def load_settings(self):
        """Load settings from JSON, or prompt and create if missing."""
        import sys
        if getattr(sys, 'frozen', False):
            script_dir = os.path.dirname(sys.executable)
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        settings_path = os.path.join(script_dir, "MosaicToolsSettings.json")
        
        if os.path.exists(settings_path):
            try:
                with open(settings_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading settings: {e}")
        
        # Missing or corrupt settings - Onboarding
        return self.show_onboarding(settings_path)

    def show_onboarding(self, settings_path):
        """Prompt user for their name and create the settings file."""
        root = tk.Tk()
        root.withdraw()
        
        msg = ("Welcome to Mosaic Tools!\n\n"
               "The tool needs to know your last name so it can correctly identify "
               "and filter your name out of clinical statements when parsing Clario exam notes.\n\n"
               "Example: If the note says 'Transferred Smith to Jones', and you are Dr. Smith, "
               "the tool will know that Jones is the contact person.")
        
        messagebox.showinfo("First Run Onboarding", msg)
        
        name = simpledialog.askstring("Doctor Name", "Enter your last name (e.g., Smith):", parent=root)
        if not name or name.strip() == "":
            name = "Radiologist" # Fallback
        
        name = name.strip()
        
        settings = {
            "_instructions": [
                "This file stores settings for MosaicTools.",
                "doctor_name: Your last name. This is used to filter you out of 'EXAM NOTE' scrapes.",
                "If you change your name, update it here and restart the tool."
            ],
            "doctor_name": name,
            "beep_enabled": True,
            "window_x": 100,
            "window_y": 100
        }
        
        try:
            with open(settings_path, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")
            
        root.destroy()
        return settings

    def reposition_toasts(self):
        """Stack active toasts from bottom-right upwards."""
        if not self.active_toasts: return
        
        # Get screen dimensions from root (or primary monitor)
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        
        # Bottom anchor
        current_y = screen_h - 100 # Start higher up to clear taskbar
        
        # Iterate backwards (newest at bottom) or forwards? 
        # Let's say newest is at the bottom (index -1), oldest at top (index 0).
        # We want to stack them upwards from bottom.
        
        # Filter out destroyed widgets just in case
        self.active_toasts = [t for t in self.active_toasts if t.winfo_exists()]
        
        for toast in reversed(self.active_toasts):
            try:
                w = toast.winfo_width()
                h = toast.winfo_height()
                if w <= 1 or h <= 1: # Not fully rendered yet, use requested check or default
                    w = toast.winfo_reqwidth()
                    h = toast.winfo_reqheight()
                
                x = screen_w - w - 20
                y = current_y - h
                
                toast.geometry(f"+{x}+{y}")
                current_y = y - 10 # Padding between toasts
            except: pass

    def status_toast(self, message, duration=5000):
        """Show a temporary, non-blocking status window that stacks."""
        log_trace(f"Toast: {message}")
        toast = tk.Toplevel(self.root)
        toast.withdraw() # Hide immediately
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.attributes("-alpha", 0.9)
        
        # Style
        label = tk.Label(toast, text=message, bg="#333333", fg="white", 
                         padx=15, pady=8, font=("Segoe UI", 10, "bold"))
        label.pack()
        
        # Force size calc
        toast.update_idletasks()
        
        # Add to stack
        self.active_toasts.append(toast)
        self.reposition_toasts()
        
        # Show
        toast.deiconify()
        
        # Auto-destroy closure
        def _destroy():
            if toast in self.active_toasts:
                self.active_toasts.remove(toast)
            if toast.winfo_exists():
                toast.destroy()
            self.root.after(10, self.reposition_toasts) # Re-stack after removal
            
        self.root.after(duration, _destroy)

    def show_results_window(self, formatted, raw):
        """Show results in a single text window with both formatted and raw."""
        win = tk.Toplevel(self.root)
        win.withdraw() # Hide immediately
        win.title("Clario Scrape - Troubleshooting")
        win.attributes("-topmost", True)
        # Position in center, auto-size height
        # Position in center, auto-size height
        win.geometry("700x500") # Reasonable default size

        win.deiconify() # Show at correct position
        
        # Single Text Box
        frame = tk.Frame(win)
        frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        text_box = tk.Text(frame, wrap=tk.WORD, font=("Segoe UI", 11), padx=10, pady=10)
        text_box.pack(expand=True, fill=tk.BOTH)
        
        # Combined content
        combined = f"FORMATTED:\n{formatted}\n\n{'='*50}\n\nRAW:\n{raw}"
        text_box.insert(tk.END, combined)
        
        # Buttons
        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="Copy All", width=20, 
                  command=lambda: pyperclip.copy(combined)).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Close", width=15, 
                  command=win.destroy).pack(side=tk.LEFT, padx=10)
        
        # Copy to clipboard immediately
        pyperclip.copy(formatted)

    def find_clario_window(self):
        """Find the Chrome window with Clario - Worklist and return a WindowSpecification."""
        desktop = Desktop(backend="uia")
        try:
            # We first find the actual window to get its exact title
            for w in desktop.windows(visible_only=True):
                try:
                    title = w.window_text()
                    if "clario" in title.lower() and "worklist" in title.lower():
                        log_trace(f"Found Clario Window: '{title}'")
                        # Return a SPECIFICATION so we can call .child_window() etc.
                        return desktop.window(handle=w.handle)
                except: continue
        except Exception as e:
            log_trace(f"Error finding Clario window: {e}")
        return None

    def find_chrome_content_area(self, chrome_window_spec):
        """DEPRECATED: Now returns the spec directly."""
        return chrome_window_spec

    def get_exam_note_elements(self, window_spec, depth=None):
        """Exhaustive search for EXAM NOTE DataItems using WindowSpecification with 30s timeout."""
        try:
            log_trace(f"Starting deep EXAM NOTE search (30s timeout)...")
            
            # 1. Strict Search (DataItem)
            # Create a specification for the element
            spec = window_spec.child_window(title_re=".*EXAM NOTE.*", control_type="DataItem")
            
            # Using a loop to check existence with a repeating toast
            start_time = time.time()
            timeout = 30.0
            
            while time.time() - start_time < timeout:
                try:
                    # exists() is safe to call on a WindowSpecification
                    if spec.exists(timeout=1.0):
                        log_trace("Spec exists, extracting matches...")
                        results = []
                        # Correct way to get multiple matches: loop with found_index
                        for i in range(15): # Max 15 notes
                            try:
                                match_spec = window_spec.child_window(title_re=".*EXAM NOTE.*", control_type="DataItem", found_index=i)
                                if match_spec.exists(timeout=0.1):
                                    t = match_spec.window_text()
                                    if t: results.append(t)
                                else:
                                    break # No more matches
                            except: break
                        
                        if results:
                            log_trace(f"SUCCESS: Found {len(results)} matches.")
                            return results
                except Exception as e:
                    # If the spec search itself crashes (like the TypeError before), log it and break or continue
                    log_trace(f"Loop search error: {e}")
                    break

                # Update toast every 5 seconds
                elapsed = time.time() - start_time
                if int(elapsed) % 5 == 0 and int(elapsed) > 0:
                    self.show_toast("Still searching Clario notes...")
                
                time.sleep(0.5)

            # 2. Final Fallback: Weak search (title only)
            log_trace("Strict search timed out or failed, trying final weak title-only search...")
            weak_spec = window_spec.child_window(title_re=".*EXAM NOTE.*")
            if weak_spec.exists(timeout=5.0):
                matches = []
                for i in range(15):
                    try:
                        m = window_spec.child_window(title_re=".*EXAM NOTE.*", found_index=i)
                        if m.exists(timeout=0.1):
                            t = m.window_text()
                            if t: matches.append(t)
                        else: break
                    except: break
                
                if matches:
                    log_trace(f"Found {len(matches)} weak matches.")
                    return matches
                 
            log_trace("TIMEOUT: No EXAM NOTE matches found after 30s+.")
        except Exception as e:
            log_trace(f"Outer search failed: {e}")
            import traceback
            log_trace(traceback.format_exc())
            
        return []

    def find_mosaic_editor(self):
        """Exhaustive search for the editor, matching testMosaic.py logic."""
        desktop = Desktop(backend="uia")
        try:
            # 1. Find the Main Window
            mosaic_win = None
            for win in desktop.windows(visible_only=True):
                try:
                    title = win.window_text().lower()
                    if "mosaic" in title and ("reporting" in title or "info hub" in title):
                        mosaic_win = win
                        break
                except: continue
            
            if not mosaic_win:
                return None

            # 2. Find the WebView2 container (Crucial for Mosaic)
            # testMosaic.py finds this via automation_id="webView"
            search_root = None
            try:
                for child in mosaic_win.descendants(automation_id="webView"):
                    search_root = child
                    break
            except: pass
            
            if not search_root:
                search_root = mosaic_win # Fallback to window

            # 3. Exhaustive search for the specific editor box
            # We look for the text "ADDENDUM:" which the user provided in the dump
            for elem in search_root.descendants():
                try:
                    name = (elem.element_info.name or "").upper()
                    text = (elem.window_text() or "").upper()
                    cname = (elem.element_info.class_name or "").lower()
                    
                    # Match by the unique "tiptap" class as priority
                    if "TIPTAP" in cname or "PROSEMIRROR" in cname:
                        return elem
                    
                    # If we find the "ADDENDUM:" text, return its parent (the editor group)
                    if ("ADDENDUM:" in text or "ADDENDUM:" in name) and elem.element_info.control_type != "Text":
                        return elem
                    elif ("ADDENDUM:" in text or "ADDENDUM:" in name):
                        try:
                            parent = elem.parent()
                            if parent: return parent
                        except: pass
                except:
                    continue

        except Exception as e:
            print(f"Error in find_mosaic_editor: {e}")
        return None

    def format_note(self, raw_text):
        """
        Parses the Clario exam note based on specific clinical rules.
        """
        try:
            def normalize_time_str(ts):
                """Ensure there is a space before AM/PM for strptime."""
                ts = ts.upper().strip()
                ts = re.sub(r"([0-9])(AM|PM)", r"\1 \2", ts)
                return ts

            # 1. Extract Name (Look for clinician patterns)
            # Strategy: Find the transfer/connection verbs and extract the non-doctor party.
            segments = []
            
            # Look for contact verbs
            # Pattern: (Verb) (Person A) (to/with) (Person B)
            verb_pattern = r"(?:Transferred|Connected\s+with|Connected|Was\s+connected\s+with|Spoke\s+to|Spoke\s+with|Discussed\s+with)\s+(.+?)(?=\s+(?:at|@|said|confirmed|stated|reported|declined|who|are|is)|\s+\d{2}/\d{2}/|$)"
            verb_match = re.search(verb_pattern, raw_text, re.IGNORECASE)
            if verb_match:
                full_segment = verb_match.group(1)
                # Split by "to" or "with" or "w/" to get the two people involved
                parts = re.split(r"\s+(?:to|with|w/)\s+", full_segment, flags=re.IGNORECASE)
                segments.extend(parts)
            
            # Fallback/Additional: Individual title matches (Dr. or Nurse)
            name_pattern = r"((?:Dr\.|Nurse)\s+.+?)(?=\s+(?:at|with|to|w/|@|said|confirmed|stated|reported|declined|who|confirm|are|is)|\s+(?:Dr\.|Nurse)|\s+\d{2}/\d{2}/|$)"
            segments.extend(re.findall(name_pattern, raw_text, re.IGNORECASE))

            title_and_name = "Dr. / Nurse [Name not found]"
            for s in segments:
                s_clean = s.strip()
                # Remove leading prepositions that might have been accidentally captured
                s_clean = re.sub(r"^(?:with|to|at|from|and|@|w/)\s+", "", s_clean, flags=re.IGNORECASE)
                
                # Skip if it's the current doctor or too short to be a name
                if self.doctor_name not in s_clean and len(s_clean) > 2:
                    # Clean up common artifacts (e.g. if capture was "...nurse to...")
                    s_clean = re.split(r"\s+(?:to|at|@|w/)\s+", s_clean, flags=re.IGNORECASE)[0]
                    title_and_name = s_clean.title()
                    break

            # 2. Extract End Timestamp (Always Central Time)
            # Format at end: 01/07/2026 11:35 PM
            end_match = re.search(r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2}\s*(?:AM|PM))", raw_text)
            end_date = "N/A"
            end_time_ct_str = "N/A"
            dt_end = None
            
            # NOTE DIALOG FORMAT: "@ 11:02 PM Central time" or "at 11:27 PM central time" (no date)
            # Only match if there's NO date pattern in the text (to avoid false positives)
            dialog_time_match = None
            # Note: Added '.' as trigger for case "connected. 10:46PM"
            if not end_match:
                dialog_time_match = re.search(r"(?:@|at|\.)\s+(\d{1,2}:\d{2}\s*(?:AM|PM))\s*([a-zA-Z\s]+)?", raw_text, re.IGNORECASE)
            
            if end_match:
                # Full format with date
                end_date = end_match.group(1)
                end_time_ct_str = normalize_time_str(end_match.group(2))
                try:
                    dt_end = datetime.datetime.strptime(f"{end_date} {end_time_ct_str}", "%m/%d/%Y %I:%M %p")
                except:
                    pass
            elif dialog_time_match:
                # Note Dialog format - infer date from current time
                text_time_str = normalize_time_str(dialog_time_match.group(1))
                
                # Robust Timezone Logic (Same as Part 3)
                raw_tz = dialog_time_match.group(2).strip() if dialog_time_match.group(2) else ""
                timezone_str = "Central Time" # Default
                
                valid_tz_tokens = [
                    "eastern", "central", "mountain", "pacific", 
                    "east", "est", "cst", "mst", "pst", "edt", "cdt", "mdt", "pdt"
                ]
                
                # Check the first word of the capture
                first_word = raw_tz.split(' ')[0].lower()
                if first_word in valid_tz_tokens:
                    # Construct nice display name
                    if first_word in ['east', 'eastern', 'est', 'edt']: timezone_str = "Eastern Time"
                    elif first_word in ['central', 'cst', 'cdt']: timezone_str = "Central Time"
                    elif first_word in ['mountain', 'mst', 'mdt']: timezone_str = "Mountain Time"
                    elif first_word in ['pacific', 'pst', 'pdt']: timezone_str = "Pacific Time"
                
                # Timezone Offsets (vs Central)
                tz_offsets = {
                    'Eastern Time': 1, 'Central Time': 0, 'Mountain Time': -1, 'Pacific Time': -2
                }
                offset = tz_offsets.get(timezone_str, 0)
                
                # Get current time and try to infer date
                now = datetime.datetime.now()
                today = now.date()
                yesterday = today - datetime.timedelta(days=1)
                
                try:
                    # Parse just the time
                    note_time = datetime.datetime.strptime(text_time_str, "%I:%M %p").time()
                    
                    # Convert note time to Central from its timezone
                    note_hour_ct = note_time.hour - offset
                    if note_hour_ct < 0:
                        note_hour_ct += 24
                    elif note_hour_ct >= 24:
                        note_hour_ct -= 24
                    
                    # Create datetime for today with this time
                    note_dt_today = datetime.datetime.combine(today, datetime.time(note_hour_ct, note_time.minute))
                    
                    # If the note time (in Central) is in the future, it must be yesterday
                    if note_dt_today > now:
                        end_date = yesterday.strftime("%m/%d/%Y")
                    else:
                        end_date = today.strftime("%m/%d/%Y")
                    
                    end_time_ct_str = text_time_str
                    
                    # Return formatted result directly for dialog format
                    return f"Critical findings were discussed with and acknowledged by {title_and_name} at {text_time_str} {timezone_str} on {end_date}."
                    
                except Exception as e:
                    print(f"Date inference error: {e}")
                    end_time_ct_str = text_time_str

            # 3. Extract Text Time and Timezone
            # Format in text: at 12:34 AM Eastern Time or @ 8:08am east or . 10:46PM
            # Regex Explanation:
            # (?:at|@|\.)\s+     -> Trigger 'at', '@', or '.'
            # (\d{1,2}:\d{2}\s*(?:AM|PM)) -> Time (Group 1)
            # \s*                -> Gap
            # ([a-zA-Z\s]+)?     -> Potential Timezone (Group 2) - captured leniently then cleaned
            text_time_match = re.search(r"(?:at|@|\.)\s+(\d{1,2}:\d{2}\s*(?:AM|PM))\s*([a-zA-Z\s]+)?", raw_text, re.IGNORECASE)
            final_time_display = "N/A"
            
            if text_time_match and dt_end:
                text_time_str = normalize_time_str(text_time_match.group(1))
                
                # Clean up timezone string (it effectively grabs everything after time... dangerous?)
                # We need to filter it against known valid timezone strings to avoid grabbing "..confirmed"
                raw_tz = text_time_match.group(2).strip() if text_time_match.group(2) else ""
                
                # Split by space and look for valid timezone token at start
                timezone_str = "Central Time" # Default
                valid_tz_tokens = [
                    "eastern", "central", "mountain", "pacific", 
                    "east", "est", "cst", "mst", "pst", "edt", "cdt", "mdt", "pdt"
                ]
                
                # Check the first word of the capture
                first_word = raw_tz.split(' ')[0].lower()
                if first_word in valid_tz_tokens:
                    # Construct nice display name
                    if first_word in ['east', 'eastern', 'est', 'edt']: timezone_str = "Eastern Time"
                    elif first_word in ['central', 'cst', 'cdt']: timezone_str = "Central Time"
                    elif first_word in ['mountain', 'mst', 'mdt']: timezone_str = "Mountain Time"
                    elif first_word in ['pacific', 'pst', 'pdt']: timezone_str = "Pacific Time"
                
                try:
                    # Convert text time to datetime (assume same date initially)
                    dt_text = datetime.datetime.strptime(f"{end_date} {text_time_str}", "%m/%d/%Y %I:%M %p")
                    
                    # Timezone Offsets (vs Central)
                    tz_offsets = {
                        'Eastern Time': 1,
                        'Central Time': 0,
                        'Mountain Time': -1,
                        'Pacific Time': -2
                    }
                    offset = tz_offsets.get(timezone_str, 0)
                    
                    # Convert text time to Central for comparison
                    dt_text_ct = dt_text - datetime.timedelta(hours=offset)
                    
                    # Handle day wraparound
                    diff_seconds = (dt_text_ct - dt_end).total_seconds()
                    if diff_seconds > 43200: # 12 hours
                        dt_text_ct -= datetime.timedelta(days=1)
                    elif diff_seconds < -43200:
                        dt_text_ct += datetime.timedelta(days=1)
                    
                    time_diff_min = abs((dt_text_ct - dt_end).total_seconds()) / 60
                    
                    if time_diff_min <= 10:
                        final_time_display = f"{text_time_str} {timezone_str}"
                    else:
                        final_time_display = f"{text_time_str} {timezone_str} (? Diff: {int(time_diff_min)}m)"
                except:
                    final_time_display = text_time_str
            else:
                final_time_display = end_time_ct_str

            return f"Critical findings were discussed with and acknowledged by {title_and_name} at {final_time_display} on {end_date}."

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            return f"Error parsing: {e}\nRaw: {raw_text}\n\n{error_details}"

    def paste_to_target(self):
        """Standard clipboard paste."""
        pyautogui.hotkey('ctrl', 'v')

    def _activate_mosaic(self, target_label="Transcript"):
        """Robust activation and targeted focus on a specific Mosaic editor."""
        log_trace(f"Activating Mosaic (Target: {target_label})")
        activated = self._activate_window_by_title(["mosaic"], ["reporting", "info hub"])
        
        if activated and target_label:
            # Synchronous targeting: Ensure focus is set BEFORE the paste logic continues.
            self._target_editor_by_label(target_label)
            
        return activated

    def _target_editor_by_label(self, label_name):
        """Finds and clicks a specific Mosaic editor by anchoring to its label (e.g., 'Transcript' or 'Final Report')."""
        try:
            time.sleep(0.2) # Allow window to settle
            hwnds = []
            def cb(h, r):
                t = win32gui.GetWindow_Text(h).lower() if hasattr(win32gui, 'GetWindow_Text') else win32gui.GetWindowText(h).lower()
                if "mosaic" in t and ("reporting" in t or "info hub" in t): r.append(h)
            win32gui.EnumWindows(cb, hwnds)
            
            if not hwnds: 
                log_trace("Targeting: Mosaic window not found for handle scan.")
                return
            
            app = Application(backend="uia").connect(handle=hwnds[0])
            win = app.window(handle=hwnds[0])
            
            # 1. Find the Anchor: Text or Button named with label_name
            anchor = None
            try:
                # We look for the exact name (e.g. 'Transcript' or 'Final Report')
                all_elems = win.descendants()
                for el in all_elems:
                    name = el.element_info.name
                    if name == label_name:
                        anchor = el
                        log_trace(f"Targeting: Found '{label_name}' anchor at {el.rectangle()}")
                        break
            except: pass

            # 2. Find Candidates: Large Group or Edit pads
            candidates = []
            try:
                potential_pads = win.descendants(control_type="Group") + win.descendants(control_type="Edit")
                for p in potential_pads:
                    try:
                        r = p.rectangle()
                        height = r.bottom - r.top
                        if height > 200:
                            candidates.append({'e': p, 'rect': r, 'left': r.left, 'top': r.top})
                    except: pass
            except: pass

            if not candidates:
                log_trace("Targeting: No major pads found (>200px height).")
                return

            target_rect = None
            if anchor:
                anchor_rect = anchor.rectangle()
                # Find the pad that is horizontally aligned with anchor and directly below it
                best_pad = None
                min_dist_y = 9999
                
                for c in candidates:
                    pad_rect = c['rect']
                    # Look for pads starting near or to the right of anchor's left edge
                    # (In Mosaic, pads often extend further right than labels)
                    if abs(pad_rect.left - anchor_rect.left) < 60 or (pad_rect.left <= anchor_rect.left <= pad_rect.right):
                        dist_y = pad_rect.top - anchor_rect.bottom
                        if 0 <= dist_y < min_dist_y:
                            min_dist_y = dist_y
                            best_pad = c
                
                if best_pad:
                    target_rect = best_pad['rect']
                    log_trace(f"Targeting: Found pad via {label_name} Anchor. Rect={target_rect}")

            # 3. Fallback: Leftmost tall pad (only for Transcript)
            if not target_rect and label_name == "Transcript":
                candidates.sort(key=lambda x: x['left'])
                target_rect = candidates[0]['rect']
                log_trace(f"Targeting: Fallback to leftmost pad for Transcript. Rect={target_rect}")

            if target_rect:
                # 1. Initial Target: Bottom-Right offset
                target_x = target_rect.right - 15
                target_y = target_rect.bottom - 15
                
                # 2. Screen Bounds Check (Handle Scrollable Containers / Off-screen)
                try:
                    # Get monitor info for the target point
                    h_monitor = win32api.MonitorFromPoint((target_x, target_y), win32con.MONITOR_DEFAULTTONEAREST)
                    mon_info = win32api.GetMonitorInfo(h_monitor)
                    work_area = mon_info['Work'] # (left, top, right, bottom) excluding taskbar
                    
                    # Clamp Y to work area bottom (minus safety margin)
                    safe_bottom = work_area[3] - 15
                    if target_y > safe_bottom:
                        log_trace(f"Targeting: Point {target_y} is off-screen/below taskbar. Clamping to {safe_bottom}.")
                        target_y = safe_bottom
                        
                    # Also clamp X if somehow off-screen
                    safe_right = work_area[2] - 15
                    if target_x > safe_right:
                         target_x = safe_right
                         
                except Exception as e:
                    log_trace(f"Targeting Bounds Check Failed: {e}")

                # Physical click to force focus
                curr_x, curr_y = pyautogui.position()
                pyautogui.click(target_x, target_y)
                pyautogui.moveTo(curr_x, curr_y) # Restore mouse
                log_trace(f"Targeting Success: Clicked {label_name} pad at {target_x},{target_y}")
            else:
                log_trace(f"Targeting Failed: Could not find pad for {label_name}")
                
        except Exception as e:
            log_trace(f"Targeting Exception: {str(e)}")

    def _activate_mosaic_forcefully(self):
        """Highly aggressive window activation (SwitchToThisWindow) to punch through overlays."""
        mosaic_hwnd = 0
        def _find_cb(h, _):
            nonlocal mosaic_hwnd
            if mosaic_hwnd: return
            if not win32gui.IsWindowVisible(h): return
            t = win32gui.GetWindowText(h).lower()
            if "mosaic" in t and ("reporting" in t or "info hub" in t):
                mosaic_hwnd = h
        win32gui.EnumWindows(_find_cb, None)

        if mosaic_hwnd:
            try:
                # 1. Attach Thread Input (Unlock Foreground)
                cur_t = ctypes.windll.kernel32.GetCurrentThreadId()
                tgt_t = ctypes.windll.user32.GetWindowThreadProcessId(mosaic_hwnd, None)
                ctypes.windll.user32.AttachThreadInput(cur_t, tgt_t, True)
                
                # 2. SwitchToThisWindow (Deep/Aggressive Switch)
                # True = Alt-Tab behavior (unminimizes and brings to front)
                ctypes.windll.user32.SwitchToThisWindow(mosaic_hwnd, True)
                
                # 3. BringWindowToTop
                ctypes.windll.user32.BringWindowToTop(mosaic_hwnd)
                
                # 4. SetForegroundWindow
                win32gui.SetForegroundWindow(mosaic_hwnd)
                
                # Detach
                ctypes.windll.user32.AttachThreadInput(cur_t, tgt_t, False)
                
                # Double-tap: Confirm focus with gentle method (User Request)
                # This ensures input focus is truly set even if SwitchToThisWindow acted weirdly.
                self._activate_mosaic(target_label=None)
                
                return True
            except Exception as e:
                log_trace(f"Force Activate Failed: {e}")
        return False

    def _activate_inteleviewer(self):
        """Robust activation of InteleViewer window."""
        return self._activate_window_by_title(["inteleviewer"])

    def _activate_window_by_title(self, mandatory, optional=None, timeout=2.0):
        """Helper to find and activate a window (AHK WinActivate-style).
        
        Key behaviors:
        1. If already foreground, return immediately (preserves internal focus)
        2. Uses AttachThreadInput + SetForegroundWindow (like AHK)
        3. Falls back to Alt key trick if initial attempts fail (like AHK's 7th attempt)
        """
        try:
            def enum_cb(hwnd, results):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = win32gui.GetWindowText(hwnd).lower()
                if not all(k.lower() in title for k in mandatory):
                    return
                if optional and not any(k.lower() in title for k in optional):
                    return
                results.append(hwnd)
            
            hwnds = []
            win32gui.EnumWindows(enum_cb, hwnds)
            if not hwnds:
                return False
                
            hwnd = hwnds[0]
            
            # Early exit if already active (preserves internal focus)
            if win32gui.GetForegroundWindow() == hwnd:
                return True
            
            # Get thread IDs for AttachThreadInput
            current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
            target_thread = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
            
            # Attempt 1: Standard activation with thread attachment
            attached = False
            if current_thread != target_thread:
                attached = ctypes.windll.user32.AttachThreadInput(current_thread, target_thread, True)
            
            try:
                # Restore if minimized
                placement = win32gui.GetWindowPlacement(hwnd)
                if placement[1] == win32con.SW_SHOWMINIMIZED:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                # Bring to top and set foreground
                ctypes.windll.user32.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                
            finally:
                if attached:
                    ctypes.windll.user32.AttachThreadInput(current_thread, target_thread, False)
            
            # Check if it worked
            time.sleep(0.05)
            if win32gui.GetForegroundWindow() == hwnd:
                return True
            
            # Attempt 2: Alt key trick (like AHK's fallback)
            # Pressing and releasing Alt allows SetForegroundWindow to work
            VK_MENU = 0x12
            KEYEVENTF_KEYUP = 0x0002
            ctypes.windll.user32.keybd_event(VK_MENU, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
            time.sleep(0.02)
            win32gui.SetForegroundWindow(hwnd)
            
            # Active wait
            start = time.time()
            while time.time() - start < timeout:
                if win32gui.GetForegroundWindow() == hwnd:
                    return True
                time.sleep(0.02)
            
            log_trace(f"WinActivate: Timeout waiting for {mandatory} to become foreground")
            return win32gui.GetForegroundWindow() == hwnd
            
        except Exception as e:
            log_trace(f"Window Activation Failed for {mandatory}: {e}")
        return False



    def perform_beep(self):
        """Action triggered by Record Button - toggles state locally for speed."""
        # Optimistic toggle for immediate feedback
        self.dictation_active = not self.dictation_active
        self.update_indicator_state()
        
        log_trace(f"Dictation Toggled: {'ON' if self.dictation_active else 'OFF'}")

        # Trigger a background check of the actual visual state
        # with a delay to allow UI to update (User requested 2x dictation_pause)
        # DEBOUNCE LOGIC: Only the LATEST beep triggers the actual scan.
        delay = (self.dictation_pause_ms / 1000.0) * 2.0
        if delay < 1.0: delay = 1.0 
        
        self.sync_check_token += 1
        current_token = self.sync_check_token
        
        def delayed_sync(token):
            time.sleep(delay)
            # If a newer beep occurred during sleep, ABORT.
            if self.sync_check_token != token:
                return
            
            # Now acquire lock just for the scan (prevent overlap execution)
            if self.sync_lock.acquire(blocking=False):
                try:
                    self.check_sync_state(token)
                finally:
                    self.sync_lock.release()

        threading.Thread(target=delayed_sync, args=(current_token,), daemon=True).start()
        
        # Modernized Beep Logic (Respects Granular Settings & Volume)
        should_play = self.start_beep_enabled if self.dictation_active else self.stop_beep_enabled
        if should_play:
            if self.dictation_active:
                # Dictation STARTING (High Pitch) - Custom delay
                delay = self.dictation_pause_ms / 1000.0
                if delay > 0:
                    time.sleep(delay)
                play_beep(1000, 200, self.start_beep_volume)
            else:
                # Dictation STOPPED (Low Pitch)
                play_beep(500, 200, self.stop_beep_volume)
    def perform_get_prior(self):
        """Native Get Prior: Extract and format prior study information from Mosaic."""
        log_trace("Executing: Get Prior (Native)")
        self.root.after(0, lambda: self.status_toast("Extracting Prior...", 1000))
        self.is_user_active = True
        
        try:
            # 1. Grab selected text (Logic: IV needs Ctrl+Shift+R, others Ctrl+c)
            old_clip = pyperclip.paste()
            pyperclip.copy("")
            
            # Detect active window
            foreground_hwnd = win32gui.GetForegroundWindow()
            active_window_text = ""
            try:
                active_window_text = win32gui.GetWindowText(foreground_hwnd).lower()
            except: pass
            
            log_trace(f"Active Window: {active_window_text}")
            
            # If not in InteleViewer, try to FIND it and alert user or just fail
            if "inteleviewer" not in active_window_text:
                log_trace(f"InteleViewer not active, skipping. Active: {active_window_text}")
                self.root.after(0, lambda: self.status_toast("InteleViewer must be active!", 2000))
                return
                
            log_trace(f"InteleViewer detected, sending '{self.iv_report_hotkey}'")
            
            # Ensure physical keys are released (Critical for when triggered via Keyboard Shortcut)
            time.sleep(0.3) 
            for key in ['ctrl', 'shift', 'alt']:
                pyautogui.keyUp(key)
            
            # Parse configured hotkey (e.g. "ctrl+shift+r" -> ['ctrl', 'shift', 'r'])
            # "v" -> ['v']
            hotkey_str = self.iv_report_hotkey.lower()
            if not hotkey_str: hotkey_str = "v" # Fallback
            
            keys_to_send = [k.strip() for k in hotkey_str.split('+')]
            pyautogui.hotkey(*keys_to_send)
            time.sleep(0.1)
            
            log_trace("Sending Ctrl+c to copy prior text")
            pyperclip.copy("")
            pyautogui.hotkey('ctrl', 'c')
            
            # Wait up to 1.5 seconds for clipboard content (snappy polling)
            raw_text = ""
            for i in range(15):
                raw_text = pyperclip.paste()
                if raw_text and len(raw_text.strip()) >= 5:
                    break
                time.sleep(0.1)

            log_trace(f"Clipboard Content Grabbed (Len: {len(raw_text)})")

            if not raw_text or len(raw_text.strip()) < 5:
                # One last try if it's slow?
                time.sleep(0.2)
                raw_text = pyperclip.paste()
                log_trace(f"Retry Clipboard Content (Len: {len(raw_text)})")

            if not raw_text or len(raw_text.strip()) < 5:
                log_trace(f"Raw Text too short: '{raw_text[:50]}'")
                pyperclip.copy(old_clip)
                self.root.after(0, lambda: self.status_toast("Nothing selected!", 1500))
                return

            # 2. Parse Date and Time
            prior_date = ""
            prior_time_formatted = ""
            
            # PhraseSearch := "i)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).*?(19[0-9][0-9]|20[0-9][0-9])"
            date_match = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).*?(19\d{2}|20\d{2})', raw_text, re.I)
            
            # TimezonePattern := "i)(\d{1,2}:\d{2}:\d{2}\s+(?:E|C|M|P|AK?|H)[SD]T)"
            time_match = re.search(r'(\d{1,2}:\d{2}:\d{2}\s+(?:E|C|M|P|AK?|H)[SD]T)', raw_text, re.I)
            
            if date_match and time_match:
                # Build M/D/YYYY from date_match content
                month_map = {"JAN":1,"FEB":2,"MAR":3,"APR":4,"MAY":5,"JUN":6,"JUL":7,"AUG":8,"SEP":9,"OCT":10,"NOV":11,"DEC":12}
                raw_date_str = date_match.group(0)
                year = date_match.group(2)
                
                month_num = 0
                for m, n in month_map.items():
                    if m in raw_date_str.upper():
                        month_num = n
                        break
                
                # Extract day (1 or 2 digits followed by space or colon)
                day_match = re.search(r'(\d{1,2})\s+\d', raw_date_str)
                day = day_match.group(1) if day_match else "1"
                prior_date = f"{month_num}/{int(day)}/{year}"
                
                # Format time (remove seconds)
                prior_time_formatted = re.sub(r'(\d{1,2}:\d{2}):\d{2}', r'\1', time_match.group(1))

            # 3. Handle Status Flags
            prior_original = raw_text
            for flag in ["IN_PROGRESS", "NO_HL7_ORDER", "UNKNOWN", "SIGNED"]:
                prior_original = prior_original.replace("\t" + flag, " SIGNXED")
            
            prior_images = ""
            if "\tNO_IMAGES" in prior_original:
                prior_images = "No Prior Images. "

            # 4. Modality Specific Processing
            prior_descript1 = ""
            
            # --- ULTRASOUND ---
            if "\tUS" in prior_original:
                desc_match = re.search(r'US.*SIGNXED', prior_original, re.I)
                if desc_match:
                    desc = desc_match.group(0)
                    desc = re.sub(r'^US', '', desc, count=1, flags=re.I)
                    desc = desc.replace(" SIGNXED", "").replace(" abd.", " abdomen.").strip()
                    prior_descript1 = self._reorder_laterality(desc.lower())
                    if " with and without" in prior_descript1:
                        prior_descript1 = re.sub(r'(\s+)(with and without)', r' ultrasound\2', prior_descript1, flags=re.I)
                    elif " without" in prior_descript1:
                        prior_descript1 = re.sub(r'(\s+)(without)', r' ultrasound\2', prior_descript1, flags=re.I)
                    elif " with" in prior_descript1:
                        prior_descript1 = re.sub(r'(\s+)(with)', r' ultrasound\2', prior_descript1, flags=re.I)
                    else:
                        prior_descript1 += " ultrasound"

            # --- MR ---
            elif "\tMR" in prior_original:
                desc_match = re.search(r'MR.*SIGNXED', prior_original, re.I)
                if desc_match:
                    desc = desc_match.group(0)
                    desc = re.sub(r'^MR', '', desc, count=1, flags=re.I)
                    desc = desc.replace(" SIGNXED", "").strip()
                    desc = desc.replace(" + ", " and ").replace(" W/O", " without").replace(" W/", " with")
                    desc = desc.replace(" W WO", " with and without").replace(" WO", " without").replace(" IV ", " ")
                    
                    prior_descript1 = self._reorder_laterality(desc.lower())
                    
                    modifier_found = False
                    if " mra" in prior_descript1 or " mrv" in prior_descript1:
                        modifier_found = True
                    elif " angiography" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(angiography)', r' MR \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " venography" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(venography)', r' MR \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " with and without" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(with and without)', r' MR \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " without" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(without)', r' MR \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " with" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(with)', r' MR \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    
                    if not modifier_found:
                        prior_descript1 += " MR"

            # --- NUCLEAR MEDICINE ---
            elif "\tNM" in prior_original:
                desc_match = re.search(r'NM.*SIGNXED', prior_original, re.I)
                if desc_match:
                    desc = desc_match.group(0).replace("NM", "nuclear medicine").replace(" SIGNXED", "").strip().lower()
                    prior_descript1 = self._reorder_laterality(desc)

            # --- X-RAY / RADIOGRAPH ---
            elif "\tXR" in prior_original or "\tCR" in prior_original or "\tX-ray" in prior_original:
                if "\tXR" in prior_original:
                    match = re.search(r'XR.*SIGNXED', prior_original, re.I)
                elif "\tCR" in prior_original:
                    match = re.search(r'CR.*SIGNXED', prior_original, re.I)
                else: # X-ray
                    match = re.search(r'X-ray.*SIGNXED', prior_original, re.I)

                if match:
                    desc = match.group(0)
                    desc = re.sub(r'^(XR|CR|X-ray, OT|X-ray)\t?', '', desc, flags=re.I)
                    desc = desc.replace(" SIGNXED", "").strip()
                    prior_descript1 = self._process_radiograph(desc.lower())

            # --- CT ---
            elif "\tCT" in prior_original:
                prior_original_tmp = re.sub(r'Oct', 'OcX', prior_original, flags=re.I)
                desc_match = re.search(r'CT.*SIGNXED', prior_original_tmp, re.I)
                if desc_match:
                    desc = desc_match.group(0)
                    desc = desc.replace("CTA - ", "CTA ")
                    desc = desc.replace("CTA", "CTAPLACEHOLDER")
                    desc = re.sub(r'^CT(\s|$)', r'\1', desc, flags=re.I)
                    desc = desc.replace("CTAPLACEHOLDER", "CTA").replace(" SIGNXED", "").strip()
                    
                    desc = desc.replace(" + ", " and ").replace("+", " and ").replace(" imags", "").replace("Head Or Brain", "brain")
                    desc = desc.replace(" W/CONTRST INCL W/O", " with and without contrast").replace(" W/O", " without")
                    desc = desc.replace(" W/ ", " with ").replace(" W/", " with ")
                    desc = re.sub(r'\s+W\s+', ' with ', desc, flags=re.I)
                    desc = desc.replace(" W WO", " with and without").replace(" WO", " without").replace(" IV ", " ")
                    
                    desc = desc.replace("ab pe", "abdomen and pelvis").replace("abd & pelvis", "abdomen and pelvis")
                    desc = desc.replace("abd/pelvis", "abdomen and pelvis").replace(" abd pel ", " abdomen and pelvis ")
                    desc = desc.replace("abdomen/pelvis", "abdomen and pelvis").replace("chest/abdomen/pelvis", "chest, abdomen, and pelvis")
                    desc = desc.replace("Thorax", "chest").replace("thorax", "chest").replace("P.E", "PE").replace("p.e", "PE")
                    desc = re.sub(r'\s+protocol\s*$', '', desc, flags=re.I)
                    
                    prior_descript1 = self._reorder_laterality(desc.lower())
                    
                    modifier_found = False
                    has_cta = False
                    cta_moved_from_start = False
                    
                    if re.match(r'^cta\s+', prior_descript1, re.I):
                        prior_descript1 = re.sub(r'^cta\s+', '', prior_descript1, count=1, flags=re.I)
                        has_cta = True
                        cta_moved_from_start = True
                        modifier_found = True
                    elif " cta" in prior_descript1:
                        has_cta = True
                        modifier_found = True
                    elif " angiography" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(angiography)', r' CT \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " with and without" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(with and without)', r' CT \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " without" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(without)', r' CT \1', prior_descript1, flags=re.I)
                        modifier_found = True
                    elif " with" in prior_descript1:
                        prior_descript1 = re.sub(r'\s+(with)', r' CT \1', prior_descript1, flags=re.I)
                        modifier_found = True

                    if cta_moved_from_start:
                        if re.match(r'^(right|left|bilateral)\s+', prior_descript1, re.I):
                            prior_descript1 = re.sub(r'^((?:right|left|bilateral)\s+\w+)\s*', r'\1 CTA ', prior_descript1, flags=re.I)
                        else:
                            prior_descript1 = re.sub(r'^(\w+)\s*', r'\1 CTA ', prior_descript1, flags=re.I)
                    
                    if not modifier_found and not has_cta:
                        prior_descript1 += " CT"

            # 5. Final Formatting
            prior_descript1 = re.sub(r'\bcta\b', 'CTA', prior_descript1, flags=re.I)
            prior_descript1 = re.sub(r'\bct\b', 'CT', prior_descript1, flags=re.I)
            for term in ["MRA", "MRV", "MRI", "MR", "PA", "PE"]:
                prior_descript1 = re.sub(rf'\b{term}\b', term, prior_descript1, flags=re.I)
            
            include_time = False
            if prior_date and prior_time_formatted:
                try:
                    p_dt = datetime.datetime.strptime(prior_date, "%m/%d/%Y")
                    now = datetime.datetime.now()
                    diff_days = (now - p_dt).days
                    if 0 <= diff_days <= 2:
                        include_time = True
                except: pass

            final_text = f" COMPARISON: {prior_date} {' ' + prior_time_formatted if include_time else ''} {prior_descript1}. {prior_images}"
            final_text = re.sub(r' {2,}', ' ', final_text).replace(" .", ".").strip()

            # 6. Paste to Mosaic
            log_trace(f"Final Comparison Text: {final_text}")
            
            # Activate Mosaic but don't click into a specific box (preserves user's focus)
            self._activate_mosaic(target_label=None)
            
            # Final Paste
            pyperclip.copy(final_text)
            time.sleep(0.1) 
            pyautogui.hotkey('ctrl', 'v')
            log_trace(f"Final Paste Sent: {final_text}")
            self.root.after(0, lambda: self.status_toast("Prior Formatted!", 1500))

        except Exception as e:
            log_trace(f"Native Get Prior failed: {e}")
            import traceback
            trace_msg = traceback.format_exc()
            log_trace(f"Traceback: {trace_msg}")
            self.root.after(0, lambda: self.status_toast("Extraction Error!", 2000))
        finally:
            self.is_user_active = False

    def _reorder_laterality(self, text):
        """Helper for Get Prior: shoulder right -> right shoulder."""
        if re.match(r'^(right|left|bilateral)\b', text, re.I):
            return text
        match = re.search(r'\b(right|left|bilateral)\b', text, re.I)
        if match:
            term = match.group(1)
            cleaned = re.sub(rf'\s*\b{term}\b\s*', ' ', text, count=1, flags=re.I)
            return f"{term} {cleaned.strip()}".strip()
        return text

    def _process_radiograph(self, text):
        """Helper for Get Prior: view expansions and radiograph insertion."""
        text = text.replace(" vw", " view(s)").replace(" 2v", " PA and lateral")
        text = text.replace(" pa lat", " PA and lateral").replace(" (kub)", "").strip()
        
        text = self._reorder_laterality(text)
        
        modifier_found = False
        # Numeric view patterns
        if re.search(r'\d+\s+or\s+more\s+radiograph\s+views?', text, re.I):
            text = re.sub(r'(\d+\s+or\s+more)\s+radiograph\s+(views?)', r'radiograph \1 \2', text, flags=re.I)
            modifier_found = True
        elif re.search(r'\s\d+\s*view', text):
            text = re.sub(r'(\s)(\d+\s*view)', r'\1radiograph \2', text)
            modifier_found = True
        elif " pa and lateral" in text:
            text = text.replace(" pa and lateral", " radiograph PA and lateral")
            modifier_found = True
        elif " view" in text:
            text = text.replace(" view", " radiograph view")
            modifier_found = True
            
        if not modifier_found:
            text += " radiograph"
        return text

    def perform_show_report(self):
        """Show Mosaic Report Popup."""
        # Toggle Logic: If window exists, close it
        if hasattr(self, 'report_popup') and self.report_popup and self.report_popup.winfo_exists():
            self.report_popup.destroy()
            self.report_popup = None
            return

        print("Executing: Show Report")
        self.status_toast("Generating Report...", 1000)
        
        raw_text = get_report_text_content(find_mosaic_report_window()) # Uses uia
        formatted = format_report_text(raw_text)
        
        # Display logic needs to handle non-blocking popup
        # We can implement a simple popup class similar to MosaicReportApp inside here or launch separate
        # Given complexity, let's create a Toplevel here similar to the standalone app
        
        # self.report_window code here...
        self.report_popup = self.show_report_popup(formatted)
        
    def show_report_popup(self, text):
        popup = tk.Toplevel(self.root)
        popup.withdraw() # Hide immediately
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.config(bg='#666666')
        
        # Saved pos and width
        popup_width = self.settings.get("report_popup_width", 750)
        x = self.settings.get("report_popup_x", 300)
        y = self.settings.get("report_popup_y", 300)
        
        # Main Frame with grey border effect
        main = tk.Frame(popup, bg='#666666')
        main.pack(fill='both', expand=True)
        
        # Drag bar / Header area
        header_bar = tk.Frame(main, bg='black', height=20)
        header_bar.pack(fill='x', padx=1, pady=(1, 0))
        
        drag_handle = tk.Label(header_bar, text="⋯", font=("Segoe UI", 10), bg='black', fg='#666', cursor='fleur')
        drag_handle.pack(side=tk.RIGHT, padx=5)
        
        # Resizable Container logic
        content_frame = tk.Frame(main, bg='black')
        content_frame.pack(fill='both', expand=True, padx=1, pady=(0, 1))
        
        # Resize grip on the right
        grip = tk.Frame(main, bg='#444444', width=5, cursor='sb_h_double_arrow')
        grip.place(relx=1.0, rely=0, relheight=1.0, anchor='ne', x=-1, y=1)

        inner = tk.Frame(content_frame, bg='black')
        inner.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Wrap length needs to be dynamic for all labels
        popup.labels = []

        def update_wraplength(new_w):
            for lbl in popup.labels:
                lbl.config(wraplength=new_w - 30)

        # Drag logic
        def start_move(event):
            popup.x = event.x
            popup.y = event.y
        def do_move(event):
            deltax = event.x - popup.x
            deltay = event.y - popup.y
            nx = popup.winfo_x() + deltax
            ny = popup.winfo_y() + deltay
            popup.geometry(f"+{nx}+{ny}")
        def stop_move(event):
            self.settings["report_popup_x"] = popup.winfo_x()
            self.settings["report_popup_y"] = popup.winfo_y()
            self.save_settings_file()
            
        header_bar.bind("<Button-1>", start_move)
        header_bar.bind("<B1-Motion>", do_move)
        header_bar.bind("<ButtonRelease-1>", stop_move)
        drag_handle.bind("<Button-1>", start_move)
        drag_handle.bind("<B1-Motion>", do_move)
        drag_handle.bind("<ButtonRelease-1>", stop_move)

        # Resize logic
        def start_resize(event):
            popup.start_x = event.x_root
            popup.start_w = popup.winfo_width()
        def do_resize(event):
            dw = event.x_root - popup.start_x
            nw = max(400, popup.start_w + dw)
            popup.geometry(f"{nw}x{popup.winfo_height()}")
            update_wraplength(nw)
        def stop_resize(event):
            self.settings["report_popup_width"] = popup.winfo_width()
            self.save_settings_file()

        grip.bind("<Button-1>", start_resize)
        grip.bind("<B1-Motion>", do_resize)
        grip.bind("<ButtonRelease-1>", stop_resize)
        
        # Content
        lines = text.split('\n')
        for line in lines:
            if not line.strip():
                tk.Frame(inner, bg='black', height=8).pack(fill='x')
                continue
            
            fg = "#E0E0E0" if "IMPRESSION" in line or "FINDINGS" in line else "#AAAAAA"
            font_style = ("Segoe UI", 11, "bold") if fg=="#E0E0E0" else ("Segoe UI", 11)
            
            lbl = tk.Label(inner, text=line, font=font_style, fg=fg, bg='black', anchor='w', justify='left', wraplength=popup_width-30)
            lbl.pack(fill='x')
            lbl.bind("<Button-1>", lambda e: popup.destroy()) # Click content to close
            popup.labels.append(lbl)
            
        # Position with fixed width but AUTO height (achieved by not providing height or setting it to 1)
        # Using f"+{x}+{y}" allows total auto-sizing, but we want to honor the saved width.
        # Tkinter will expand the window to fit the packed content even if geometry is just +X+Y
        # so we set the initial geometry to just position and rely on wraplength for width.
        popup.geometry(f"+{x}+{y}")
        popup.deiconify() 
        popup.update() 
        return popup

    def perform_capture_series(self):
        """Capture Yellow Box -> OCR -> Paste."""
        log_trace("Executing: Capture Series")
        
        if not pil_lib or not np_lib:
            self.root.after(0, lambda: self.status_toast("Missing dependencies: Pillow or Numpy. Capture disabled.", 3000))
            return
            
        self.root.after(0, lambda: self.status_toast("Capturing...", 1000))
        self.is_user_active = True
        
        try:
            # 1. Find Yellow Box
            # 1. Determine Capture Region (Fallback Logic)
            log_trace("Determining capture region...")
            target = self.ocr_engine.find_yellow_target()
            bbox = None
            
            if target:
                tx, ty = target
                bbox = (tx, ty, tx+300, ty+300)
                log_trace(f"Source: Yellow Box at {target}")
            else:
                # Fallback: InteleViewer Header or Mouse
                iv_win = self._activate_window_by_title(["inteleviewer"]) # Just check existence/activate
                # Note: _activate returns bool, we need the actual window rect if possible.
                # Since we don't have pywinauto window object readily available here without re-scan,
                # let's use the mouse fallback as the primary alternative which is most robust manually.
                
                mx, my = pyautogui.position()
                bbox = (mx-250, my-50, mx+250, my+150)
                log_trace(f"Source: Mouse (Fallback) at {mx},{my}")
                self.root.after(0, lambda: self.status_toast("Yellow Box not found, using Mouse Pos", 1000))

            # 2. Capture
            log_trace("Grabbing screenshot of yellow box...")
            raw_img = ImageGrab.grab(bbox=bbox, all_screens=True)
            img_np = np.array(raw_img)
            
            # 3. OCR (Running in background)
            log_trace(f"Running OCR on captured image... Shape: {img_np.shape}")
            raw_text = self.ocr_engine.run_rapid(img_np)
            final_str = extract_series_image_numbers(raw_text)
            
            if final_str:
                log_trace(f"OCR Result Formatted: {final_str}")
                
                # 4. Activate Mosaic JUST-IN-TIME (Forceful to overcome any overlays)
                self._activate_mosaic_forcefully()
                time.sleep(0.2) # Let window settle
                
                # 5. Paste
                pyperclip.copy(final_str)
                time.sleep(0.05)
                pyautogui.hotkey('ctrl', 'v')
                log_trace(f"Final Paste Sent: {final_str}")
                self.root.after(0, lambda: self.status_toast(f"Pasted: {final_str}", 2000))
            else:
                preview = raw_text[:50] if raw_text else "None"
                log_trace(f"OCR Failed or no numbers: '{preview}'")
                self.root.after(0, lambda: self.status_toast("OCR Failed / No numbers", 2000))
                
        except Exception as e:
            log_trace(f"Capture Series Failed: {e}")
            import traceback
            log_trace(f"Traceback: {traceback.format_exc()}")
            self.root.after(0, lambda: self.status_toast("Capture Error!", 2000))
        finally:
            self.is_user_active = False
            time.sleep(0.2)

    def perform_toggle_record(self, desired_state=None):
        """Toggle recording in Mosaic via Alt+R. 
        desired_state: None=Toggle, True=Start, False=Stop
        Uses raw keybd_event to send Alt+R without any UIA interaction that could change focus.
        """
        log_trace(f"Executing: Toggle Recording (Alt+R, desired={desired_state})")
        
        # Smart Check (Registry Lookup)
        if desired_state is not None:
             current_real = self.is_dictation_active_registry()
             if current_real == desired_state:
                 log_trace(f"Smart Toggle: Already {'Recording' if desired_state else 'Stopped'}. Skipping action.")
                 # Sync internal tracking just in case
                 if self.dictation_active != current_real:
                      self.dictation_active = current_real
                      self.update_indicator_state()
                 # Still play beep to acknowledge user's button press
                 # Still play beep to acknowledge user's button press
                 should_play = self.start_beep_enabled if desired_state else self.stop_beep_enabled
                 if should_play:
                     freq = 1000 if desired_state else 500
                     # Add delay for START beep
                     delay = (self.dictation_pause_ms / 1000.0) if (desired_state and self.dictation_pause_ms > 0) else 0
                     def play_delayed_beep():
                         if delay > 0:
                             time.sleep(delay)
                         play_beep(freq, 200)
                     threading.Thread(target=play_delayed_beep, daemon=True).start()
                 return # EXIT EARLY - No keystroke needed

        # Pre-check reality for smarter toggle (know what we are toggling FROM)
        pre_toggle_real_state = None
        if desired_state is None:
             pre_toggle_real_state = self.is_dictation_active_registry()

        try:
            # Find the Mosaic window
            hwnds = []
            def cb(h, r):
                if not win32gui.IsWindowVisible(h):
                    return
                t = win32gui.GetWindowText(h).lower()
                if "mosaic" in t and ("reporting" in t or "info hub" in t):
                    r.append(h)
            win32gui.EnumWindows(cb, hwnds)
            
            if not hwnds:
                self.status_toast("Mosaic Not Found", 1000)
                return
            
            hwnd = hwnds[0]
            
            # Check if Mosaic is already foreground
            already_active = (win32gui.GetForegroundWindow() == hwnd)
            
            # Only activate if not already foreground
            if not already_active:
                if not self._activate_mosaic(target_label=None):
                    self.status_toast("Mosaic Not Found", 1000)
                    return
                # Brief pause for activation
                time.sleep(0.1)
            
            # Send Alt+R using raw keybd_event (lowest level, no UIA)
            # Virtual key codes: VK_MENU (Alt) = 0x12, VK_R = 0x52
            VK_MENU = 0x12
            VK_R = 0x52
            KEYEVENTF_KEYUP = 0x0002
            
            # Press Alt
            ctypes.windll.user32.keybd_event(VK_MENU, 0, 0, 0)
            # Press R
            ctypes.windll.user32.keybd_event(VK_R, 0, 0, 0)
            # Release R
            ctypes.windll.user32.keybd_event(VK_R, 0, KEYEVENTF_KEYUP, 0)
            # Release Alt
            ctypes.windll.user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
            
            log_trace("Sent Alt+R to Mosaic via keybd_event")
            
            # Update internal state
            if desired_state is not None:
                self.dictation_active = desired_state
            else:
                if pre_toggle_real_state is not None:
                    # We toggled FROM `pre_toggle_real_state`, so we are now NOT that.
                    log_trace(f"Smart Toggle: Registry said {pre_toggle_real_state} -> Now {not pre_toggle_real_state}")
                    self.dictation_active = not pre_toggle_real_state
                else:
                    self.dictation_active = not self.dictation_active
            self.update_indicator_state()
            
            # Play beep feedback
            should_play = self.start_beep_enabled if self.dictation_active else self.stop_beep_enabled
            if should_play:
                freq = 1000 if self.dictation_active else 500
                # Add delay for START beep
                delay = (self.dictation_pause_ms / 1000.0) if (self.dictation_active and self.dictation_pause_ms > 0) else 0
                def play_delayed_beep():
                    if delay > 0:
                        time.sleep(delay)
                    play_beep(freq, 200)
                threading.Thread(target=play_delayed_beep, daemon=True).start()
                
        except Exception as e:
            log_trace(f"Toggle Record Failed: {e}")
            import traceback
            log_trace(traceback.format_exc())


    def perform_process_report(self, inject_always=True):

        """Trigger 'Process Report' in Mosaic via Alt+P."""
        log_trace("Executing: Process Report (Alt+P)")
        
        try:
            # Save current mouse position
            old_x, old_y = pyautogui.position()
            
            # FORCEFUL ACTIVATION (Process Report Specific)
            if self._activate_mosaic_forcefully():
                time.sleep(0.3)
                
                did_auto_stop = False
                
                # Auto-Stop Safety
                if self.auto_stop_dictation and self.dictation_active:
                    log_trace("Auto-Stop Triggered")
                    # 1. Send Alt+R (Raw keybd_event for consistency)
                    VK_MENU = 0x12
                    VK_R = 0x52
                    KEYEVENTF_KEYUP = 0x0002
                    
                    ctypes.windll.user32.keybd_event(VK_MENU, 0, 0, 0)
                    ctypes.windll.user32.keybd_event(VK_R, 0, 0, 0)
                    ctypes.windll.user32.keybd_event(VK_R, 0, KEYEVENTF_KEYUP, 0)
                    ctypes.windll.user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
                    
                    # 2. Update Internal State
                    self.dictation_active = False 
                    self.update_indicator_state()
                    
                    # 2.5 Play Sound (Feedback)
                    log_trace(f"Attempting Auto-Stop Beep (Enabled={self.stop_beep_enabled})")
                    if self.stop_beep_enabled:
                         threading.Thread(target=lambda: play_beep(500, 200, self.stop_beep_volume), daemon=True).start()

                    # 3. Pause
                    time.sleep(0.5) 
                    did_auto_stop = True

                # TRIGGER LOGIC:
                # 1. If we Auto-Stopped, we MUST send Alt+P to recover the interrupted action.
                # 2. If 'inject_always' is True (Explicit Hotkey/Mapping), we MUST send Alt+P.
                # 3. If neither (Fallback Button w/o Auto-Stop), we DO NOTHING (Native handles it).
                should_send_p = did_auto_stop or inject_always

                if should_send_p:
                    # Robust Alt+P
                    pyautogui.keyDown('alt')
                    pyautogui.press('p')
                    pyautogui.keyUp('alt')
                    log_trace("Sent Alt+P")
                else:
                    log_trace("Process Report: Skipped Injection (Native Fallback)")
                
                # Restore mouse
                pyautogui.moveTo(old_x, old_y)
            else:
                self.status_toast("Mosaic Not Found", 1000)
                
        except Exception as e:
            log_trace(f"Process Report Failed: {e}")

    def perform_sign_report(self):
        """Trigger 'Sign Report' in Mosaic via Alt+F."""
        log_trace("Executing: Sign Report (Alt+F)")
        try:
            # Save current mouse position
            old_x, old_y = pyautogui.position()
            
            # Activate Mosaic
            if self._activate_mosaic(target_label=None):
                time.sleep(0.3)
                
                # Robust Alt+F
                pyautogui.keyDown('alt')
                pyautogui.press('f')
                pyautogui.keyUp('alt')
                
                log_trace("Sent Alt+F to Mosaic")
                
                # Restore mouse
                pyautogui.moveTo(old_x, old_y)
            else:
                self.status_toast("Mosaic Not Found", 1000)
                
        except Exception as e:
            log_trace(f"Sign Report Failed: {e}")

    def is_dictation_active_uia(self):
        """Check if Mosaic is currently recording by looking for recording indicators (UIA method - slow).
        
        LEGACY/BACKUP: This method scans the UI tree which can be slow and miss indicators.
        Prefer is_dictation_active_registry() for faster, more reliable detection.
        """
        try:
            desktop = Desktop(backend="uia")
            
            # BROAD SEARCH: Iterate ALL visible windows
            # The 'Microphone recording' indicator might be a top-level Pane or Window
            # that is distinct from the main 'SlimHub' window in the UIA tree.
            for window in desktop.windows(visible_only=True):
                try:
                    name = window.element_info.name or ""
                    
                    # 1. Direct Window Name Match
                    if "Microphone recording" in name:
                        return True
                        
                    # 2. Check Children if it's a Mosaic/SlimHub window
                    if "Mosaic" in name or "SlimHub" in name:
                        try:
                            # We use descendants() because 'child_window' is not available on Wrapper objects
                            # Depth of 14 should capture the recording pane or button (User Request)
                            children = window.descendants(depth=14)
                            for child in children:
                                c_name = child.element_info.name or ""
                                if "Microphone recording" in c_name:
                                    return True
                                if "accessing your microphone" in c_name:
                                    return True
                        except: pass
                except: pass
            
            return False
        except Exception as e:
            print(f"Error checking dictation state (UIA): {e}")
            return False

    def is_dictation_active_registry(self):
        """Check if microphone is in use via Windows registry (faster method).
        
        Uses Windows CapabilityAccessManager to detect if msedgewebview2.exe (Mosaic's WebView)
        has an active microphone session. This is much faster than UIA scanning.
        
        Returns:
            True: Microphone is actively recording
            False: Microphone is not recording
            None: Registry check failed (caller should fallback to UIA)
        """
        import winreg
        
        # Target executables used by Mosaic for WebView2
        # Updated to include Packaged App names found in diagnostics (MosaicInfoHub)
        TARGET_EXES = ["msedgewebview2.exe", "webviewhost.exe", "mosaic"]
        
        # Registry paths where Windows tracks microphone access
        MIC_ROOTS = [
            r"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged",
            r"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
        ]
        
        try:
            seen_target = False
            for root_path in MIC_ROOTS:
                try:
                    root_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, root_path)
                except FileNotFoundError:
                    continue
                    
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(root_key, i)
                        for exe in TARGET_EXES:
                            if exe.lower() in subkey_name.lower():
                                seen_target = True
                                try:
                                    subkey = winreg.OpenKey(root_key, subkey_name)
                                    start_val, _ = winreg.QueryValueEx(subkey, "LastUsedTimeStart")
                                    stop_val, _ = winreg.QueryValueEx(subkey, "LastUsedTimeStop")
                                    winreg.CloseKey(subkey)
                                    
                                    # Active if start != 0 and stop == 0
                                    if start_val != 0 and stop_val == 0:
                                        winreg.CloseKey(root_key)
                                        return True
                                except (FileNotFoundError, OSError):
                                    pass
                        i += 1
                    except OSError:
                        break
                winreg.CloseKey(root_key)
            
            # If we saw a target exe entry but it's not active, return False
            # If we never saw any target, return None to trigger UIA fallback
            if seen_target:
                return False
            else:
                return None  # No registry data for target - fallback to UIA
                
        except Exception as e:
            log_trace(f"Registry dictation check failed: {e}")
            return None  # Signals fallback to UIA

    def is_dictation_active(self):
        """Check if Mosaic is currently recording.
        
        Primary detection: Uses fast registry-based check.
        Fallback: Uses slower UIA-based check if registry fails.
        """
        result = self.is_dictation_active_registry()
        if result is None:
            # Registry check failed or no data - fallback to UIA
            return self.is_dictation_active_uia()
        return result

    def check_sync_state(self, expected_token=None):
        """
        Periodic check to ensure local state matches reality.
        expected_token: If provided, ensures that we only update state if 
        no NEW sync request (beep) has occurred since this check started.
        """
        real_state = self.is_dictation_active()
        
        # VALIDATE TOKEN: If a new beep happened during is_dictation_active(), 
        # self.sync_check_token will have incremented. We must abort to avoid 
        # overwriting the user's latest action with old data.
        if expected_token is not None and self.sync_check_token != expected_token:
            log_trace(f"Sync Aborted: Token Mismatch ({expected_token} vs {self.sync_check_token})")
            return

        # STRICTLY CONSERVATIVE SYNC:
        # We only correct the state if we POSITIVELY detect recording (Real=True).
        # We DO NOT correct if we detect 'Stopped' (Real=False), because UI scans can miss the icon (False Negative).
        # A False Negative here would reset the user's state to OFF while they are dictating, causing a double-start beep.
        
        real_state = self.is_dictation_active()
        
        if real_state and not self.dictation_active:
            print(f"SYNC FIX: Local=False, Real=True. Detected recording icon. Correcting to TRUE.")
            self.dictation_active = True
            self.root.after(0, self.update_indicator_state)
        # else: Ignore. Trust local state.

    def perform_clario_scrape(self):
        """Helper to handle the tiered Clario scrape."""
        window = self.find_clario_window()
        if not window: 
            log_trace("Clario window NOT found.")
            return None
        
        # Ensure window is focused/active to populate UIA tree
        try:
            log_trace("Focusing Clario window before scrape...")
            window.set_focus()
            time.sleep(0.5)
        except:
            pass
        
        # FAST PATH: Check if Note Dialog is open
        note_dialog_text = self.get_note_dialog_text(window)
        if note_dialog_text:
            log_trace("SUCCESS: Found open Note Dialog.")
            return note_dialog_text
        
        # NORMAL PATH: Exhaustive Search on whole window
        notes = self.get_exam_note_elements(window)
        
        if notes:
            # Heuristic: longest match is usually the one with the actual note text
            best_note = max(notes, key=len)
            log_trace(f"SUCCESS: Returning best match (len={len(best_note)})")
            return best_note
            
        return None

    def get_note_dialog_text(self, window):
        """Check if the Note Dialog is open and extract text from it."""
        try:
            # Look for the note dialog by automation_id
            for elem in window.descendants():
                try:
                    auto_id = elem.element_info.automation_id or ""
                    if auto_id == "content_patient_note_dialog_Main":
                        # Found the dialog! Now find the note text field
                        for child in elem.descendants():
                            try:
                                child_id = child.element_info.automation_id or ""
                                # noteFieldMessage-XXXX-inputEl contains the actual text
                                if "noteFieldMessage" in child_id and "inputEl" in child_id:
                                    note_text = child.element_info.name or child.window_text() or ""
                                    if note_text:
                                        return note_text.strip()
                            except:
                                pass
                        # Dialog found but no text field - return None to fall back
                        return None
                except:
                    pass
        except Exception as e:
            print(f"Error checking for note dialog: {e}")
        return None

    def start_async_action(self, target_func, progress_msg=None):
        """Start a cancellable background action with optional progress verification."""
        self.action_token += 1
        token = self.action_token
        self.is_user_active = True
        
        # Run in thread
        threading.Thread(target=target_func, args=(token,), daemon=True).start()
        
        # Start Progress Loop if requested
        if progress_msg:
            def _progress_loop():
                # Check if this specific action is still the active one
                if self.action_token == token and self.is_user_active:
                    self.status_toast(progress_msg, 5000)
                    # Schedules next check slightly before this toast expires to ensure continuity
                    self.root.after(4500, _progress_loop)
            
            # Start the loop after the initial toast (assumed 5s) would be fading
            self.root.after(4500, _progress_loop)

    def scrape_and_debug_async(self, token):
        """Debug workflow with pop-up result (Threaded)."""
        try:
            self.root.after(0, lambda: self.status_toast("Debug Scrape Started...", 5000))
            
            # Check Cancel
            if self.action_token != token: return

            raw_note = self.perform_clario_scrape()
            
            # Check Cancel
            if self.action_token != token: return

            if raw_note:
                formatted_note = self.format_note(raw_note)
                
                # Check Cancel
                if self.action_token != token: return

                self.root.after(0, lambda: self.show_results_window(formatted_note, raw_note))
                self.root.after(0, lambda: self.status_toast("Debug Window Opened", 1000))
            else:
                self.root.after(0, lambda: self.status_toast("No Clario Note Found", 2000))
                self.root.after(0, lambda: messagebox.showwarning("Not Found", "No 'EXAM NOTE' elements found in Clario."))
        except Exception as e:
            log_trace(f"Debug Scrape Failed: {e}")
            self.root.after(0, lambda: self.status_toast("Debug Error!", 2000))
        finally:
            # Only unset active if we are still the current token
            if self.action_token == token:
                self.is_user_active = False

    def scrape_and_process_async(self, token):
        """Main scraping workflow called by button press (Threaded)."""
        try:
            self.root.after(0, lambda: self.status_toast("Scraping Clario...", 5000))
            
            # Check Cancel
            if self.action_token != token: return

            # We do this in series. UIA has internal locks that can often make 
            # parallel searches MUCH slower than sequential ones.
            raw_note = self.perform_clario_scrape()
            
            # Check Cancel
            if self.action_token != token: return
            
            if raw_note:
                self.root.after(0, lambda: self.status_toast("Note found! Locating Mosaic...", 2000))
                
                # UIA search for editor
                editor = self.find_mosaic_editor()
                
                formatted_note = self.format_note(raw_note)
                
                # Check Cancel
                if self.action_token != token: return
                
                # Copy to clipboard
                pyperclip.copy(formatted_note)
                
                # AUTOMATIC PASTE INTO MOSAIC
                if editor:
                    try:
                        # Save current mouse position
                        old_x, old_y = pyautogui.position()
                        
                        # Use centralized activation - DO NOT force focus to Final Report anymore
                        self._activate_mosaic(target_label=None)
                        
                        # Final Paste
                        pyautogui.hotkey('ctrl', 'v')
                        
                        # Restore mouse position
                        pyautogui.moveTo(old_x, old_y)
                        
                        self.root.after(0, lambda: self.status_toast("Successfully Pasted to Mosaic!", 1500))
                    except Exception as e:
                        print(f"Paste error: {e}")
                        self.root.after(0, lambda: self.status_toast("Focus failed, result copied to clipboard.", 2000))
                else:
                    self.root.after(0, lambda: self.status_toast("Mosaic editor not found, result copied.", 2000))
                    
                print(f"Scraped & Formatted: {formatted_note}")
            else:
                print("No EXAM NOTE elements found in Clario.")
                self.root.after(0, lambda: self.status_toast("No Note Found", 2000))
                self.root.after(0, lambda: messagebox.showwarning("Not Found", "No 'EXAM NOTE' elements found in Clario."))
        except Exception as e:
            log_trace(f"Scrape Process Failed: {e}")
            self.root.after(0, lambda: self.status_toast("Scrape Error!", 2000))
        finally:
            if self.action_token == token:
                self.is_user_active = False

    def trigger_action_by_name(self, action, is_fallback=False):
        """Centralized action dispatcher used by HID, Keyboard, and WinMsg."""
        log_trace(f"Triggering Action: {action} (Fallback={is_fallback})")
        
        if action == ACTION_NONE:
            return
        elif action == ACTION_BEEP:
            # Run in thread to not block
            threading.Thread(target=self.perform_beep, daemon=True).start()
            # Debounce is handled by caller usually, but harmless here
        elif action == ACTION_GET_PRIOR:
            # Run in thread to not block hotkey/main loop
            threading.Thread(target=self.perform_get_prior, daemon=True).start()
        elif action == ACTION_SCRAPE:
            self.start_async_action(self.scrape_and_process_async, progress_msg="Still scraping Clario...")
        elif action == ACTION_DEBUG:
            self.start_async_action(self.scrape_and_debug_async, progress_msg="Still debugging...")
        elif action == ACTION_SHOW_REPORT:
            self.perform_show_report()
        elif action == ACTION_CAPTURE_SERIES:
            self.perform_capture_series()
        elif action == ACTION_TOGGLE_RECORD:
            self.perform_toggle_record()
        elif action == ACTION_PROCESS_REPORT:
            # Pass explicit flag: If NOT fallback, we inject always.
            self.perform_process_report(inject_always=not is_fallback)
        elif action == ACTION_SIGN_REPORT:
            self.perform_sign_report()

    def setup_hotkeys(self):
        """Register global hotkeys via keyboard library."""
        if not keyboard: return
        
        try:
            keyboard.unhook_all() # Clear previous
            
            for action_name, config in self.action_mappings.items():
                hotkey = config.get("hotkey", "")
                if hotkey and hotkey.strip():
                    log_trace(f"Registering Hotkey: '{hotkey}' -> {action_name}")
                    try:
                        # IMPORTANT: Do NOT use suppress=True - it blocks keyboard input
                        # until the handler returns. Instead, let the key pass through and
                        # run our action asynchronously.
                        def on_press(a=action_name):
                            try:
                                log_trace(f"HOTKEY DETECTED: {a}")
                                # Run action in a separate thread to not block the keyboard hook
                                def async_action():
                                    try:
                                        self.root.after(0, lambda: self.status_toast(f"Recognized: {a}", 5000))
                                        self.root.after(0, lambda: self.trigger_action_by_name(a))
                                    except Exception as ex:
                                        log_trace(f"HOTKEY ASYNC ERROR: {ex}")
                                threading.Thread(target=async_action, daemon=True).start()
                            except Exception as ex:
                                log_trace(f"HOTKEY ERROR: {ex}")
                        
                        keyboard.add_hotkey(hotkey, on_press, suppress=False)

                    except Exception as e:
                        log_trace(f"Failed to register hotkey '{hotkey}': {e}")
            
            # Watchdog: Schedule a refresh to prevent hook timeouts/drops
            # Only schedule if we haven't already (to avoid exponential loops if called recursively)
            if not hasattr(self, '_watchdog_scheduled'):
                 self._watchdog_scheduled = True
                 self.root.after(30000, self.hotkey_watchdog)

        except Exception as e:
            log_trace(f"Error setting up hotkeys: {e}")


    def hotkey_watchdog(self):
        """Periodically refresh hotkeys to prevent silent hook drops by Windows."""
        if not self.running: return
        try:
            # log_trace("Maintenance: Refreshing Hotkey Hooks...")
            # We don't verify full re-setup, just ensure hooks are alive.
            # actually, calling setup_hotkeys unhooks all and re-adds. 
            # This is "nuclear" but effective for dropped hooks.
            self.setup_hotkeys() 
        except: pass
        
        self.root.after(30000, self.hotkey_watchdog)

    def background_listener(self):
        """Listens for PowerMic button presses via HID."""
        if not hid:
            print("HID library not found. PowerMic buttons will not work.")
            return

        device = None
        last_sync_time = time.time()
        
        last_rec_down = False # Track Record Button state for PTT
        
        while self.running:
            # Sync Check every 30 seconds (Hybrid Approach)
            if (time.time() - last_sync_time > 30) and not self.is_user_active:
                threading.Thread(target=self.check_sync_state, daemon=True).start()
                last_sync_time = time.time()
            
            if not device:
                found_any = False
                for d in hid.enumerate():
                    if d['vendor_id'] in VENDOR_IDS:
                        log_trace(f"Attempting to connect to HID: {d.get('product_string', 'PowerMic')} at {d['path']}")
                        try:
                            device = hid.device()
                            device.open_path(d['path'])
                            device.set_nonblocking(1)
                            # Notify user of successful connection
                            msg = f"Connected to {d.get('product_string', 'PowerMic')}"
                            log_trace(f"HID Connection Successful: {msg}")
                            self.root.after(0, lambda m=msg: self.status_toast(m, 3000))
                            found_any = True
                            last_rec_down = False # Reset state on reconnect
                            break
                        except Exception as e:
                            log_trace(f"HID Connection Failed: {e}")
                            device = None
                
                if not found_any:
                    time.sleep(5)
                    continue

            try:
                data = device.read(64)
                if data:
                    btn_data = list(data[:3])
                    
                    if self.dead_man_switch:
                        # Check Byte 1, Bit 2 (Mask 0x04)
                        rec_down = (btn_data[1] & 0x04) != 0
                        
                        # PRESS: Ensure START (skip if a stop retry is in progress)
                        if rec_down and not last_rec_down:
                            if self.ptt_busy:
                                log_trace("PTT: Press ignored (busy with stop retry)")
                            else:
                                log_trace("PTT: Record Down -> Enforce START")
                                self.perform_toggle_record(desired_state=True)
                             
                        # RELEASE: Ensure STOP (with retry, skip if already busy)
                        elif not rec_down and last_rec_down:
                            if self.ptt_busy:
                                log_trace("PTT: Release ignored (already retrying)")
                            else:
                                def ptt_stop_with_retry():
                                    self.ptt_busy = True
                                    try:
                                        for attempt in range(10):
                                            real_state = self.is_dictation_active_registry()
                                            if not real_state:
                                                log_trace(f"PTT Stop: Success (Registry confirmed inactive)")
                                                if self.dictation_active:
                                                    self.dictation_active = False
                                                    self.root.after(0, self.update_indicator_state)
                                                break
                                            
                                            log_trace(f"PTT Stop: Attempt {attempt+1}/10 (Registry active)")
                                            self._activate_mosaic_forcefully()
                                            time.sleep(0.1)
                                            self.perform_toggle_record(desired_state=False)
                                            time.sleep(1.0)
                                        else:
                                            log_trace("PTT Stop: FAILED after 10 attempts!")
                                    finally:
                                        self.ptt_busy = False

                                threading.Thread(target=ptt_stop_with_retry, daemon=True).start()
                        
                        last_rec_down = rec_down
                        
                        # CRITICAL: Mask Record button so standard handler doesn't fire blind toggle
                        btn_data[1] = btn_data[1] & ~0x04

                    # Match Button Name
                    matched_button = None
                    for name, code in BUTTON_DEFINITIONS.items():
                        if btn_data == code:
                            matched_button = name
                            break
                    
                    if matched_button:
                        # Lookup Action from Unified Mappings
                        action_to_trigger = ACTION_NONE
                        for act_name, config in self.action_mappings.items():
                            if config.get("mic_button") == matched_button:
                                action_to_trigger = act_name
                                break
                        
                        # FALLBACK: If no mapping found (or mapped to None), check Hardcoded Defaults
                        # The UI claims these are "Hardcoded", so we honor that promise.
                        if action_to_trigger == ACTION_NONE:
                            if matched_button == "Skip Back":
                                action_to_trigger = ACTION_PROCESS_REPORT
                            elif matched_button == "Checkmark":
                                action_to_trigger = ACTION_SIGN_REPORT
                            elif matched_button == "Record Button":
                                # For PTT mode, Record Button MUST trigger toggle
                                # For regular mode, default is often Beep (via mapping)
                                if self.dead_man_switch:
                                    action_to_trigger = ACTION_TOGGLE_RECORD
                        
                        if action_to_trigger != ACTION_NONE:
                            log_trace(f"HID Trigger: {matched_button} -> {action_to_trigger}")
                            
                            # Determine if this was a fallback trigger
                            is_fallback = (action_to_trigger not in [self.action_mappings.get(a, {}).get("mic_button") for a in [ACTION_SCRAPE, ACTION_DEBUG]])
                            # Simplification: If we hit the loop break via normal mapping, act_name is set.
                            # If we hit fallback, we manually set action_to_trigger.
                            # Better check:
                            is_smart_fallback = False
                            if matched_button == "Skip Back" and action_to_trigger == ACTION_PROCESS_REPORT:
                                # Check if it was REALLY mapped explicitly?
                                explicit_map = self.action_mappings.get(ACTION_PROCESS_REPORT, {}).get("mic_button")
                                if explicit_map != "Skip Back":
                                    is_smart_fallback = True
                            
                            # Debounce logic specifically for physical buttons if needed
                            if action_to_trigger == ACTION_BEEP:
                                self.trigger_action_by_name(action_to_trigger, is_fallback=is_smart_fallback)
                                time.sleep(0.8) 
                            elif action_to_trigger in [ACTION_SCRAPE, ACTION_DEBUG]:
                                self.trigger_action_by_name(action_to_trigger, is_fallback=is_smart_fallback)
                                time.sleep(1.0)
                            else:
                                self.trigger_action_by_name(action_to_trigger, is_fallback=is_smart_fallback)
                                
                time.sleep(0.01)
            except:
                device = None
                time.sleep(2)

    def setup_message_listener(self):
        """Create a hidden window to listen for custom Windows messages (Stable win32gui version)."""
        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_TRIGGER_SCRAPE:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_SCRAPE))
                return 0
            elif msg == WM_TRIGGER_DEBUG:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_DEBUG))
                return 0
            elif msg == WM_TRIGGER_BEEP:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_BEEP))
                return 0
            elif msg == WM_TRIGGER_SHOW_REPORT:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_SHOW_REPORT))
                return 0
            elif msg == WM_TRIGGER_CAPTURE_SERIES:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_CAPTURE_SERIES))
                return 0
            elif msg == WM_TRIGGER_GET_PRIOR:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_GET_PRIOR))
                return 0
            elif msg == WM_TRIGGER_TOGGLE_RECORD:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_TOGGLE_RECORD))
                return 0
            elif msg == WM_TRIGGER_PROCESS_REPORT:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_PROCESS_REPORT))
                return 0
            elif msg == WM_TRIGGER_SIGN_REPORT:
                self.root.after(0, lambda: self.trigger_action_by_name(ACTION_SIGN_REPORT))
                return 0
            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = wnd_proc
        wc.lpszClassName = "MosaicToolsMessageWindow"
        wc.hInstance = win32gui.GetModuleHandle(None)
        
        try:
            class_atom = win32gui.RegisterClass(wc)
            self.msg_hwnd = win32gui.CreateWindow(
                class_atom, "Mosaic Tools Service", 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None
            )
            log_trace(f"Windows message listener established on HWND: {self.msg_hwnd}")
            win32gui.PumpMessages()
        except Exception as e:
            print(f"Could not establish message listener: {e}")

if __name__ == "__main__":
    try:
        log_trace("--- NEW SESSION ---")
        app = MosaicToolsApp()
    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        log_trace(f"TOP LEVEL CRASH: {e}\n{err_msg}")
        with open("mosaic_crash_log.txt", "w") as f:
            f.write(f"CRASH AT STARTUP: {e}\n")
            f.write(err_msg)
        # Show message box if possible
        try:
            import tkinter.messagebox as mb
            mb.showerror("MosaicTools Startup Error", f"Application failed to start:\n\n{e}\n\nCheck mosaic_crash_log.txt for details.")
        except: pass
