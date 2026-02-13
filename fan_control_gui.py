#!/usr/bin/env python3
"""
Fan Control GUI — Lenovo Yoga Pro 9i Gen 9 (2024)

A dark-themed tkinter GUI for controlling fan speed via the EC mailbox.
Uses fan_control.py as the backend.

MUST BE RUN AS ADMINISTRATOR.
"""

import tkinter as tk
from tkinter import messagebox
import threading
import time
import math
import sys
import os
import json
import atexit
import tempfile
import subprocess
import ctypes
import ctypes.wintypes as wintypes

from PIL import Image, ImageDraw
import pystray

# Add script directory to path so we can import fan_control
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fan_control import WinRing0, ECMailbox, is_admin

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# ==============================================================================
# Theme / Colors
# ==============================================================================

COLORS = {
    "bg_dark":      "#0f0f1a",
    "bg_panel":     "#1a1a2e",
    "bg_card":      "#16213e",
    "bg_input":     "#0a0a1a",
    "accent":       "#00d2ff",
    "accent_dim":   "#0077aa",
    "accent_glow":  "#00e5ff",
    "text":         "#e0e0e0",
    "text_dim":     "#808090",
    "text_bright":  "#ffffff",
    "success":      "#00e676",
    "warning":      "#ffab00",
    "danger":       "#ff1744",
    "preset_quiet": "#4a6741",
    "preset_bal":   "#4a5a6a",
    "preset_perf":  "#6a5a3a",
    "preset_full":  "#6a3a3a",
    "slider_trough":"#0a0e1a",
    "arc_bg":       "#1e2a40",
}

FONT_FAMILY = "Segoe UI"

# Config file lives next to the exe/script (not in _MEIPASS temp dir)
if getattr(sys, 'frozen', False):
    _app_dir = os.path.dirname(sys.executable)
else:
    _app_dir = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_app_dir, "fan_config.json")

# Safety limits
MIN_FAN_SPEED = 18   # Values 1-17 cause pulsing; 0 (off) is fine
SAFE_MAX = 48        # EC's normal maximum; above this requires confirmation
_above_safe_confirmed = False  # User has acknowledged going above SAFE_MAX

def clamp_fan_speed(val):
    """Clamp to valid range: 0 or 18-100 (1-17 causes fan pulsing)."""
    val = max(0, min(100, int(val)))
    if 1 <= val < MIN_FAN_SPEED:
        return MIN_FAN_SPEED
    return val

# Built-in presets: (name, fan%, button_color)
BUILTIN_PRESETS = [
    ("OFF",      0,   "#3a3a4a"),
    ("Min",      18,  COLORS["preset_quiet"]),
    ("Med",      22,  COLORS["preset_bal"]),
    ("Med-High", 30,  COLORS["preset_perf"]),
    ("High",     48,  COLORS["preset_full"]),
]


# ==============================================================================
# Arc Gauge Widget
# ==============================================================================

class ArcGauge(tk.Canvas):
    """A circular arc gauge widget showing fan speed percentage."""

    def __init__(self, parent, label="Fan", size=160, **kwargs):
        super().__init__(parent, width=size, height=size + 30,
                         bg=COLORS["bg_panel"], highlightthickness=0, **kwargs)
        self.size = size
        self.label = label
        self.value = 0
        self._draw()

    def _draw(self):
        self.delete("all")
        cx, cy = self.size // 2, self.size // 2
        r = self.size // 2 - 15
        pad = 15

        # Background arc (270 degrees, from 135 to -135)
        self.create_arc(pad, pad, self.size - pad, self.size - pad,
                        start=225, extent=-270, style="arc",
                        outline=COLORS["arc_bg"], width=10)

        # Value arc
        extent = -(self.value / 100.0) * 270
        if self.value > 0:
            color = self._get_color_for_value(self.value)
            self.create_arc(pad, pad, self.size - pad, self.size - pad,
                            start=225, extent=extent, style="arc",
                            outline=color, width=10)

        # Center text - percentage
        self.create_text(cx, cy - 5,
                         text=f"{self.value}%",
                         font=(FONT_FAMILY, 24, "bold"),
                         fill=COLORS["text_bright"])

        # Label below gauge
        self.create_text(cx, self.size + 10,
                         text=self.label,
                         font=(FONT_FAMILY, 11),
                         fill=COLORS["text_dim"])

    def _get_color_for_value(self, val):
        """Gradient from cyan (low) to orange (mid) to red (high)."""
        if val <= 40:
            return COLORS["accent"]
        elif val <= 70:
            return COLORS["warning"]
        else:
            return COLORS["danger"]

    def set_value(self, value):
        value = max(0, min(100, int(value)))
        if value != self.value:
            self.value = value
            self._draw()


# ==============================================================================
# Styled Slider Widget
# ==============================================================================

class FanSlider(tk.Frame):
    """A styled horizontal slider for fan speed control."""

    def __init__(self, parent, label="Fan 1", on_change=None, **kwargs):
        super().__init__(parent, bg=COLORS["bg_panel"], **kwargs)
        self.on_change = on_change
        self._dragging = False

        # Label row
        top = tk.Frame(self, bg=COLORS["bg_panel"])
        top.pack(fill="x", padx=10, pady=(5, 0))

        self.label = tk.Label(top, text=label,
                              font=(FONT_FAMILY, 11),
                              fg=COLORS["text_dim"], bg=COLORS["bg_panel"])
        self.label.pack(side="left")

        self.value_label = tk.Label(top, text="50%",
                                    font=(FONT_FAMILY, 11, "bold"),
                                    fg=COLORS["accent"], bg=COLORS["bg_panel"])
        self.value_label.pack(side="right")

        # Scale widget
        self.scale = tk.Scale(self, from_=0, to=100, orient="horizontal",
                              length=280, sliderlength=20,
                              bg=COLORS["bg_panel"],
                              fg=COLORS["text"],
                              troughcolor=COLORS["slider_trough"],
                              activebackground=COLORS["accent"],
                              highlightthickness=0,
                              borderwidth=0,
                              showvalue=False,
                              command=self._on_slide)
        self.scale.set(30)
        self.scale.pack(fill="x", padx=10, pady=(0, 5))

        # Bind release event for actual command sending
        self.scale.bind("<ButtonRelease-1>", self._on_release)
        self.scale.bind("<ButtonPress-1>", self._on_press)

    def _on_press(self, event):
        self._dragging = True

    def _on_slide(self, value):
        v = clamp_fan_speed(int(float(value)))
        # Snap the slider to the clamped value if different
        if v != int(float(value)):
            self.scale.set(v)
        self.value_label.config(text=f"{v}%")
        # Visual warning when above safe max
        if v > SAFE_MAX:
            self.value_label.config(fg=COLORS["danger"])
        else:
            self.value_label.config(fg=COLORS["accent"])

    def _on_release(self, event):
        self._dragging = False
        if self.on_change:
            self.on_change(self.get_value())

    def get_value(self):
        return int(self.scale.get())

    def set_value(self, val, trigger=False):
        val = clamp_fan_speed(val)
        self.scale.set(val)
        self.value_label.config(text=f"{val}%")
        if val > SAFE_MAX:
            self.value_label.config(fg=COLORS["danger"])
        else:
            self.value_label.config(fg=COLORS["accent"])
        if trigger and self.on_change:
            self.on_change(val)


# ==============================================================================
# Main Application
# ==============================================================================

class FanControlApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Fan Control — Yoga Pro 9i")
        self.root.configure(bg=COLORS["bg_dark"])
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._minimize_to_tray)

        # State
        self.ec = None
        self.driver = None
        self.connected = False
        self.hold_active = False
        self.auto_mode = True
        self.running = True
        self._hold_thread = None
        self._monitor_thread = None
        self._tray_icon = None

        # Generate and set app icon
        self._app_icon = self._create_fan_icon(256)
        self._tray_icon_image = self._create_fan_icon(64)
        self._set_window_icon()

        self._build_ui()
        self._connect()
        self._load_config()
        self._start_monitor()
        self._setup_tray()
        self._setup_power_events()
        self._setup_shutdown_handler()
        self._ensure_startup_safety_task()

        atexit.register(self._safety_restore)

    # ── Shutdown Detection ───────────────────────────────────────────────

    def _setup_shutdown_handler(self):
        """Create a hidden native Win32 window to catch WM_QUERYENDSESSION.

        Tkinter's withdrawn window may not receive shutdown broadcasts.
        A raw hidden Win32 window with its own message pump is reliable.
        """
        try:
            user32 = ctypes.windll.user32
            kernel32_dll = ctypes.windll.kernel32

            # 64-bit safe WNDPROC type
            WNDPROC = ctypes.WINFUNCTYPE(
                ctypes.c_longlong,       # LRESULT
                ctypes.c_void_p,         # HWND
                wintypes.UINT,           # MSG
                ctypes.c_ulonglong,      # WPARAM
                ctypes.c_longlong,       # LPARAM
            )

            WM_QUERYENDSESSION = 0x0011
            WM_ENDSESSION = 0x0016

            # Set DefWindowProcW types
            user32.DefWindowProcW.restype = ctypes.c_longlong
            user32.DefWindowProcW.argtypes = [
                ctypes.c_void_p, wintypes.UINT,
                ctypes.c_ulonglong, ctypes.c_longlong,
            ]

            def _shutdown_wndproc(hwnd, msg, wparam, lparam):
                if msg in (WM_QUERYENDSESSION, WM_ENDSESSION):
                    # System is shutting down — restore auto fan control NOW
                    self.running = False
                    self.hold_active = False
                    if self.connected:
                        try:
                            self.ec.restore_auto()
                        except Exception:
                            pass
                        try:
                            self.driver.close()
                        except Exception:
                            pass
                        try:
                            self.driver.stop_driver()
                        except Exception:
                            pass
                        self.connected = False
                    return 1 if msg == WM_QUERYENDSESSION else 0
                return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

            # Must prevent garbage collection
            self._shutdown_wndproc_ref = WNDPROC(_shutdown_wndproc)

            # WNDCLASSEXW structure
            class WNDCLASSEXW(ctypes.Structure):
                _fields_ = [
                    ("cbSize", wintypes.UINT),
                    ("style", wintypes.UINT),
                    ("lpfnWndProc", WNDPROC),
                    ("cbClsExtra", ctypes.c_int),
                    ("cbWndExtra", ctypes.c_int),
                    ("hInstance", ctypes.c_void_p),
                    ("hIcon", ctypes.c_void_p),
                    ("hCursor", ctypes.c_void_p),
                    ("hbrBackground", ctypes.c_void_p),
                    ("lpszMenuName", wintypes.LPCWSTR),
                    ("lpszClassName", wintypes.LPCWSTR),
                    ("hIconSm", ctypes.c_void_p),
                ]

            kernel32_dll.GetModuleHandleW.restype = ctypes.c_void_p
            kernel32_dll.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
            hInstance = kernel32_dll.GetModuleHandleW(None)

            wc = WNDCLASSEXW()
            wc.cbSize = ctypes.sizeof(WNDCLASSEXW)
            wc.lpfnWndProc = self._shutdown_wndproc_ref
            wc.hInstance = hInstance
            wc.lpszClassName = "FanControlShutdownWatcher"

            user32.RegisterClassExW.restype = wintypes.ATOM
            user32.RegisterClassExW.argtypes = [ctypes.POINTER(WNDCLASSEXW)]

            atom = user32.RegisterClassExW(ctypes.byref(wc))
            if not atom:
                return

            user32.CreateWindowExW.restype = ctypes.c_void_p
            user32.CreateWindowExW.argtypes = [
                wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                wintypes.DWORD,
                ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                ctypes.c_void_p, ctypes.c_void_p,
                ctypes.c_void_p, ctypes.c_void_p,
            ]

            # Create hidden top-level window (NOT message-only — must receive broadcasts)
            self._shutdown_hwnd = user32.CreateWindowExW(
                0, "FanControlShutdownWatcher", "FanCtrl Shutdown",
                0,  # not visible
                0, 0, 0, 0,
                None, None, hInstance, None,
            )
            if not self._shutdown_hwnd:
                return

            # Message pump in background thread
            def _shutdown_msg_pump():
                msg = wintypes.MSG()
                while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))

            threading.Thread(target=_shutdown_msg_pump, daemon=True).start()
        except Exception:
            pass  # Best effort

    # ── Icon Generation ─────────────────────────────────────────────────────

    @staticmethod
    def _create_fan_icon(size):
        """Generate a fan-blade icon programmatically."""
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        cx, cy = size // 2, size // 2
        r = size // 2 - 2

        # Background circle
        draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(15, 15, 26, 255))

        # Draw 3 fan blades
        blade_color = (0, 210, 255, 255)  # cyan accent
        for angle_offset in [0, 120, 240]:
            points = []
            for a in range(0, 80, 2):
                angle = math.radians(angle_offset + a)
                # Blade shape: starts thin at center, widens outward
                dist = r * 0.2 + (r * 0.7) * (a / 80.0)
                width = r * 0.08 + r * 0.15 * (a / 80.0)
                px = cx + dist * math.cos(angle)
                py = cy + dist * math.sin(angle)
                # Offset perpendicular for width
                nx = width * math.cos(angle + math.pi/2)
                ny = width * math.sin(angle + math.pi/2)
                points.append((px + nx, py + ny))

            # Return path
            for a in range(78, -1, -2):
                angle = math.radians(angle_offset + a)
                dist = r * 0.2 + (r * 0.7) * (a / 80.0)
                width = r * 0.08 + r * 0.15 * (a / 80.0)
                px = cx + dist * math.cos(angle)
                py = cy + dist * math.sin(angle)
                nx = width * math.cos(angle - math.pi/2)
                ny = width * math.sin(angle - math.pi/2)
                points.append((px + nx, py + ny))

            if len(points) >= 3:
                draw.polygon(points, fill=blade_color)

        # Center hub
        hub_r = int(r * 0.18)
        draw.ellipse([cx-hub_r, cy-hub_r, cx+hub_r, cy+hub_r],
                     fill=(0, 180, 220, 255), outline=(0, 229, 255, 255), width=2)

        return img

    def _set_window_icon(self):
        """Set the window/taskbar icon from the generated image."""
        try:
            self._icon_path = os.path.join(
                tempfile.gettempdir(), "fan_control_icon.ico"
            )
            self._app_icon.save(self._icon_path, format='ICO',
                                sizes=[(256, 256), (64, 64), (32, 32), (16, 16)])
            self.root.iconbitmap(self._icon_path)
        except Exception:
            pass

    # ── System Tray ───────────────────────────────────────────────────────

    def _setup_tray(self):
        """Create the system tray icon with a right-click menu."""
        menu = pystray.Menu(
            pystray.MenuItem("Show", self._show_from_tray, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit_from_tray),
        )
        self._tray_icon = pystray.Icon(
            "FanControl",
            self._tray_icon_image,
            "Fan Control — Yoga Pro 9i",
            menu
        )
        # Run tray icon in background thread
        threading.Thread(target=self._tray_icon.run, daemon=True).start()

    def _minimize_to_tray(self):
        """Hide window to tray — unless the system is shutting down."""
        # SM_SHUTTINGDOWN (0x2000) is TRUE during shutdown/restart
        SM_SHUTTINGDOWN = 0x2000
        if ctypes.windll.user32.GetSystemMetrics(SM_SHUTTINGDOWN):
            # System is shutting down — do a real close with auto-restore
            self._on_close()
            return
        self.root.withdraw()

    def _show_from_tray(self, icon=None, item=None):
        """Restore window from system tray."""
        self.root.after(0, self._restore_window)

    def _restore_window(self):
        """Restore and focus the window."""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _quit_from_tray(self, icon=None, item=None):
        """Actually quit the application from tray menu."""
        self.root.after(0, self._on_close)

    # ── Power Event Handling (Sleep/Resume) ──────────────────────────────

    def _setup_power_events(self):
        """Register for power suspend/resume via powrprof callback (no window needed)."""
        try:
            powrprof = ctypes.windll.powrprof

            # Callback type: ULONG callback(PVOID Context, ULONG Type, PVOID Setting)
            DEVICE_NOTIFY_CALLBACK_ROUTINE = ctypes.CFUNCTYPE(
                wintypes.ULONG,     # return
                ctypes.c_void_p,    # Context
                wintypes.ULONG,     # Type (PBT_APM* values)
                ctypes.c_void_p,    # Setting
            )

            class DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS(ctypes.Structure):
                _fields_ = [
                    ("Callback", DEVICE_NOTIFY_CALLBACK_ROUTINE),
                    ("Context", ctypes.c_void_p),
                ]

            powrprof.PowerRegisterSuspendResumeNotification.restype = wintypes.DWORD
            powrprof.PowerRegisterSuspendResumeNotification.argtypes = [
                wintypes.DWORD,
                ctypes.POINTER(DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS),
                ctypes.POINTER(wintypes.HANDLE),
            ]

            PBT_APMSUSPEND = 0x0004
            PBT_APMRESUMESUSPEND = 0x0007
            PBT_APMRESUMEAUTOMATIC = 0x0012
            DEVICE_NOTIFY_CALLBACK = 2

            def power_callback(context, event_type, setting):
                if event_type == PBT_APMSUSPEND:
                    self._on_suspend()
                elif event_type in (PBT_APMRESUMESUSPEND, PBT_APMRESUMEAUTOMATIC):
                    self._on_resume()
                return 0

            # Must keep references to prevent garbage collection
            self._power_callback_func = DEVICE_NOTIFY_CALLBACK_ROUTINE(power_callback)
            self._power_params = DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS()
            self._power_params.Callback = self._power_callback_func
            self._power_params.Context = None
            self._power_reg_handle = wintypes.HANDLE()

            result = powrprof.PowerRegisterSuspendResumeNotification(
                DEVICE_NOTIFY_CALLBACK,
                ctypes.byref(self._power_params),
                ctypes.byref(self._power_reg_handle),
            )
            # result == 0 means success (ERROR_SUCCESS)
        except Exception:
            pass  # Best effort — sleep handling is a nice-to-have

    def _on_suspend(self):
        """Called when system is about to sleep. Restore auto + stop driver."""
        # Restore auto fan control first
        if self.connected:
            try:
                self.ec.restore_auto()
            except Exception:
                pass

        # Close driver handle and stop the service
        if self.driver:
            try:
                self.driver.close()
            except Exception:
                pass
            try:
                self.driver.stop_driver()
            except Exception:
                pass
        self.connected = False

    def _on_resume(self):
        """Called when system resumes from sleep. Restart driver."""
        def _delayed_reconnect():
            time.sleep(3)  # Give hardware time to settle
            self.root.after(0, self._connect)
        threading.Thread(target=_delayed_reconnect, daemon=True).start()

    # ── Boot Safety ───────────────────────────────────────────────────

    def _ensure_startup_safety_task(self):
        """Install a Windows scheduled task to restore auto fan control on boot."""
        TASK_NAME = "FanControlAutoRestore"
        try:
            # Check if task already exists
            result = subprocess.run(
                ['schtasks', '/query', '/tn', TASK_NAME],
                capture_output=True, text=True
            )
            if result.returncode == 0:
                return  # Already installed

            # Get exe/script path for the task
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = sys.executable  # python.exe

            if getattr(sys, 'frozen', False):
                task_cmd = f'"{exe_path}" --startup-safety'
            else:
                script_path = os.path.abspath(__file__)
                task_cmd = f'"{exe_path}" "{script_path}" --startup-safety'

            # Create the scheduled task: runs at system startup as SYSTEM
            subprocess.run([
                'schtasks', '/create',
                '/tn', TASK_NAME,
                '/tr', task_cmd,
                '/sc', 'onstart',        # Run at system boot
                '/ru', 'SYSTEM',         # Run as SYSTEM (has admin rights)
                '/rl', 'HIGHEST',        # Highest privileges
                '/f',                    # Force (overwrite if exists)
            ], capture_output=True, text=True)
        except Exception:
            pass  # Non-critical — best effort

    # ── UI Construction ────────────────────────────────────────────────────

    def _build_ui(self):
        # Main container
        main = tk.Frame(self.root, bg=COLORS["bg_dark"], padx=15, pady=10)
        main.pack(fill="both", expand=True)

        # ── Status Bar ──
        status_frame = tk.Frame(main, bg=COLORS["bg_dark"])
        status_frame.pack(fill="x", pady=(0, 8))

        self.status_dot = tk.Canvas(status_frame, width=12, height=12,
                                     bg=COLORS["bg_dark"], highlightthickness=0)
        self.status_dot.pack(side="left", padx=(0, 6))
        self._draw_status_dot(False)

        self.status_label = tk.Label(status_frame, text="Connecting...",
                                     font=(FONT_FAMILY, 9),
                                     fg=COLORS["text_dim"], bg=COLORS["bg_dark"])
        self.status_label.pack(side="left")

        # CPU temp on right
        self.temp_label = tk.Label(status_frame, text="",
                                   font=(FONT_FAMILY, 9),
                                   fg=COLORS["text_dim"], bg=COLORS["bg_dark"])
        self.temp_label.pack(side="right")

        # ── Gauges ──
        gauge_frame = tk.Frame(main, bg=COLORS["bg_panel"])
        gauge_frame.pack(fill="x", pady=(0, 8))

        # Card-like background
        gauge_card = tk.Frame(gauge_frame, bg=COLORS["bg_panel"], padx=15, pady=10)
        gauge_card.pack(fill="x")

        self.gauge1 = ArcGauge(gauge_card, label="Fan 1", size=150)
        self.gauge1.pack(side="left", padx=(15, 10), pady=5)

        self.gauge2 = ArcGauge(gauge_card, label="Fan 2", size=150)
        self.gauge2.pack(side="right", padx=(10, 15), pady=5)

        # ── Sliders Card ──
        slider_card = tk.Frame(main, bg=COLORS["bg_panel"], padx=5, pady=8)
        slider_card.pack(fill="x", pady=(0, 8))

        # Link toggle
        link_frame = tk.Frame(slider_card, bg=COLORS["bg_panel"])
        link_frame.pack(fill="x", padx=10, pady=(2, 0))

        self.linked_var = tk.BooleanVar(value=True)
        self.link_cb = tk.Checkbutton(link_frame, text="🔗 Link fans together",
                                       variable=self.linked_var,
                                       font=(FONT_FAMILY, 9),
                                       fg=COLORS["text_dim"], bg=COLORS["bg_panel"],
                                       selectcolor=COLORS["bg_input"],
                                       activebackground=COLORS["bg_panel"],
                                       activeforeground=COLORS["text"])
        self.link_cb.pack(side="left")

        self.slider1 = FanSlider(slider_card, label="Fan 1",
                                  on_change=self._on_slider1_change)
        self.slider1.pack(fill="x")

        self.slider2 = FanSlider(slider_card, label="Fan 2",
                                  on_change=self._on_slider2_change)
        self.slider2.pack(fill="x")

        # ── Presets ──
        preset_card = tk.Frame(main, bg=COLORS["bg_panel"], padx=10, pady=10)
        preset_card.pack(fill="x", pady=(0, 8))

        preset_label = tk.Label(preset_card, text="PRESETS",
                                font=(FONT_FAMILY, 8, "bold"),
                                fg=COLORS["text_dim"], bg=COLORS["bg_panel"])
        preset_label.pack(pady=(0, 6))

        self.preset_frame = tk.Frame(preset_card, bg=COLORS["bg_panel"])
        self.preset_frame.pack(fill="x")

        # Will be populated by _rebuild_presets()
        self.custom_presets = []  # list of {"name": str, "speed": int}
        self._rebuild_presets()

        # ── Hold + Auto ──
        bottom_card = tk.Frame(main, bg=COLORS["bg_panel"], padx=10, pady=10)
        bottom_card.pack(fill="x", pady=(0, 5))

        controls = tk.Frame(bottom_card, bg=COLORS["bg_panel"])
        controls.pack()

        self.hold_var = tk.BooleanVar(value=False)
        self.hold_cb = tk.Checkbutton(controls, text="🔄 Hold Mode",
                                       variable=self.hold_var,
                                       font=(FONT_FAMILY, 10),
                                       fg=COLORS["text"], bg=COLORS["bg_panel"],
                                       selectcolor=COLORS["bg_input"],
                                       activebackground=COLORS["bg_panel"],
                                       activeforeground=COLORS["text"],
                                       command=self._toggle_hold)
        self.hold_cb.pack(side="left", padx=(0, 20))

        self.auto_btn = tk.Button(controls, text="↺ Auto Mode",
                                   font=(FONT_FAMILY, 11, "bold"),
                                   fg=COLORS["text_bright"],
                                   bg=COLORS["accent_dim"],
                                   activebackground=COLORS["accent"],
                                   activeforeground=COLORS["text_bright"],
                                   relief="flat", padx=20, pady=6,
                                   cursor="hand2",
                                   command=self._restore_auto)
        self.auto_btn.pack(side="left")
        self.auto_btn.bind("<Enter>", lambda e: self.auto_btn.config(
            bg=COLORS["accent"]))
        self.auto_btn.bind("<Leave>", lambda e: self.auto_btn.config(
            bg=COLORS["accent_dim"]))

        # ── Feedback bar ──
        self.feedback_label = tk.Label(main, text="",
                                        font=(FONT_FAMILY, 9),
                                        fg=COLORS["text_dim"],
                                        bg=COLORS["bg_dark"])
        self.feedback_label.pack(pady=(2, 0))

    # ── Helpers ────────────────────────────────────────────────────────────

    def _lighten(self, hex_color, factor=0.15):
        """Lighten a hex color by a factor."""
        hex_color = hex_color.lstrip("#")
        r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:], 16)
        r = min(255, int(r + (255 - r) * factor))
        g = min(255, int(g + (255 - g) * factor))
        b = min(255, int(b + (255 - b) * factor))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw_status_dot(self, connected):
        self.status_dot.delete("all")
        color = COLORS["success"] if connected else COLORS["danger"]
        self.status_dot.create_oval(2, 2, 10, 10, fill=color, outline="")

    def _set_feedback(self, text, color=None):
        if color is None:
            color = COLORS["text_dim"]
        self.feedback_label.config(text=text, fg=color)

    # ── Connection ────────────────────────────────────────────────────────

    def _connect(self):
        try:
            self.driver = WinRing0()
            self.driver.open()
            self.ec = ECMailbox(self.driver)
            self.connected = True
            self._draw_status_dot(True)
            self.status_label.config(text="Connected", fg=COLORS["success"])

            # Safety: always restore auto mode on connect (protects against
            # persisted fan speeds across reboots or crashes)
            try:
                self.ec.restore_auto()
            except Exception:
                pass

            self._set_feedback("Ready. Use sliders or presets to control fan speed.")
        except Exception as e:
            self.connected = False
            self._draw_status_dot(False)
            self.status_label.config(text="Not connected", fg=COLORS["danger"])
            self._set_feedback(f"Error: {e}", COLORS["danger"])

    # ── Monitor Thread ────────────────────────────────────────────────────

    def _start_monitor(self):
        self._monitor_thread = threading.Thread(target=self._monitor_loop,
                                                 daemon=True)
        self._monitor_thread.start()

    def _monitor_loop(self):
        while self.running:
            if self.connected:
                try:
                    f1 = self.ec.read_fan1()
                    f2 = self.ec.read_fan2()
                    self.root.after(0, self.gauge1.set_value, f1)
                    self.root.after(0, self.gauge2.set_value, f2)
                except Exception:
                    pass

                # CPU temperature
                if HAS_PSUTIL:
                    try:
                        temps = psutil.sensors_temperatures()
                        if temps:
                            for name, entries in temps.items():
                                if entries:
                                    temp = entries[0].current
                                    self.root.after(0, self.temp_label.config,
                                                    {"text": f"CPU: {temp:.0f}°C"})
                                    break
                    except Exception:
                        pass

            time.sleep(1.5)

    # ── Slider Callbacks ──────────────────────────────────────────────────

    def _on_slider1_change(self, value):
        if self.linked_var.get():
            self.slider2.set_value(value, trigger=False)
        self._apply_fan_speeds()

    def _on_slider2_change(self, value):
        if self.linked_var.get():
            self.slider1.set_value(value, trigger=False)
        self._apply_fan_speeds()

    def _apply_fan_speeds(self):
        global _above_safe_confirmed
        if not self.connected:
            self._set_feedback("Not connected to EC", COLORS["danger"])
            return

        f1 = self.slider1.get_value()
        f2 = self.slider2.get_value()

        # Safety gate: confirm before exceeding EC's normal maximum
        if (f1 > SAFE_MAX or f2 > SAFE_MAX) and not _above_safe_confirmed:
            result = messagebox.askokcancel(
                "⚠️ Above Normal Range",
                f"You are setting fans above {SAFE_MAX}%, which exceeds\n"
                f"the EC's normal maximum operating range.\n\n"
                f"Fan 1: {f1}%  |  Fan 2: {f2}%\n\n"
                f"This will make the fans significantly louder.\n"
                f"Continue?",
                icon="warning"
            )
            if not result:
                # Reset sliders to safe max
                self.slider1.set_value(min(f1, SAFE_MAX))
                self.slider2.set_value(min(f2, SAFE_MAX))
                self._set_feedback(f"Capped at {SAFE_MAX}% (safe range)")
                return
            _above_safe_confirmed = True

        self.auto_mode = False

        def _send():
            try:
                ok1 = self.ec.set_fan1(f1)
                ok2 = self.ec.set_fan2(f2)
                if ok1 and ok2:
                    label = f"Set Fan 1: {f1}%, Fan 2: {f2}%"
                    color = COLORS["accent"]
                    if f1 > SAFE_MAX or f2 > SAFE_MAX:
                        label += "  ⚠️"
                        color = COLORS["warning"]
                    self.root.after(0, self._set_feedback, label, color)
                else:
                    self.root.after(0, self._set_feedback,
                                    "Warning: EC did not confirm", COLORS["warning"])
            except Exception as e:
                self.root.after(0, self._set_feedback, f"Error: {e}", COLORS["danger"])

        threading.Thread(target=_send, daemon=True).start()

    # ── Presets ───────────────────────────────────────────────────────────

    def _rebuild_presets(self):
        """Rebuild the preset buttons from built-in + custom presets."""
        # Clear existing
        for w in self.preset_frame.winfo_children():
            w.destroy()

        # Row 1: built-in presets
        row1 = tk.Frame(self.preset_frame, bg=COLORS["bg_panel"])
        row1.pack(pady=(0, 4))

        for name, speed, color in BUILTIN_PRESETS:
            btn = tk.Button(row1, text=f"{name}\n{speed}%",
                            font=(FONT_FAMILY, 9),
                            fg=COLORS["text"], bg=color,
                            activebackground=color,
                            activeforeground=COLORS["text_bright"],
                            relief="flat", padx=8, pady=5, width=8,
                            cursor="hand2",
                            command=lambda s=speed: self._apply_preset(s, s))
            btn.pack(side="left", padx=2)
            btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(
                bg=self._lighten(c, 0.15)))
            btn.bind("<Leave>", lambda e, b=btn, c=color: b.config(bg=c))

        # Row 2: custom presets + add button
        if self.custom_presets or True:  # always show row for '+' button
            row2 = tk.Frame(self.preset_frame, bg=COLORS["bg_panel"])
            row2.pack(pady=(2, 0))

            for i, preset in enumerate(self.custom_presets):
                pname = preset["name"]
                pspeed = preset["speed"]
                color = COLORS["accent_dim"] if pspeed <= SAFE_MAX else COLORS["preset_full"]
                btn = tk.Button(row2, text=f"{pname}\n{pspeed}%",
                                font=(FONT_FAMILY, 9),
                                fg=COLORS["text"], bg=color,
                                activebackground=color,
                                activeforeground=COLORS["text_bright"],
                                relief="flat", padx=8, pady=5, width=8,
                                cursor="hand2",
                                command=lambda s=pspeed: self._apply_preset(s, s))
                btn.pack(side="left", padx=2)
                btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(
                    bg=self._lighten(c, 0.15)))
                btn.bind("<Leave>", lambda e, b=btn, c=color: b.config(bg=c))
                # Right-click to delete
                btn.bind("<Button-3>", lambda e, idx=i: self._delete_preset(idx))

            # '+' button to add custom preset
            add_btn = tk.Button(row2, text="＋",
                                font=(FONT_FAMILY, 12, "bold"),
                                fg=COLORS["accent"], bg=COLORS["bg_card"],
                                activebackground=COLORS["bg_panel"],
                                activeforeground=COLORS["accent_glow"],
                                relief="flat", padx=8, pady=3, width=3,
                                cursor="hand2",
                                command=self._add_preset_dialog)
            add_btn.pack(side="left", padx=4)
            add_btn.bind("<Enter>", lambda e: add_btn.config(bg=COLORS["bg_panel"]))
            add_btn.bind("<Leave>", lambda e: add_btn.config(bg=COLORS["bg_card"]))

    def _add_preset_dialog(self):
        """Open a dialog to create a custom preset."""
        dlg = tk.Toplevel(self.root)
        dlg.title("New Custom Preset")
        dlg.configure(bg=COLORS["bg_panel"])
        dlg.resizable(False, False)
        dlg.grab_set()  # modal

        # Center on parent
        dlg.geometry(f"+{self.root.winfo_x() + 60}+{self.root.winfo_y() + 200}")

        # Name
        tk.Label(dlg, text="Preset Name:", font=(FONT_FAMILY, 10),
                 fg=COLORS["text"], bg=COLORS["bg_panel"]).pack(padx=15, pady=(12, 2), anchor="w")
        name_entry = tk.Entry(dlg, font=(FONT_FAMILY, 10),
                              bg=COLORS["bg_input"], fg=COLORS["text"],
                              insertbackground=COLORS["text"], width=20)
        name_entry.pack(padx=15, pady=(0, 8))

        # Speed
        tk.Label(dlg, text="Fan Speed (0 or 18-100%):", font=(FONT_FAMILY, 10),
                 fg=COLORS["text"], bg=COLORS["bg_panel"]).pack(padx=15, pady=(0, 2), anchor="w")
        speed_entry = tk.Entry(dlg, font=(FONT_FAMILY, 10),
                               bg=COLORS["bg_input"], fg=COLORS["text"],
                               insertbackground=COLORS["text"], width=20)
        speed_entry.pack(padx=15, pady=(0, 8))

        # Warning label
        warn_label = tk.Label(dlg, text="", font=(FONT_FAMILY, 8),
                              fg=COLORS["warning"], bg=COLORS["bg_panel"])
        warn_label.pack(padx=15)

        def _save():
            name = name_entry.get().strip()
            speed_str = speed_entry.get().strip()

            if not name:
                warn_label.config(text="Please enter a name.", fg=COLORS["danger"])
                return

            try:
                speed = int(speed_str)
            except ValueError:
                warn_label.config(text="Speed must be a number.", fg=COLORS["danger"])
                return

            if 1 <= speed < MIN_FAN_SPEED:
                warn_label.config(
                    text=f"Values 1-{MIN_FAN_SPEED - 1}% cause fan pulsing.\nUse 0 (off) or {MIN_FAN_SPEED}-100.",
                    fg=COLORS["danger"])
                return

            if speed < 0 or speed > 100:
                warn_label.config(text="Speed must be 0-100%.", fg=COLORS["danger"])
                return

            if speed > SAFE_MAX:
                if not messagebox.askokcancel(
                    "⚠️ Above Normal Range",
                    f"{speed}% exceeds the EC's normal max ({SAFE_MAX}%).\n"
                    f"Create this preset anyway?",
                    parent=dlg, icon="warning"
                ):
                    return

            self.custom_presets.append({"name": name, "speed": speed})
            self._rebuild_presets()
            self._save_config()
            dlg.destroy()

        # Buttons
        btn_frame = tk.Frame(dlg, bg=COLORS["bg_panel"])
        btn_frame.pack(pady=(8, 12))

        tk.Button(btn_frame, text="Save", font=(FONT_FAMILY, 10, "bold"),
                  fg=COLORS["text_bright"], bg=COLORS["accent_dim"],
                  activebackground=COLORS["accent"],
                  relief="flat", padx=15, pady=3, cursor="hand2",
                  command=_save).pack(side="left", padx=5)

        tk.Button(btn_frame, text="Cancel", font=(FONT_FAMILY, 10),
                  fg=COLORS["text"], bg=COLORS["bg_card"],
                  activebackground=COLORS["bg_panel"],
                  relief="flat", padx=15, pady=3, cursor="hand2",
                  command=dlg.destroy).pack(side="left", padx=5)

    def _delete_preset(self, index):
        """Delete a custom preset (right-click)."""
        preset = self.custom_presets[index]
        if messagebox.askyesno(
            "Delete Preset",
            f"Delete custom preset '{preset['name']}'?"
        ):
            self.custom_presets.pop(index)
            self._rebuild_presets()
            self._save_config()

    def _apply_preset(self, f1, f2):
        self.slider1.set_value(clamp_fan_speed(f1))
        self.slider2.set_value(clamp_fan_speed(f2))
        self._apply_fan_speeds()

    # ── Hold Mode ─────────────────────────────────────────────────────────

    def _toggle_hold(self):
        if self.hold_var.get():
            self.hold_active = True
            self.auto_mode = False
            self._set_feedback("Hold mode ON — re-sending every 3s", COLORS["warning"])
            self._hold_thread = threading.Thread(target=self._hold_loop, daemon=True)
            self._hold_thread.start()
        else:
            self.hold_active = False
            self._set_feedback("Hold mode OFF")

    def _hold_loop(self):
        while self.hold_active and self.running:
            if self.connected:
                f1 = self.slider1.get_value()
                f2 = self.slider2.get_value()
                try:
                    self.ec.set_fan1(f1)
                    self.ec.set_fan2(f2)
                except Exception:
                    pass
            time.sleep(3)

    # ── Auto Mode ─────────────────────────────────────────────────────────

    def _restore_auto(self):
        global _above_safe_confirmed
        if not self.connected:
            self._set_feedback("Not connected to EC", COLORS["danger"])
            return

        self.hold_active = False
        self.hold_var.set(False)
        self.auto_mode = True
        _above_safe_confirmed = False  # Reset confirmation for next time

        def _send():
            try:
                ok = self.ec.restore_auto()
                if ok:
                    self.root.after(0, self._set_feedback,
                                    "Automatic fan control restored", COLORS["success"])
                else:
                    self.root.after(0, self._set_feedback,
                                    "Warning: auto restore may not have confirmed",
                                    COLORS["warning"])
            except Exception as e:
                self.root.after(0, self._set_feedback, f"Error: {e}", COLORS["danger"])

        threading.Thread(target=_send, daemon=True).start()

    # ── Config Persistence ────────────────────────────────────────────────

    def _load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f:
                    cfg = json.load(f)
                self.slider1.set_value(cfg.get("fan1", 30))
                self.slider2.set_value(cfg.get("fan2", 30))
                self.linked_var.set(cfg.get("linked", True))
                self.custom_presets = cfg.get("custom_presets", [])
                self._rebuild_presets()
        except Exception:
            pass

    def _save_config(self):
        try:
            cfg = {
                "fan1": self.slider1.get_value(),
                "fan2": self.slider2.get_value(),
                "linked": self.linked_var.get(),
                "custom_presets": self.custom_presets,
            }
            with open(CONFIG_FILE, "w") as f:
                json.dump(cfg, f)
        except Exception:
            pass

    # ── Cleanup ───────────────────────────────────────────────────────────

    def _safety_restore(self):
        """Restore auto mode on exit for safety (atexit handler)."""
        # Always restore — don't check auto_mode flag, this is a safety net
        if self.connected:
            try:
                self.ec.restore_auto()
            except Exception:
                pass

    def _on_close(self):
        """Full shutdown — called from tray Quit."""
        self.running = False
        self.hold_active = False
        self._save_config()

        # Stop tray icon
        if self._tray_icon:
            try:
                self._tray_icon.stop()
            except Exception:
                pass

        # Restore auto if we changed anything
        if self.connected and not self.auto_mode:
            try:
                self.ec.restore_auto()
            except Exception:
                pass

        if self.driver:
            try:
                self.driver.close()
            except Exception:
                pass
            try:
                self.driver.stop_driver()
            except Exception:
                pass

        self.root.destroy()

    # ── Run ───────────────────────────────────────────────────────────────

    def run(self):
        # Center window on screen
        self.root.update_idletasks()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"+{x}+{y}")

        self.root.mainloop()


# ==============================================================================
# Entry Point
# ==============================================================================

def _run_startup_safety():
    """Headless mode: restore EC auto fan control and exit silently."""
    try:
        driver = WinRing0()
        driver.open()
        ec = ECMailbox(driver)
        ec.restore_auto()
        driver.close()
        driver.stop_driver()
    except Exception:
        pass
    sys.exit(0)


def main():
    # Check for startup safety mode (runs silently at boot)
    if '--startup-safety' in sys.argv:
        _run_startup_safety()
        return

    if not is_admin():
        messagebox.showerror(
            "Administrator Required",
            "This tool must be run as Administrator.\n\n"
            "Right-click your terminal or this script and select\n"
            "'Run as administrator'."
        )
        sys.exit(1)

    app = FanControlApp()
    app.run()


if __name__ == "__main__":
    main()
