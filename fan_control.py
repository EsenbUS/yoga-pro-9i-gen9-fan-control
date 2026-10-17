#!/usr/bin/env python3
"""
Fan Control Tool - Lenovo Yoga Pro 9i Gen 9 (2024)

Controls fan speed through the firmware's own WMI interface
(root\\WMI : LENOVO_OTHER_METHOD, ACPI method \\_SB.GZFD.WMAE).

The firmware method talks to the EC mailbox at 0x5C0/0x5C4 itself, holding
the EC0.LFCM mutex and the ACPI Global Lock while it does so. That means:
  - No kernel driver is needed (WinRing0 is gone).
  - Our transactions can never interleave with the firmware's or Lenovo
    services' own EC mailbox traffic.

Feature IDs (from the DSDT):
  0x04030001  Fan 1   Get: percent * 100   Set: 0 = auto, 1 = off, N = N/100 %
  0x04030002  Fan 2   (same encoding)

MUST BE RUN AS ADMINISTRATOR (the Lenovo WMI classes deny non-admin access).

Use GUI. CLI is not updated for now.
"""

import ctypes
import sys
import os
import time
import signal
import argparse
import queue
import shutil
import subprocess
import threading
from concurrent.futures import Future

import pythoncom
import win32com.client


# ==============================================================================
# Lenovo WMI Interface
# ==============================================================================

WMI_NAMESPACE = r"winmgmts:{impersonationLevel=impersonate}!\\.\root\WMI"
WMI_CLASS     = "LENOVO_OTHER_METHOD"

FEATURE_FAN1 = 0x04030001
FEATURE_FAN2 = 0x04030002

# SetFeatureValue values for the fan features (see WMAE in the DSDT)
FAN_VALUE_AUTO = 0
FAN_VALUE_OFF  = 1

MIN_FAN_SPEED = 18
SAFE_MAX      = 48

CALL_TIMEOUT = 10.0  # seconds to wait for a single WMI call


def clamp_fan_speed(val):
    """Clamp to valid range: 0 or 18-100."""
    val = max(0, min(100, int(val)))
    if 1 <= val < MIN_FAN_SPEED:
        return MIN_FAN_SPEED
    return val


def _fan_set_value(percent):
    """Encode a percentage the way WMAE expects it."""
    percent = clamp_fan_speed(percent)
    if percent == 0:
        return FAN_VALUE_OFF
    return percent * 100


class FanControlPaused(RuntimeError):
    """Raised when EC access is requested while the controller is paused."""


class LenovoWMI:
    """Thin wrapper around LENOVO_OTHER_METHOD. Must be used from one COM thread."""

    def __init__(self):
        self._svc = win32com.client.GetObject(WMI_NAMESPACE)
        instances = list(self._svc.InstancesOf(WMI_CLASS))
        if not instances:
            raise RuntimeError(f"WMI class {WMI_CLASS} has no instances.")
        self._path = instances[0].Path_.Path
        self._cls = self._svc.Get(WMI_CLASS)

    def _call(self, method, **args):
        params = self._cls.Methods_(method).InParameters.SpawnInstance_()
        for name, value in args.items():
            params.Properties_.Item(name).Value = value
        return self._svc.ExecMethod(self._path, method, params)

    def get_feature(self, feature_id):
        out = self._call("GetFeatureValue", IDs=feature_id)
        return int(out.Properties_.Item("value").Value)

    def set_feature(self, feature_id, value):
        self._call("SetFeatureValue", IDs=feature_id, value=value)
        return True


# ==============================================================================
# High-Level FanController
# ==============================================================================

class FanController:
    """
    High-level fan controller.

    All WMI calls run on one dedicated worker thread (COM objects are bound
    to the thread that created them), which also serializes them.

    pause() blocks all further EC access until unpause(). The GUI pauses
    whenever the display is off or the system is suspending, so nothing
    touches the EC around sleep / hibernate transitions.
    """

    def __init__(self):
        self._jobs = queue.Queue()
        self._paused = threading.Event()
        self._ready = threading.Event()
        self._init_error = None
        self._thread = threading.Thread(target=self._worker, name="FanWMI", daemon=True)
        self._thread.start()
        if not self._ready.wait(CALL_TIMEOUT):
            raise RuntimeError("Timed out connecting to Lenovo WMI.")
        if self._init_error:
            raise RuntimeError(f"Cannot open Lenovo WMI ({self._init_error}).")

    # ── Worker ───────────────────────────────────────────────────────────

    def _worker(self):
        pythoncom.CoInitialize()
        try:
            try:
                wmi = LenovoWMI()
            except Exception as e:
                self._init_error = e
                return
            finally:
                self._ready.set()

            while True:
                job = self._jobs.get()
                if job is None:
                    break
                fn, fut = job
                if not fut.set_running_or_notify_cancel():
                    continue
                # Checked at execution time, so jobs queued before pause() never run after it
                if self._paused.is_set():
                    fut.set_exception(FanControlPaused("Fan control is paused"))
                    continue
                try:
                    fut.set_result(fn(wmi))
                except Exception as e:
                    fut.set_exception(e)
        finally:
            pythoncom.CoUninitialize()

    def _run(self, fn):
        if not self._thread.is_alive():
            raise RuntimeError("Fan controller is closed.")
        fut = Future()
        self._jobs.put((fn, fut))
        return fut.result(timeout=CALL_TIMEOUT)

    # ── Fan Access ───────────────────────────────────────────────────────

    def read_fan1(self):
        return self._run(lambda w: w.get_feature(FEATURE_FAN1) // 100)

    def read_fan2(self):
        return self._run(lambda w: w.get_feature(FEATURE_FAN2) // 100)

    def read_fans(self):
        """Read both fans. Returns (f1, f2)."""
        return self._run(lambda w: (w.get_feature(FEATURE_FAN1) // 100,
                                    w.get_feature(FEATURE_FAN2) // 100))

    def set_fan1(self, percent):
        return self._run(lambda w: w.set_feature(FEATURE_FAN1, _fan_set_value(percent)))

    def set_fan2(self, percent):
        return self._run(lambda w: w.set_feature(FEATURE_FAN2, _fan_set_value(percent)))

    def set_fans(self, f1, f2):
        def _both(w):
            r1 = w.set_feature(FEATURE_FAN1, _fan_set_value(f1))
            r2 = w.set_feature(FEATURE_FAN2, _fan_set_value(f2))
            return r1, r2
        return self._run(_both)

    def restore_auto(self):
        def _auto(w):
            r1 = w.set_feature(FEATURE_FAN1, FAN_VALUE_AUTO)
            r2 = w.set_feature(FEATURE_FAN2, FAN_VALUE_AUTO)
            return r1 and r2
        return self._run(_auto)

    # ── Pause / Lifecycle ────────────────────────────────────────────────

    @property
    def paused(self):
        return self._paused.is_set()

    def pause(self):
        """Block all EC access. Does not touch the EC itself."""
        self._paused.set()

    def unpause(self):
        self._paused.clear()

    def close(self):
        """Stop the worker thread."""
        if self._thread.is_alive():
            self._jobs.put(None)
            self._thread.join(timeout=CALL_TIMEOUT)


# ==============================================================================
# Legacy WinRing0 Cleanup
# ==============================================================================

LEGACY_SERVICE_NAME = "WinRing0_Transient"


def cleanup_legacy_driver():
    """Remove the WinRing0 service and driver copy left by older versions."""
    try:
        subprocess.run(['sc', 'stop', LEGACY_SERVICE_NAME], capture_output=True,
                       creationflags=subprocess.CREATE_NO_WINDOW)
        subprocess.run(['sc', 'delete', LEGACY_SERVICE_NAME], capture_output=True,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass

    appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    driver_dir = os.path.join(appdata, 'Yoga Fan Control')
    if os.path.exists(driver_dir):
        try:
            shutil.rmtree(driver_dir)
        except Exception:
            pass


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ==============================================================================
# CLI Interface
# ==============================================================================

_fc = None
_auto_restore_on_exit = False


def restore_auto_on_exit():
    global _fc, _auto_restore_on_exit
    if _fc and _auto_restore_on_exit:
        try:
            _fc.restore_auto()
            print("\n[Safety] Restored automatic fan control.")
        except Exception:
            pass


def signal_handler(sig, frame):
    restore_auto_on_exit()
    sys.exit(0)


def cmd_read(fc):
    f1, f2 = fc.read_fans()
    print(f"Fan 1: {f1}%")
    print(f"Fan 2: {f2}%")


def cmd_set(fc, fan1_pct, fan2_pct=None):
    if fan2_pct is None:
        fan2_pct = fan1_pct
    fan1_pct = clamp_fan_speed(fan1_pct)
    fan2_pct = clamp_fan_speed(fan2_pct)

    if fan1_pct > SAFE_MAX or fan2_pct > SAFE_MAX:
        print(f"WARNING: Setting fans above {SAFE_MAX}% exceeds the EC's normal range.")
        confirm = input("Continue? (y/N): ").strip().lower()
        if confirm != 'y':
            print("Aborted.")
            return

    print(f"Setting Fan 1 to {fan1_pct}%...", end=" ")
    ok1 = fc.set_fan1(fan1_pct)
    print("OK" if ok1 else "FAILED")

    print(f"Setting Fan 2 to {fan2_pct}%...", end=" ")
    ok2 = fc.set_fan2(fan2_pct)
    print("OK" if ok2 else "FAILED")


def cmd_auto(fc):
    print("Restoring automatic fan control...", end=" ")
    ok = fc.restore_auto()
    print("OK" if ok else "FAILED")


def cmd_monitor(fc):
    print("Monitoring fan speeds (Ctrl+C to stop)...\n")
    try:
        while True:
            f1, f2 = fc.read_fans()
            print(f"\rFan 1: {f1:3d}%  |  Fan 2: {f2:3d}%", end="", flush=True)
            time.sleep(2)
    except KeyboardInterrupt:
        print("\n\nMonitoring stopped.")


def cmd_hold(fc, fan1_pct, fan2_pct=None):
    global _auto_restore_on_exit
    _auto_restore_on_exit = True

    if fan2_pct is None:
        fan2_pct = fan1_pct
    fan1_pct = clamp_fan_speed(fan1_pct)
    fan2_pct = clamp_fan_speed(fan2_pct)

    if fan1_pct > SAFE_MAX or fan2_pct > SAFE_MAX:
        print(f"WARNING: Holding fans above {SAFE_MAX}% exceeds the EC's normal range.")
        confirm = input("Continue? (y/N): ").strip().lower()
        if confirm != 'y':
            print("Aborted.")
            return

    print(f"Holding fans at Fan 1: {fan1_pct}%, Fan 2: {fan2_pct}%")
    print("Press Ctrl+C to stop and restore automatic control.\n")

    try:
        while True:
            fc.set_fans(fan1_pct, fan2_pct)
            actual1, actual2 = fc.read_fans()
            print(f"\rTarget: {fan1_pct}%/{fan2_pct}%  |  Actual: {actual1}%/{actual2}%",
                  end="", flush=True)
            time.sleep(3)
    except KeyboardInterrupt:
        print()


def main():
    parser = argparse.ArgumentParser(
        description="Fan Control (Lenovo WMI) — Lenovo Yoga Pro 9i Gen 9",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  read              Read current fan speeds
  set <f1> [f2]     Set fan speed(s) in percent (0-100)
  auto              Restore automatic EC fan control
  monitor           Continuously display fan speeds
  hold <f1> [f2]    Set and maintain fan speeds (re-sends every 3s)
        """
    )
    parser.add_argument("command", choices=["read", "set", "auto", "monitor", "hold"])
    parser.add_argument("fan1", nargs="?", type=int, default=None)
    parser.add_argument("fan2", nargs="?", type=int, default=None)
    args = parser.parse_args()

    if args.command in ("set", "hold") and args.fan1 is None:
        parser.error(f"'{args.command}' requires at least one fan speed")

    if not is_admin():
        print("ERROR: Must be run as Administrator.")
        sys.exit(1)

    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGBREAK, signal_handler)

    global _fc
    try:
        _fc = FanController()

        if args.command == "read":
            cmd_read(_fc)
        elif args.command == "set":
            cmd_set(_fc, args.fan1, args.fan2)
        elif args.command == "auto":
            cmd_auto(_fc)
        elif args.command == "monitor":
            cmd_monitor(_fc)
        elif args.command == "hold":
            cmd_hold(_fc, args.fan1, args.fan2)
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    finally:
        restore_auto_on_exit()


if __name__ == "__main__":
    main()
