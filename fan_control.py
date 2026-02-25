#!/usr/bin/env python3
"""
EC Mailbox Fan Control Tool - Lenovo Yoga Pro 9i Gen 9 (2024)

Controls fan speed via the EC mailbox interface at I/O ports 0x5C0/0x5C4.

Uses a TRANSIENT WinRing0 driver approach:
  - The driver service is created, started, used, stopped, and DELETED
    in a single atomic operation (milliseconds).
  - The service never persists between fan writes.
  - This prevents the HAL_INITIALIZATION_FAILED (0x5C) crash during
    Modern Standby: Windows cannot send a power IRP to a service that
    doesn't exist.

MUST BE RUN AS ADMINISTRATOR.

Usage:
    python fan_control.py read              Read current fan speeds
    python fan_control.py set <f1%> [f2%]   Set fan speed(s), 0-100
    python fan_control.py auto              Restore automatic EC control
    python fan_control.py monitor           Continuously display fan speeds
    python fan_control.py hold <f1%> [f2%]  Set and maintain fan speeds
"""

import ctypes
import ctypes.wintypes as wintypes
import sys
import os
import time
import signal
import argparse
import struct
import shutil
import threading

# ==============================================================================
# WinRing0 Driver Interface — Transient Service Model
# ==============================================================================

# IOCTLs for WinRing0 driver
OLS_TYPE = 40000  # 0x9C40

def CTL_CODE(DeviceType, Function, Method, Access):
    return (DeviceType << 16) | (Access << 14) | (Function << 2) | Method

METHOD_BUFFERED = 0
FILE_READ_ACCESS = 1
FILE_WRITE_ACCESS = 2

IOCTL_OLS_READ_IO_PORT_BYTE  = CTL_CODE(OLS_TYPE, 0x833, METHOD_BUFFERED, FILE_READ_ACCESS)
IOCTL_OLS_WRITE_IO_PORT_BYTE = CTL_CODE(OLS_TYPE, 0x836, METHOD_BUFFERED, FILE_WRITE_ACCESS)

# Windows API constants
GENERIC_READ  = 0x80000000
GENERIC_WRITE = 0x40000000
FILE_SHARE_READ  = 0x01
FILE_SHARE_WRITE = 0x02
OPEN_EXISTING = 3
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

kernel32  = ctypes.windll.kernel32
advapi32  = ctypes.windll.advapi32

# CreateFile
CreateFileW = kernel32.CreateFileW
CreateFileW.restype = wintypes.HANDLE
CreateFileW.argtypes = [
    wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
    wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE
]

# DeviceIoControl
DeviceIoControl = kernel32.DeviceIoControl
DeviceIoControl.restype = wintypes.BOOL
DeviceIoControl.argtypes = [
    wintypes.HANDLE, wintypes.DWORD, wintypes.LPVOID, wintypes.DWORD,
    wintypes.LPVOID, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID
]

CloseHandle = kernel32.CloseHandle
CloseHandle.restype = wintypes.BOOL
CloseHandle.argtypes = [wintypes.HANDLE]

# SCM API
SC_HANDLE = wintypes.HANDLE
SC_MANAGER_ALL_ACCESS = 0xF003F
SERVICE_ALL_ACCESS    = 0xF01FF
SERVICE_KERNEL_DRIVER = 0x01
SERVICE_DEMAND_START  = 0x03
SERVICE_ERROR_NORMAL  = 0x01
SERVICE_CONTROL_STOP  = 0x01

advapi32.OpenSCManagerW.restype = SC_HANDLE
advapi32.OpenSCManagerW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD]

advapi32.OpenServiceW.restype = SC_HANDLE
advapi32.OpenServiceW.argtypes = [SC_HANDLE, wintypes.LPCWSTR, wintypes.DWORD]

advapi32.CreateServiceW.restype = SC_HANDLE
advapi32.CreateServiceW.argtypes = [
    SC_HANDLE, wintypes.LPCWSTR, wintypes.LPCWSTR,
    wintypes.DWORD, wintypes.DWORD, wintypes.DWORD, wintypes.DWORD,
    wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPVOID,
    wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR,
]

advapi32.StartServiceW.restype = wintypes.BOOL
advapi32.StartServiceW.argtypes = [SC_HANDLE, wintypes.DWORD, wintypes.LPVOID]

advapi32.ControlService.restype = wintypes.BOOL
advapi32.ControlService.argtypes = [SC_HANDLE, wintypes.DWORD, ctypes.c_void_p]

advapi32.DeleteService.restype = wintypes.BOOL
advapi32.DeleteService.argtypes = [SC_HANDLE]

advapi32.CloseServiceHandle.restype = wintypes.BOOL
advapi32.CloseServiceHandle.argtypes = [SC_HANDLE]


def _get_driver_path():
    """Find WinRing0x64.sys — copy to a stable (non-temp) directory."""
    appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    driver_dir = os.path.join(appdata, 'Yoga Fan Control')
    os.makedirs(driver_dir, exist_ok=True)
    dest = os.path.join(driver_dir, 'WinRing0x64.sys')

    if os.path.exists(dest):
        return dest

    # Locate source
    if getattr(sys, '_MEIPASS', None):
        src = os.path.join(sys._MEIPASS, 'WinRing0x64.sys')
    elif getattr(sys, 'frozen', False):
        src = os.path.join(os.path.dirname(sys.executable), 'WinRing0x64.sys')
    else:
        src = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WinRing0x64.sys')

    if not os.path.exists(src):
        raise FileNotFoundError(
            "Cannot find WinRing0x64.sys. "
            "Place it in the same folder as this program."
        )
    shutil.copy2(src, dest)
    return dest


class TransientWinRing0:
    """
    Transient WinRing0 driver session.

    Creates the service, opens a handle, exposes I/O port read/write,
    closes the handle, stops and DELETES the service — all within a
    single context manager invocation.

    Because the service is deleted before Python yields control back,
    it will never be alive during a sleep transition.
    """

    SERVICE_NAME   = "WinRing0_Transient"
    DEVICE_NAME    = r"\\.\WinRing0_1_2_0"

    def __init__(self):
        self._sc_manager = None
        self._service    = None
        self.handle      = None
        self._lock       = threading.Lock()

    # ── Lifecycle ────────────────────────────────────────────────────────

    def open(self):
        """Install, start, and open the driver. Idempotent if already open."""
        if self.handle and self.handle != INVALID_HANDLE_VALUE:
            return

        driver_path = _get_driver_path()

        self._sc_manager = advapi32.OpenSCManagerW(
            None, None, SC_MANAGER_ALL_ACCESS
        )
        if not self._sc_manager:
            raise RuntimeError(
                f"Cannot open SCM (error {kernel32.GetLastError()}). "
                "Run as Administrator."
            )

        # Clean up any leftover service from a previous crash
        self._cleanup_existing_service()

        # Create fresh service
        self._service = advapi32.CreateServiceW(
            self._sc_manager,
            self.SERVICE_NAME,            # service name
            "WinRing0 (Transient)",       # display name
            SERVICE_ALL_ACCESS,
            SERVICE_KERNEL_DRIVER,
            SERVICE_DEMAND_START,
            SERVICE_ERROR_NORMAL,
            driver_path,
            None, None, None, None, None,
        )
        if not self._service:
            err = kernel32.GetLastError()
            raise RuntimeError(
                f"Cannot create WinRing0 service (error {err})."
            )

        # Start service
        if not advapi32.StartServiceW(self._service, 0, None):
            err = kernel32.GetLastError()
            ERROR_SERVICE_ALREADY_RUNNING = 1056
            if err != ERROR_SERVICE_ALREADY_RUNNING:
                advapi32.DeleteService(self._service)
                raise RuntimeError(
                    f"Cannot start WinRing0 service (error {err})."
                )

        # Open device handle
        self.handle = CreateFileW(
            self.DEVICE_NAME,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None, OPEN_EXISTING, 0, None,
        )
        if self.handle == INVALID_HANDLE_VALUE:
            err = kernel32.GetLastError()
            self._stop_and_delete()
            raise RuntimeError(
                f"Cannot open WinRing0 device handle (error {err})."
            )

    def close(self):
        """Close handle, stop service, DELETE service entry. Sleep-safe."""
        if self.handle and self.handle != INVALID_HANDLE_VALUE:
            CloseHandle(self.handle)
            self.handle = None
        self._stop_and_delete()

    def _cleanup_existing_service(self):
        """Remove any leftover service with our name."""
        svc = advapi32.OpenServiceW(
            self._sc_manager, self.SERVICE_NAME, SERVICE_ALL_ACCESS
        )
        if svc:
            status_buf = ctypes.create_string_buffer(28)
            advapi32.ControlService(svc, SERVICE_CONTROL_STOP, status_buf)
            time.sleep(0.1)  # brief wait for stop
            advapi32.DeleteService(svc)
            advapi32.CloseServiceHandle(svc)

    def _stop_and_delete(self):
        """Stop and delete the service so it never sees a sleep IRP."""
        if self._service:
            status_buf = ctypes.create_string_buffer(28)
            advapi32.ControlService(self._service, SERVICE_CONTROL_STOP, status_buf)
            # Give the kernel a moment to process the stop
            time.sleep(0.05)
            advapi32.DeleteService(self._service)
            advapi32.CloseServiceHandle(self._service)
            self._service = None
        if self._sc_manager:
            advapi32.CloseServiceHandle(self._sc_manager)
            self._sc_manager = None

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, *args):
        self.close()

    # ── I/O Port Access ──────────────────────────────────────────────────

    def read_io_port_byte(self, port):
        """Read a single byte from an I/O port via WinRing0 IOCTL."""
        in_buf = struct.pack('<I', port)
        out_buf = ctypes.create_string_buffer(4)
        returned = wintypes.DWORD(0)
        ok = DeviceIoControl(
            self.handle, IOCTL_OLS_READ_IO_PORT_BYTE,
            in_buf, len(in_buf),
            out_buf, len(out_buf),
            ctypes.byref(returned), None,
        )
        if not ok:
            raise RuntimeError(f"I/O read failed on port 0x{port:X}")
        return struct.unpack('<I', out_buf.raw)[0] & 0xFF

    def write_io_port_byte(self, port, value):
        """Write a single byte to an I/O port via WinRing0 IOCTL."""
        in_buf = struct.pack('<II', port, value & 0xFF)
        returned = wintypes.DWORD(0)
        ok = DeviceIoControl(
            self.handle, IOCTL_OLS_WRITE_IO_PORT_BYTE,
            in_buf, len(in_buf),
            None, 0,
            ctypes.byref(returned), None,
        )
        if not ok:
            raise RuntimeError(f"I/O write 0x{value:02X} failed on port 0x{port:X}")


# ==============================================================================
# EC Mailbox Protocol
# ==============================================================================

PORT_DATA = 0x5C0
PORT_CMD  = 0x5C4

TIMEOUT_ITERATIONS = 0x10000
POLL_DELAY = 0.00001  # 10 µs

CMD_FAN       = 0xEF
SUBCMD_SET_FAN1 = 0x61
SUBCMD_SET_FAN2 = 0x62
SUBCMD_QUERY    = 0x63

QUERY_READ_FAN1 = 0x01
QUERY_READ_FAN2 = 0x02
QUERY_AUTO_MODE = 0x03

SUCCESS_CODE = 0xAC

MIN_FAN_SPEED = 18
SAFE_MAX      = 48


def clamp_fan_speed(val):
    """Clamp to valid range: 0 or 18-100."""
    val = max(0, min(100, int(val)))
    if 1 <= val < MIN_FAN_SPEED:
        return MIN_FAN_SPEED
    return val


class ECMailbox:
    """EC Mailbox interface using a TransientWinRing0 driver session."""

    def __init__(self, driver: TransientWinRing0):
        self.drv = driver

    def _inb(self, port):
        return self.drv.read_io_port_byte(port)

    def _outb(self, port, value):
        self.drv.write_io_port_byte(port, value)

    def _wait_ibe(self):
        for _ in range(TIMEOUT_ITERATIONS):
            if (self._inb(PORT_CMD) & 0x02) == 0:
                return
            time.sleep(POLL_DELAY)
        raise TimeoutError("EC: input buffer not empty (IBE timeout)")

    def _wait_obf(self):
        for _ in range(TIMEOUT_ITERATIONS):
            if (self._inb(PORT_CMD) & 0x01) == 1:
                return
            time.sleep(POLL_DELAY)
        raise TimeoutError("EC: output buffer not full (OBF timeout)")

    def _wait_obe(self):
        for _ in range(TIMEOUT_ITERATIONS):
            status = self._inb(PORT_CMD)
            if (status & 0x01) == 0:
                return
            self._inb(PORT_DATA)  # drain stale byte
            time.sleep(POLL_DELAY)
        raise TimeoutError("EC: output buffer not empty (OBE timeout)")

    def mbey(self, cmd, subcmd, arg):
        """Execute one EC mailbox transaction and return the result byte."""
        self._wait_ibe()
        self._wait_obe()
        self._outb(PORT_CMD, cmd)
        self._wait_ibe()
        self._outb(PORT_DATA, subcmd)
        self._wait_ibe()
        self._outb(PORT_DATA, arg)
        self._wait_ibe()
        self._wait_obf()
        return self._inb(PORT_DATA)

    def set_fan1(self, percent):
        percent = max(0, min(100, int(percent)))
        return self.mbey(CMD_FAN, SUBCMD_SET_FAN1, percent) == SUCCESS_CODE

    def set_fan2(self, percent):
        percent = max(0, min(100, int(percent)))
        return self.mbey(CMD_FAN, SUBCMD_SET_FAN2, percent) == SUCCESS_CODE

    def read_fan1(self):
        return self.mbey(CMD_FAN, SUBCMD_QUERY, QUERY_READ_FAN1)

    def read_fan2(self):
        return self.mbey(CMD_FAN, SUBCMD_QUERY, QUERY_READ_FAN2)

    def restore_auto(self):
        return self.mbey(CMD_FAN, SUBCMD_QUERY, QUERY_AUTO_MODE) == SUCCESS_CODE


# ==============================================================================
# High-Level FanController — Manages Transient Driver Sessions
# ==============================================================================

class FanController:
    """
    High-level fan controller that uses transient WinRing0 sessions.

    Each public method:
      1. Creates the driver service
      2. Performs the EC mailbox transaction
      3. Destroys the driver service immediately after

    This means the driver is never alive during idle periods, sleep
    transitions, or resume — eliminating the HAL crash.

    A threading lock prevents concurrent access.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self._driver_path = None  # cached for speed

    def _run(self, fn):
        """Run fn(ec: ECMailbox) inside a transient driver session."""
        with self._lock:
            drv = TransientWinRing0()
            drv.open()
            try:
                ec = ECMailbox(drv)
                return fn(ec)
            finally:
                drv.close()

    def read_fan1(self):
        return self._run(lambda ec: ec.read_fan1())

    def read_fan2(self):
        return self._run(lambda ec: ec.read_fan2())

    def read_fans(self):
        """Read both fans in a single transient driver session. Returns (f1, f2)."""
        def _both(ec):
            return ec.read_fan1(), ec.read_fan2()
        return self._run(_both)

    def set_fan1(self, percent):
        return self._run(lambda ec: ec.set_fan1(clamp_fan_speed(percent)))

    def set_fan2(self, percent):
        return self._run(lambda ec: ec.set_fan2(clamp_fan_speed(percent)))

    def set_fans(self, f1, f2):
        def _both(ec):
            r1 = ec.set_fan1(clamp_fan_speed(f1))
            r2 = ec.set_fan2(clamp_fan_speed(f2))
            return r1, r2
        return self._run(_both)

    def restore_auto(self):
        return self._run(lambda ec: ec.restore_auto())

    # Legacy compat — these are no-ops since there's no persistent session
    def open(self):
        pass

    def close(self):
        pass

    def stop_driver(self):
        pass

    def uninstall(self):
        """Remove driver files and service if somehow still registered."""
        # Clean up any leftover service
        sc = advapi32.OpenSCManagerW(None, None, SC_MANAGER_ALL_ACCESS)
        if sc:
            for name in (TransientWinRing0.SERVICE_NAME, "WinRing0_1_2_0"):
                svc = advapi32.OpenServiceW(sc, name, SERVICE_ALL_ACCESS)
                if svc:
                    sb = ctypes.create_string_buffer(28)
                    advapi32.ControlService(svc, SERVICE_CONTROL_STOP, sb)
                    advapi32.DeleteService(svc)
                    advapi32.CloseServiceHandle(svc)
            advapi32.CloseServiceHandle(sc)

        # Remove driver files
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        driver_dir = os.path.join(appdata, 'Yoga Fan Control')
        if os.path.exists(driver_dir):
            try:
                shutil.rmtree(driver_dir)
            except Exception:
                pass


# ==============================================================================
# Compat alias — GUI imports WinRing0 and ECMailbox by name
# ==============================================================================

# The GUI uses WMIWrapper now, but fan_control.py is still imported for is_admin()
# and the CLI. Keep this alias in case anything imports directly.
WinRing0 = FanController


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
    f1 = fc.read_fan1()
    f2 = fc.read_fan2()
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
            f1 = fc.read_fan1()
            f2 = fc.read_fan2()
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
            ok1, ok2 = fc.set_fans(fan1_pct, fan2_pct)
            actual1 = fc.read_fan1()
            actual2 = fc.read_fan2()
            print(f"\rTarget: {fan1_pct}%/{fan2_pct}%  |  Actual: {actual1}%/{actual2}%",
                  end="", flush=True)
            time.sleep(3)
    except KeyboardInterrupt:
        print()


def main():
    parser = argparse.ArgumentParser(
        description="EC Mailbox Fan Control — Lenovo Yoga Pro 9i Gen 9",
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
    _fc = FanController()

    try:
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
    except TimeoutError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    finally:
        restore_auto_on_exit()


if __name__ == "__main__":
    main()
