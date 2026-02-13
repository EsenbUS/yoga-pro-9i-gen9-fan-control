#!/usr/bin/env python3
"""
EC Mailbox Fan Control Tool — Lenovo Yoga Pro 9i Gen 9 (2024)

Controls fan speed via the EC mailbox interface at I/O ports 0x5C0/0x5C4.
Uses the WinRing0 kernel driver (already installed from NotebookFanControl)
for privileged I/O port access.

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

# ==============================================================================
# WinRing0 Driver Interface
# ==============================================================================

# IOCTLs for WinRing0 driver (from WinRing0 source code)
OLS_TYPE = 40000  # 0x9C40

def CTL_CODE(DeviceType, Function, Method, Access):
    return (DeviceType << 16) | (Access << 14) | (Function << 2) | Method

METHOD_BUFFERED = 0
FILE_ANY_ACCESS = 0
FILE_READ_ACCESS = 1
FILE_WRITE_ACCESS = 2

IOCTL_OLS_READ_IO_PORT_BYTE  = CTL_CODE(OLS_TYPE, 0x833, METHOD_BUFFERED, FILE_READ_ACCESS)
IOCTL_OLS_WRITE_IO_PORT_BYTE = CTL_CODE(OLS_TYPE, 0x836, METHOD_BUFFERED, FILE_WRITE_ACCESS)

# Windows API constants
GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
FILE_SHARE_READ = 0x01
FILE_SHARE_WRITE = 0x02
OPEN_EXISTING = 3
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

# Windows API functions
kernel32 = ctypes.windll.kernel32

CreateFileW = kernel32.CreateFileW
CreateFileW.restype = wintypes.HANDLE
CreateFileW.argtypes = [
    wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
    wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE
]

DeviceIoControl = kernel32.DeviceIoControl
DeviceIoControl.restype = wintypes.BOOL
DeviceIoControl.argtypes = [
    wintypes.HANDLE, wintypes.DWORD, wintypes.LPVOID, wintypes.DWORD,
    wintypes.LPVOID, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID
]

CloseHandle = kernel32.CloseHandle
CloseHandle.restype = wintypes.BOOL
CloseHandle.argtypes = [wintypes.HANDLE]


class WinRing0:
    """Interface to the WinRing0 kernel driver for I/O port access.

    Self-installs and manages the WinRing0 driver — no NBFC dependency needed.
    Looks for WinRing0x64.sys in the same directory as this script.
    """

    DEVICE_NAME = r"\\.\WinRing0_1_2_0"
    SERVICE_NAME = "WinRing0_1_2_0"
    DRIVER_FILENAME = "WinRing0x64.sys"

    # SCM constants
    SC_MANAGER_ALL_ACCESS = 0xF003F
    SERVICE_ALL_ACCESS = 0xF01FF
    SERVICE_KERNEL_DRIVER = 0x01
    SERVICE_DEMAND_START = 0x03
    SERVICE_ERROR_NORMAL = 0x01
    SERVICE_RUNNING = 0x04
    SERVICE_STOPPED = 0x01

    def __init__(self):
        self.handle = None
        self._service_installed_by_us = False

        # SC Manager API — must declare types for 64-bit correctness
        advapi32 = ctypes.windll.advapi32

        # SC_HANDLE is a pointer-sized value
        SC_HANDLE = wintypes.HANDLE

        advapi32.OpenSCManagerW.restype = SC_HANDLE
        advapi32.OpenSCManagerW.argtypes = [
            wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD
        ]

        advapi32.OpenServiceW.restype = SC_HANDLE
        advapi32.OpenServiceW.argtypes = [
            SC_HANDLE, wintypes.LPCWSTR, wintypes.DWORD
        ]

        advapi32.CreateServiceW.restype = SC_HANDLE
        advapi32.CreateServiceW.argtypes = [
            SC_HANDLE,          # hSCManager
            wintypes.LPCWSTR,   # lpServiceName
            wintypes.LPCWSTR,   # lpDisplayName
            wintypes.DWORD,     # dwDesiredAccess
            wintypes.DWORD,     # dwServiceType
            wintypes.DWORD,     # dwStartType
            wintypes.DWORD,     # dwErrorControl
            wintypes.LPCWSTR,   # lpBinaryPathName
            wintypes.LPCWSTR,   # lpLoadOrderGroup
            wintypes.LPVOID,    # lpdwTagId
            wintypes.LPCWSTR,   # lpDependencies
            wintypes.LPCWSTR,   # lpServiceStartName
            wintypes.LPCWSTR,   # lpPassword
        ]

        advapi32.StartServiceW.restype = wintypes.BOOL
        advapi32.StartServiceW.argtypes = [
            SC_HANDLE, wintypes.DWORD, wintypes.LPVOID
        ]

        advapi32.DeleteService.restype = wintypes.BOOL
        advapi32.DeleteService.argtypes = [SC_HANDLE]

        advapi32.ChangeServiceConfigW.restype = wintypes.BOOL
        advapi32.ChangeServiceConfigW.argtypes = [
            SC_HANDLE,          # hService
            wintypes.DWORD,     # dwServiceType
            wintypes.DWORD,     # dwStartType
            wintypes.DWORD,     # dwErrorControl
            wintypes.LPCWSTR,   # lpBinaryPathName
            wintypes.LPCWSTR,   # lpLoadOrderGroup
            wintypes.LPVOID,    # lpdwTagId
            wintypes.LPCWSTR,   # lpDependencies
            wintypes.LPCWSTR,   # lpServiceStartName
            wintypes.LPCWSTR,   # lpPassword
            wintypes.LPCWSTR,   # lpDisplayName
        ]

        advapi32.ControlService.restype = wintypes.BOOL
        advapi32.ControlService.argtypes = [
            SC_HANDLE, wintypes.DWORD, ctypes.c_void_p
        ]

        advapi32.CloseServiceHandle.restype = wintypes.BOOL
        advapi32.CloseServiceHandle.argtypes = [SC_HANDLE]

        self._advapi32 = advapi32

    @staticmethod
    def _get_permanent_driver_dir():
        """Get (or create) a permanent directory for the driver .sys file."""
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        driver_dir = os.path.join(appdata, 'FanControl')
        os.makedirs(driver_dir, exist_ok=True)
        return driver_dir

    def _find_driver(self):
        """Find WinRing0x64.sys, copying to permanent location if needed."""
        perm_dir = self._get_permanent_driver_dir()
        perm_path = os.path.join(perm_dir, self.DRIVER_FILENAME)

        # If already in permanent location, use it
        if os.path.exists(perm_path):
            return perm_path

        # Find the source file
        source = None
        if getattr(sys, '_MEIPASS', None):
            candidate = os.path.join(sys._MEIPASS, self.DRIVER_FILENAME)
            if os.path.exists(candidate):
                source = candidate
        if not source:
            if getattr(sys, 'frozen', False):
                d = os.path.dirname(sys.executable)
            else:
                d = os.path.dirname(os.path.abspath(__file__))
            candidate = os.path.join(d, self.DRIVER_FILENAME)
            if os.path.exists(candidate):
                source = candidate

        if source:
            shutil.copy2(source, perm_path)
            return perm_path

        raise FileNotFoundError(
            f"Cannot find {self.DRIVER_FILENAME}.\n"
            f"Please place WinRing0x64.sys in the same folder as this program."
        )

    def _install_driver(self):
        """Install and start the WinRing0 driver service."""
        driver_path = self._find_driver()

        # Open Service Control Manager
        sc_manager = self._advapi32.OpenSCManagerW(None, None, self.SC_MANAGER_ALL_ACCESS)
        if not sc_manager:
            raise RuntimeError(
                f"Cannot open Service Control Manager (error {kernel32.GetLastError()}).\n"
                f"Make sure you are running as Administrator."
            )

        try:
            # Try to open existing service first
            service = self._advapi32.OpenServiceW(
                sc_manager, self.SERVICE_NAME, self.SERVICE_ALL_ACCESS
            )

            if service:
                # Service exists — update its config (path, start type)
                SERVICE_NO_CHANGE = 0xFFFFFFFF
                self._advapi32.ChangeServiceConfigW(
                    service,
                    SERVICE_NO_CHANGE,           # type: no change
                    self.SERVICE_DEMAND_START,   # ensure not disabled
                    SERVICE_NO_CHANGE,           # error: no change
                    driver_path,                 # update path
                    None, None, None, None, None, None
                )
            else:
                # Service doesn't exist — create it
                service = self._advapi32.CreateServiceW(
                    sc_manager,
                    self.SERVICE_NAME,
                    self.SERVICE_NAME,
                    self.SERVICE_ALL_ACCESS,
                    self.SERVICE_KERNEL_DRIVER,
                    self.SERVICE_DEMAND_START,
                    self.SERVICE_ERROR_NORMAL,
                    driver_path,
                    None, None, None, None, None
                )
                if not service:
                    raise RuntimeError(
                        f"Cannot create driver service (error {kernel32.GetLastError()})."
                    )
                self._service_installed_by_us = True

            # Start the service (ignore error if already running)
            if not self._advapi32.StartServiceW(service, 0, None):
                error = kernel32.GetLastError()
                ERROR_SERVICE_ALREADY_RUNNING = 1056
                if error != ERROR_SERVICE_ALREADY_RUNNING:
                    raise RuntimeError(
                        f"Cannot start WinRing0 driver (error {error})."
                    )

            self._advapi32.CloseServiceHandle(service)
        finally:
            self._advapi32.CloseServiceHandle(sc_manager)

    def open(self):
        """Open a handle to the WinRing0 driver, installing it if needed."""
        # First try to open directly (driver may already be running)
        self.handle = CreateFileW(
            self.DEVICE_NAME,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None,
            OPEN_EXISTING,
            0,
            None
        )
        if self.handle != INVALID_HANDLE_VALUE:
            return  # Already running (e.g. from a previous session)

        # Driver not running — install and start it ourselves
        self._install_driver()

        # Try opening again
        self.handle = CreateFileW(
            self.DEVICE_NAME,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None,
            OPEN_EXISTING,
            0,
            None
        )
        if self.handle == INVALID_HANDLE_VALUE:
            error = ctypes.get_last_error() or kernel32.GetLastError()
            raise RuntimeError(
                f"Cannot open WinRing0 driver after installation (error {error}).\n"
                f"Make sure you are running as Administrator."
            )

    def close(self):
        """Close the driver handle."""
        if self.handle and self.handle != INVALID_HANDLE_VALUE:
            CloseHandle(self.handle)
            self.handle = None

    def stop_driver(self):
        """Stop the WinRing0 driver service (needed before sleep)."""
        try:
            sc_manager = self._advapi32.OpenSCManagerW(
                None, None, self.SC_MANAGER_ALL_ACCESS
            )
            if not sc_manager:
                return
            service = self._advapi32.OpenServiceW(
                sc_manager, self.SERVICE_NAME, self.SERVICE_ALL_ACCESS
            )
            if service:
                # SERVICE_STATUS is 7 DWORDs = 28 bytes
                status_buf = ctypes.create_string_buffer(28)
                SERVICE_CONTROL_STOP = 0x01
                self._advapi32.ControlService(service, SERVICE_CONTROL_STOP, status_buf)
                self._advapi32.CloseServiceHandle(service)
            self._advapi32.CloseServiceHandle(sc_manager)
        except Exception:
            pass

    def read_io_port_byte(self, port):
        """Read a single byte from an I/O port."""
        # Input: 4 bytes (DWORD port number)
        in_buf = struct.pack('<I', port)
        out_buf = ctypes.create_string_buffer(4)
        bytes_returned = wintypes.DWORD(0)

        success = DeviceIoControl(
            self.handle,
            IOCTL_OLS_READ_IO_PORT_BYTE,
            in_buf, len(in_buf),
            out_buf, len(out_buf),
            ctypes.byref(bytes_returned),
            None
        )
        if not success:
            raise RuntimeError(f"Failed to read I/O port 0x{port:X}")

        return struct.unpack('<I', out_buf.raw)[0] & 0xFF

    def write_io_port_byte(self, port, value):
        """Write a single byte to an I/O port."""
        # Input: 8 bytes (DWORD port, DWORD value)
        in_buf = struct.pack('<II', port, value & 0xFF)
        bytes_returned = wintypes.DWORD(0)

        success = DeviceIoControl(
            self.handle,
            IOCTL_OLS_WRITE_IO_PORT_BYTE,
            in_buf, len(in_buf),
            None, 0,
            ctypes.byref(bytes_returned),
            None
        )
        if not success:
            raise RuntimeError(f"Failed to write 0x{value:02X} to I/O port 0x{port:X}")

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, *args):
        self.close()


# ==============================================================================
# EC Mailbox Protocol
# ==============================================================================

PORT_DATA = 0x5C0  # X5C0: data register
PORT_CMD  = 0x5C4  # X5C4: command/status register

TIMEOUT_ITERATIONS = 0x10000  # Matches DSDT WIBE/WOBF/WOBE loop count
POLL_DELAY = 0.00001  # 10 microseconds between polls

CMD_FAN = 0xEF  # EC fan control command byte

# Sub-commands
SUBCMD_SET_FAN1 = 0x61   # Set Fan 1 duty cycle
SUBCMD_SET_FAN2 = 0x62   # Set Fan 2 duty cycle
SUBCMD_QUERY    = 0x63   # Query / restore

# Query arguments
QUERY_READ_FAN1  = 0x01  # Read Fan 1 current speed
QUERY_READ_FAN2  = 0x02  # Read Fan 2 current speed
QUERY_AUTO_MODE  = 0x03  # Restore automatic mode

SUCCESS_CODE = 0xAC  # MBEY returns this on success

# Safety limits
MIN_FAN_SPEED = 18   # Values 1-17 cause pulsing; 0 (off) is fine
SAFE_MAX = 48        # EC's normal maximum; above this is unusual

def clamp_fan_speed(val):
    """Clamp to valid range: 0 or 18-100 (1-17 causes fan pulsing)."""
    val = max(0, min(100, int(val)))
    if 1 <= val < MIN_FAN_SPEED:
        return MIN_FAN_SPEED
    return val


class ECMailbox:
    """EC Mailbox interface for fan control via WinRing0 I/O."""

    def __init__(self, driver):
        self.drv = driver

    def _inb(self, port):
        return self.drv.read_io_port_byte(port)

    def _outb(self, port, value):
        self.drv.write_io_port_byte(port, value)

    def wait_ibe(self):
        """Wait for Input Buffer Empty: bit 1 of status == 0."""
        for _ in range(TIMEOUT_ITERATIONS):
            status = self._inb(PORT_CMD)
            if (status & 0x02) == 0:
                return True
            time.sleep(POLL_DELAY)
        raise TimeoutError("EC mailbox: input buffer not empty (WIBE timeout)")

    def wait_obf(self):
        """Wait for Output Buffer Full: bit 0 of status == 1."""
        for _ in range(TIMEOUT_ITERATIONS):
            status = self._inb(PORT_CMD)
            if (status & 0x01) == 1:
                return True
            time.sleep(POLL_DELAY)
        raise TimeoutError("EC mailbox: output buffer not full (WOBF timeout)")

    def wait_obe(self):
        """Wait for Output Buffer Empty: bit 0 of status == 0.
        If data is pending, drain it."""
        for _ in range(TIMEOUT_ITERATIONS):
            status = self._inb(PORT_CMD)
            if (status & 0x01) == 0:
                return True
            # Drain stale data
            _ = self._inb(PORT_DATA)
            time.sleep(POLL_DELAY)
        raise TimeoutError("EC mailbox: output buffer not empty (WOBE timeout)")

    def mbey(self, cmd, subcmd, arg):
        """
        Execute a full MBEY mailbox transaction.

        Protocol (from DSDT):
          1. WIBE  — wait for input buffer empty
          2. WOBE  — wait for output buffer empty (drain stale)
          3. Write cmd to X5C4 (command port)
          4. WIBE  — wait for input buffer empty
          5. Write subcmd to X5C0 (data port)
          6. WIBE  — wait for input buffer empty
          7. Write arg to X5C0 (data port)
          8. WIBE  — wait for input buffer empty
          9. WOBF  — wait for output buffer full
         10. Read result from X5C0 (data port)

        Returns the result byte.
        """
        self.wait_ibe()
        self.wait_obe()
        self._outb(PORT_CMD, cmd)
        self.wait_ibe()
        self._outb(PORT_DATA, subcmd)
        self.wait_ibe()
        self._outb(PORT_DATA, arg)
        self.wait_ibe()
        self.wait_obf()
        result = self._inb(PORT_DATA)
        return result

    def set_fan1(self, percent):
        """Set Fan 1 duty cycle (0-100%)."""
        percent = max(0, min(100, int(percent)))
        result = self.mbey(CMD_FAN, SUBCMD_SET_FAN1, percent)
        return result == SUCCESS_CODE

    def set_fan2(self, percent):
        """Set Fan 2 duty cycle (0-100%)."""
        percent = max(0, min(100, int(percent)))
        result = self.mbey(CMD_FAN, SUBCMD_SET_FAN2, percent)
        return result == SUCCESS_CODE

    def read_fan1(self):
        """Read Fan 1 current speed. Returns percentage value."""
        return self.mbey(CMD_FAN, SUBCMD_QUERY, QUERY_READ_FAN1)

    def read_fan2(self):
        """Read Fan 2 current speed. Returns percentage value."""
        return self.mbey(CMD_FAN, SUBCMD_QUERY, QUERY_READ_FAN2)

    def restore_auto(self):
        """Restore automatic EC fan control."""
        result = self.mbey(CMD_FAN, SUBCMD_QUERY, QUERY_AUTO_MODE)
        return result == SUCCESS_CODE


# ==============================================================================
# CLI Interface
# ==============================================================================

# Global reference for signal handler cleanup
_ec_mailbox = None
_auto_restore_on_exit = False


def restore_auto_on_exit():
    """Safety: restore automatic fan control when the program exits."""
    global _ec_mailbox, _auto_restore_on_exit
    if _ec_mailbox and _auto_restore_on_exit:
        try:
            _ec_mailbox.restore_auto()
            print("\n[Safety] Restored automatic fan control.")
        except Exception:
            pass


def signal_handler(sig, frame):
    """Handle Ctrl+C gracefully."""
    restore_auto_on_exit()
    sys.exit(0)


def is_admin():
    """Check if running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def cmd_read(ec):
    """Read and display current fan speeds."""
    fan1 = ec.read_fan1()
    fan2 = ec.read_fan2()
    print(f"Fan 1: {fan1}%")
    print(f"Fan 2: {fan2}%")


def cmd_set(ec, fan1_pct, fan2_pct=None):
    """Set fan speed(s)."""
    if fan2_pct is None:
        fan2_pct = fan1_pct

    # Enforce minimum (0 is fine, 1-17 clamps to 18)
    fan1_pct = clamp_fan_speed(fan1_pct)
    fan2_pct = clamp_fan_speed(fan2_pct)

    # Warn above safe max
    if fan1_pct > SAFE_MAX or fan2_pct > SAFE_MAX:
        print(f"WARNING: Setting fans above {SAFE_MAX}% exceeds the EC's normal range.")
        print(f"  Fan 1: {fan1_pct}%  |  Fan 2: {fan2_pct}%")
        confirm = input("Continue? (y/N): ").strip().lower()
        if confirm != 'y':
            print("Aborted.")
            return

    print(f"Setting Fan 1 to {fan1_pct}%...", end=" ")
    ok1 = ec.set_fan1(fan1_pct)
    print("OK" if ok1 else "FAILED")

    print(f"Setting Fan 2 to {fan2_pct}%...", end=" ")
    ok2 = ec.set_fan2(fan2_pct)
    print("OK" if ok2 else "FAILED")

    if ok1 and ok2:
        print("\nFan speeds set successfully.")
        print("Note: The EC may override these values. Use 'hold' mode to maintain them.")
    else:
        print("\nWarning: One or more fan commands did not return success code.")


def cmd_auto(ec):
    """Restore automatic fan control."""
    print("Restoring automatic fan control...", end=" ")
    ok = ec.restore_auto()
    print("OK" if ok else "FAILED")


def cmd_monitor(ec):
    """Continuously monitor fan speeds."""
    global _auto_restore_on_exit
    _auto_restore_on_exit = False  # Monitor is read-only, no need to restore

    print("Monitoring fan speeds (Ctrl+C to stop)...\n")
    try:
        while True:
            fan1 = ec.read_fan1()
            fan2 = ec.read_fan2()
            print(f"\rFan 1: {fan1:3d}%  |  Fan 2: {fan2:3d}%", end="", flush=True)
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\nMonitoring stopped.")


def cmd_hold(ec, fan1_pct, fan2_pct=None):
    """Set and continuously maintain fan speeds."""
    global _auto_restore_on_exit
    _auto_restore_on_exit = True

    if fan2_pct is None:
        fan2_pct = fan1_pct

    # Enforce minimum (0 is fine, 1-17 clamps to 18)
    fan1_pct = clamp_fan_speed(fan1_pct)
    fan2_pct = clamp_fan_speed(fan2_pct)

    # Warn above safe max
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
            ec.set_fan1(fan1_pct)
            ec.set_fan2(fan2_pct)

            # Read back actual values
            actual1 = ec.read_fan1()
            actual2 = ec.read_fan2()
            print(f"\rTarget: {fan1_pct}%/{fan2_pct}%  |  Actual: {actual1}%/{actual2}%", end="", flush=True)

            time.sleep(3)  # Re-send every 3 seconds to prevent EC override
    except KeyboardInterrupt:
        print()  # newline before cleanup message


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
  hold <f1> [f2]    Set and maintain fan speeds (re-sends periodically)

Examples:
  python fan_control.py read
  python fan_control.py set 70
  python fan_control.py set 50 60
  python fan_control.py hold 40 40
  python fan_control.py auto
        """
    )
    parser.add_argument("command", choices=["read", "set", "auto", "monitor", "hold"],
                        help="Command to execute")
    parser.add_argument("fan1", nargs="?", type=int, default=None,
                        help="Fan 1 speed percentage (0-100)")
    parser.add_argument("fan2", nargs="?", type=int, default=None,
                        help="Fan 2 speed percentage (0-100, defaults to fan1)")

    args = parser.parse_args()

    # Validate arguments
    if args.command in ("set", "hold") and args.fan1 is None:
        parser.error(f"'{args.command}' requires at least one fan speed percentage")

    if args.fan1 is not None and not (0 <= args.fan1 <= 100):
        parser.error("Fan 1 speed must be between 0 and 100")

    if args.fan2 is not None and not (0 <= args.fan2 <= 100):
        parser.error("Fan 2 speed must be between 0 and 100")

    # Check admin
    if not is_admin():
        print("ERROR: This tool must be run as Administrator.")
        print("Right-click your terminal and select 'Run as administrator'.")
        sys.exit(1)

    # Set up signal handler
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGBREAK, signal_handler)

    # Connect to driver and execute command
    global _ec_mailbox
    driver = WinRing0()

    try:
        driver.open()
        print("WinRing0 driver connected.\n")
    except RuntimeError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    ec = ECMailbox(driver)
    _ec_mailbox = ec

    try:
        if args.command == "read":
            cmd_read(ec)
        elif args.command == "set":
            cmd_set(ec, args.fan1, args.fan2)
        elif args.command == "auto":
            cmd_auto(ec)
        elif args.command == "monitor":
            cmd_monitor(ec)
        elif args.command == "hold":
            cmd_hold(ec, args.fan1, args.fan2)
    except TimeoutError as e:
        print(f"\nERROR: {e}")
        print("The EC did not respond in time. This may indicate:")
        print("  - The mailbox ports (0x5C0/0x5C4) are not correct for your system")
        print("  - The EC is busy or in a bad state")
        print("  - Try rebooting and running again")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    finally:
        restore_auto_on_exit()
        driver.close()


if __name__ == "__main__":
    main()
