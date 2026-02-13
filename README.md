> **Project uses WinRing0, so you are required to exclude it from Windows Defender/whatever antivirus you use if you want to use this program. 
See [Microsoft Article](https://support.microsoft.com/en-us/windows/microsoft-defender-antivirus-alert-vulnerabledriver-winnt-winring0-eb057830-d77b-41a2-9a34-015a5d203c42)**

> Additionally, for safety's sake before using this program update BIOS to the latest (NKCN32WW) using this website [Lenovo Drivers](https://pcsupport.lenovo.com/us/en/products/laptops-and-netbooks/yoga-series/yoga-pro-9-16imh9/83dn/83dn0008us/pf4zs51d/downloads/driver-list)

# Fan Control - Lenovo Yoga Pro 9i Gen 9

<p align="center">
  <img src="https://cdn.imgchest.com/files/b6bbc12244b4.png" alt="Yoga Fan Control GUI" width="600">
</p>

A standalone fan control utility for the **Lenovo Yoga Pro 9i Gen 9 (2024)** with a dark-themed GUI, live monitoring, and full manual override of the EC (Embedded Controller).

> **No third-party services required** - this tool communicates directly with the EC via the mailbox protocol on I/O ports `0x5C0`/`0x5C4`, using the WinRing0 kernel driver for privileged I/O access.

![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

---

## Features

| Feature | Description |
|---------|-------------|
| **Manual Fan Control** | Dual sliders for Fan 1 and Fan 2, with optional linking |
| **Live Monitoring** | Real-time arc gauges showing current fan speed |
| **Presets** | Built-in presets (OFF / Min 18% / Med 22% / Med-High 30% / High 48%) |
| **Custom Presets** | Create your own presets with validation and persistence. Use right-click to delete a custom preset|
| **Hold Mode** | Re-sends fan speed every 3s to prevent EC override (not really necessary for work) |
| **System Tray** | Minimizes to tray on close; quit only via right-click → Quit |
| **Sleep Safe** | Automatically restores auto fan control before system sleep |
| **Shutdown Safe** | Restores auto fan control on shutdown/restart |
| **Standalone .exe** | Single executable - no Python installation required |
| **Self-contained Driver** | Auto-installs WinRing0 driver |
| **Easy Uninstall** | Includes script for driver uninstallation |


---

## Quick Start

### Option 1: Pre-built Executable (Recommended)

1. Download `Yoga Fan Control.exe` and run
2. Accept the UAC (Administrator)
3. Control your fans

### Option 2: Run from Source

```bash
# Clone the repository
git clone https://github.com/EsenbUS/yoga-pro-9i-gen9-fan-control.git

# Install dependencies
pip install psutil Pillow pystray

# Run as Administrator
python fan_control_gui.py
```

> **Must be run as Administrator** - privileged I/O port access requires elevation.

---

## Fan Speed Safety

| Range | Behavior |
|-------|----------|
| **0%** | Fans completely off (use with caution!) |
| **1–17%** | **Blocked** - causes fan pulsing/instability |
| **18–48%** | Normal operating range |
| **49–100%** | Warning prompt - exceeds EC's normal maximum |

The tool enforces these ranges automatically. The 1–17% dead zone is a hardware limitation where the fans pulse erratically.

---

## Safety Features

This tool is designed to **never leave your laptop in a dangerous state**:

- **On Sleep** → Auto fan control is restored before the system suspends
- **On Shutdown/Restart** → Auto fan control is restored before the system powers down
- **On App Quit** → Auto fan control is restored when you exit via the tray

---

## How It Works

The tool communicates with the laptop's Embedded Controller using the **EC Mailbox Protocol**, reverse-engineered from the DSDT (ACPI firmware tables):

```
Port 0x5C0 - Data register (read/write)
Port 0x5C4 - Command/status register
```

**Protocol sequence:**
1. Wait for Input Buffer Empty (IBE) on `0x5C4`
2. Wait for Output Buffer Empty (OBE) on `0x5C4`
3. Write command byte to `0x5C4`
4. Write sub-command to `0x5C0`
5. Write data to `0x5C0`
6. Wait for Output Buffer Full (OBF) on `0x5C4`
7. Read result from `0x5C0`

| Command | Sub-cmd | Description |
|---------|---------|-------------|
| `0xD0` | `0x61` | Set Fan 1 speed (0–255) |
| `0xD0` | `0x62` | Set Fan 2 speed (0–255) |
| `0xD0` | `0x21` | Read Fan 1 speed |
| `0xD0` | `0x22` | Read Fan 2 speed |
| `0xD1` | `0x61` / `0x62` | Restore auto control |

---

## Project Structure

```
Fan_control/
├── fan_control.py          # Backend: EC mailbox protocol + WinRing0 driver management
├── fan_control_gui.py      # Frontend: Tkinter GUI with tray, presets, gauges
├── WinRing0x64.sys         # Kernel driver for I/O port access (required)
├── fan_config.json         # User settings (auto-generated, not committed)
└── README.md
```

---

## Building the Standalone Executable

```bash
pip install pyinstaller Pillow pystray psutil

pyinstaller --onefile --noconsole --uac-admin \
  --name "Fan Control" \
  --icon "fan_control.ico" \
  --add-data "WinRing0x64.sys;." \
  --add-data "fan_control.py;." \
  fan_control_gui.py
```

The output will be in `dist/Yoga Fan Control.exe`.

---

## CLI Mode

The backend script also provides a standalone CLI:

```bash
# Read current fan speeds
python fan_control.py read

# Set both fans to 30%
python fan_control.py set --fan1 30 --fan2 30

# Restore automatic fan control
python fan_control.py auto

# Monitor fan speeds in real-time
python fan_control.py monitor

# Hold fans at a specific speed (re-sends every 3s)
python fan_control.py hold --fan1 25 --fan2 25
```

---

## Disclaimer

This tool directly communicates with your laptop's Embedded Controller at the hardware level. While it has been tested on the **Lenovo Yoga Pro 9i Gen 9 (2024)**, use it at your own risk.

- **Only tested on**: Lenovo Yoga Pro 9i Gen 9 (2024)
- **May work on**: Other Lenovo laptops with the same EC mailbox protocol (no responsibility for any issues)
- **Will NOT work on**: Non-Lenovo laptops, or Lenovo models with different EC interfaces

> **If your laptop model has a different EC interface, using this tool could produce unpredictable results. Always ensure you have a way to hard-reset your laptop (hold power button for 10 seconds) before experimenting.**

---

## Contributing

Contributions are welcome! If you've tested this on another Lenovo model, please open an issue to report compatibility.

---

## License

MIT License - see [LICENSE](LICENSE) for details.
