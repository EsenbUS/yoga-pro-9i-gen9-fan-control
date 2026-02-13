# 🌀 Fan Control — Lenovo Yoga Pro 9i Gen 9

A standalone fan control utility for the **Lenovo Yoga Pro 9i Gen 9 (2024)** with a dark-themed GUI, live monitoring, and full manual override of the EC (Embedded Controller).

> **No Notebook Fan Control (NBFC), no third-party services** — this tool communicates directly with the EC via the mailbox protocol on I/O ports `0x5C0`/`0x5C4`, using the WinRing0 kernel driver for privileged I/O access.

![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🎛️ **Manual Fan Control** | Dual sliders for Fan 1 and Fan 2, with optional linking |
| 📊 **Live Monitoring** | Real-time arc gauges showing current fan RPM + CPU temperature |
| ⚡ **Presets** | Built-in presets (OFF / Min 18% / Med 22% / Med-High 30% / High 48%) |
| ➕ **Custom Presets** | Create your own presets with validation and persistence |
| 🔄 **Hold Mode** | Re-sends fan speed every 3s to prevent EC override |
| 🌙 **System Tray** | Minimizes to tray on close; quit only via right-click → Quit |
| 😴 **Sleep Safe** | Automatically restores auto fan control before system sleep |
| 🔌 **Shutdown Safe** | Restores auto fan control on shutdown/restart |
| 🛡️ **Boot Safety Net** | Scheduled task restores auto mode on every boot (crash protection) |
| 📦 **Standalone .exe** | Single executable — no Python installation required |
| 🔧 **Self-contained Driver** | Auto-installs WinRing0 driver — no NBFC dependency |

---

## 🚀 Quick Start

### Option 1: Pre-built Executable (Recommended)

1. Download `Fan.Control.exe` and `WinRing0x64.sys` from [Releases](../../releases)
2. Place both files in the same folder
3. Double-click `Fan Control.exe`
4. Accept the UAC (Administrator) prompt
5. Control your fans! 🎉

### Option 2: Run from Source

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/Fan_control.git
cd Fan_control

# Install dependencies
pip install psutil Pillow pystray

# Run as Administrator
python fan_control_gui.py
```

> ⚠️ **Must be run as Administrator** — privileged I/O port access requires elevation.

---

## 🎛️ Fan Speed Safety

| Range | Behavior |
|-------|----------|
| **0%** | Fans completely off (use with caution!) |
| **1–17%** | ❌ **Blocked** — causes fan pulsing/instability |
| **18–48%** | ✅ Normal operating range |
| **49–100%** | ⚠️ Warning prompt — exceeds EC's normal maximum |

The tool enforces these ranges automatically. The 1–17% dead zone is a hardware limitation where the fans pulse erratically.

---

## 🛡️ Safety Features

This tool is designed to **never leave your laptop in a dangerous state**:

- **On Sleep** → Auto fan control is restored before the system suspends
- **On Shutdown/Restart** → Auto fan control is restored before the system powers down
- **On App Quit** → Auto fan control is restored when you exit via the tray
- **On Boot** → A scheduled task (`FanControlAutoRestore`) restores auto mode on every system startup, protecting against crashes or unexpected reboots

---

## 🔧 How It Works

The tool communicates with the laptop's Embedded Controller using the **EC Mailbox Protocol**, reverse-engineered from the DSDT (ACPI firmware tables):

```
Port 0x5C0 — Data register (read/write)
Port 0x5C4 — Command/status register
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

## 📁 Project Structure

```
Fan_control/
├── fan_control.py          # Backend: EC mailbox protocol + WinRing0 driver management
├── fan_control_gui.py      # Frontend: Tkinter GUI with tray, presets, gauges
├── Fan Control.vbs         # Optional VBS launcher (auto-elevates, no console)
├── WinRing0x64.sys         # Kernel driver for I/O port access (required)
├── fan_config.json         # User settings (auto-generated, not committed)
└── README.md
```

---

## 🏗️ Building the Standalone Executable

```bash
pip install pyinstaller Pillow pystray psutil

pyinstaller --onefile --noconsole --uac-admin \
  --name "Fan Control" \
  --icon "fan_control.ico" \
  --add-data "WinRing0x64.sys;." \
  --add-data "fan_control.py;." \
  fan_control_gui.py
```

The output will be in `dist/Fan Control.exe`.

---

## 📝 CLI Mode

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

## ⚠️ Disclaimer

This tool directly communicates with your laptop's Embedded Controller at the hardware level. While it has been tested on the **Lenovo Yoga Pro 9i Gen 9 (83GO)**, use it at your own risk.

- **Only tested on**: Lenovo Yoga Pro 9i Gen 9 (2024)
- **May work on**: Other Lenovo laptops with the same EC mailbox protocol
- **Will NOT work on**: Non-Lenovo laptops, or Lenovo models with different EC interfaces

> **If your laptop model has a different EC interface, using this tool could produce unpredictable results. Always ensure you have a way to hard-reset your laptop (hold power button for 10 seconds) before experimenting.**

---

## 🤝 Contributing

Contributions are welcome! If you've tested this on another Lenovo model, please open an issue to report compatibility.

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.
