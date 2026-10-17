# Fan Control - Lenovo Yoga Pro 9i Gen 9

<p align="center">
  <img src="https://cdn.imgchest.com/files/d7be7615fcfb.png" alt="Yoga Fan Control GUI" width="600">
</p>

A standalone fan control utility for the **Lenovo Yoga Pro 9i Gen 9 (2024)** with a dark-themed GUI, live monitoring, and full manual override of the EC (Embedded Controller).

> **No drivers, no third-party services** - this tool calls the fan-control method that Lenovo's own firmware exposes through WMI (`root\WMI` → `LENOVO_OTHER_METHOD`). The firmware does the EC access itself, with proper locking.

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
| **System Tray** | Minimizes to tray on close; quit only via right-click → Quit |
| **Shutdown Safe** | Restores auto fan control on shutdown/restart |
| **Standalone .exe** | Single executable - no Python installation required |
| **No Kernel Driver** | Uses the firmware's WMI interface; no WinRing0, no antivirus exclusions |
| **Sleep Safe** | Stops all EC access while the display is off or the system is sleeping |

---

## Quick Start

### Option 1: Pre-built Executable (Recommended)

1. Download `Fan.Control.exe` from [Releases](https://github.com/EsenbUS/yoga-pro-9i-gen9-fan-control/releases) and run
2. Accept the UAC (Administrator)
3. Control your fans

### Option 2: Run from Source

```bash
# Clone the repository
git clone https://github.com/EsenbUS/yoga-pro-9i-gen9-fan-control.git

# Install dependencies
pip install psutil Pillow pystray pywin32

# Run as Administrator
python fan_control_gui.py
```

> **Must be run as Administrator** - Lenovo's WMI fan interface denies access to non-admin processes.

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

- **On Shutdown/Restart** → Auto fan control is restored before the system powers down (if the program is running)
- **On App Quit** → Auto fan control is restored when you exit via the tray
- **On Display Off / Sleep** → Auto fan control is restored when the display turns off, and the EC is not touched again until the display has been back on for 5 seconds. Your manual speeds are then re-applied. This keeps the app out of the way of Modern Standby and hibernate transitions.

---

## How It Works

The firmware exposes a WMI method (`LENOVO_OTHER_METHOD`, ACPI `\_SB.GZFD.WMAE`) that Lenovo's own software uses. Internally it runs the **EC Mailbox Protocol** on ports `0x5C0`/`0x5C4` while holding the EC mutex and the ACPI Global Lock, so calls can't collide with other EC traffic.

| Feature ID | `GetFeatureValue` | `SetFeatureValue` |
|------------|-------------------|-------------------|
| `0x04030001` | Fan 1 speed × 100 | `0` = auto, `1` = off, `N` = `N / 100` % |
| `0x04030002` | Fan 2 speed × 100 | same as Fan 1 |

Earlier versions wrote to the mailbox ports directly through the WinRing0 driver. That bypassed the firmware's locking and could leave the EC mid-command around sleep transitions, which caused failed hibernates and full battery drains. The app removes the leftover `WinRing0_Transient` service and driver copy on startup.

---

## Project Structure

```
Fan_control/
├── fan_control.py          # Backend: Lenovo WMI fan interface + CLI
├── fan_control_gui.py      # Frontend: Tkinter GUI with tray, presets, gauges
├── fan_config.json         # User settings (auto-generated, not committed)
└── README.md
```

---

## Building the Standalone Executable

```bash
pip install pyinstaller Pillow pystray psutil pywin32

pyinstaller --onefile --noconsole --uac-admin \
  --name "Fan Control" \
  --icon "fan_control.ico" \ # Optional
\
  --add-data "fan_control.py;." \
  fan_control_gui.py
```

The output will be in `dist/Fan Control.exe`.

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
