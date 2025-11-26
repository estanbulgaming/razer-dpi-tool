"""
DPI Tool - Razer & Logitech
Simple tool to change DPI and audio EQ without vendor software
"""

import hid
import tkinter as tk
from tkinter import ttk, messagebox
import struct
import os
import shutil
from pathlib import Path

# Vendor IDs
RAZER_VENDOR_ID = 0x1532
LOGITECH_VENDOR_ID = 0x046D

# EqualizerAPO paths
EQUALIZER_APO_PATH = Path(os.environ.get('PROGRAMFILES', 'C:\\Program Files')) / 'EqualizerAPO' / 'config'
EQUALIZER_APO_CONFIG = EQUALIZER_APO_PATH / 'config.txt'

# Known Razer Mouse Product IDs (from OpenRazer)
KNOWN_RAZER_MICE = {
    0x0013: "Orochi 2011",
    0x0016: "DeathAdder 3.5G",
    0x0020: "DeathAdder 1800",
    0x0024: "Mamba 2012 (Wired)",
    0x0025: "Mamba 2012 (Wireless)",
    0x0029: "DeathAdder 3.5G Black",
    0x002E: "Naga 2012",
    0x002F: "Naga 2012 Lefty",
    0x0032: "Naga Hex",
    0x0034: "Abyssus 1800",
    0x0036: "DeathAdder 2013",
    0x0037: "DeathAdder Chroma",
    0x0038: "Naga 2014",
    0x0039: "Naga Hex V2",
    0x003E: "Orochi 2013",
    0x003F: "Naga Epic Chroma",
    0x0040: "Mamba (Wired)",
    0x0041: "Mamba Chroma (Wired)",
    0x0042: "Abyssus V2",
    0x0043: "DeathAdder 3500",
    0x0044: "Mamba TE",
    0x0045: "Mamba Chroma (Wireless)",
    0x0046: "Naga Chroma",
    0x0048: "Diamondback Chroma",
    0x004C: "Abyssus 2000",
    0x004F: "Deathadder V2 X HyperSpeed",
    0x0050: "Naga Left-Handed",
    0x0053: "Naga Trinity",
    0x0054: "Abyssus Elite (D.Va)",
    0x005B: "DeathAdder Elite",
    0x005C: "DeathAdder 1500",
    0x005E: "Lancehead Wired",
    0x0060: "Lancehead Wireless",
    0x0062: "Naga Hex V2 (Overwatch)",
    0x0064: "Abyssus Essential",
    0x0065: "Basilisk",
    0x0067: "Naga X",
    0x006A: "Viper Mini",
    0x006B: "DeathAdder Essential White",
    0x006C: "Mamba Elite",
    0x006E: "DeathAdder Essential",
    0x006F: "Lancehead TE",
    0x0070: "Atheris (Receiver)",
    0x0071: "Atheris (Bluetooth)",
    0x0072: "Basilisk Essential",
    0x0073: "Naga Trinity (C)",
    0x0078: "Viper",
    0x007A: "Viper Ultimate (Wired)",
    0x007B: "Viper Ultimate (Wireless)",
    0x007C: "DeathAdder V2 Pro (Wired)",
    0x007D: "DeathAdder V2 Pro (Wireless)",
    0x0083: "Lancehead Wireless (Receiver)",
    0x0084: "Basilisk X HyperSpeed",
    0x0085: "DeathAdder V2",
    0x0086: "DeathAdder V2 Mini",
    0x0088: "Naga Left-Handed 2020",
    0x008A: "Viper Mini SE (Wired)",
    0x008B: "Viper Mini SE (Wireless)",
    0x008C: "Basilisk V2",
    0x008D: "Basilisk Ultimate (Wired)",
    0x008F: "Basilisk Ultimate (Wireless)",
    0x0090: "Orochi V2 (Receiver)",
    0x0091: "Orochi V2 (Bluetooth)",
    0x0094: "DeathAdder V2 Lite",
    0x0096: "Naga Pro (Wired)",
    0x0098: "Naga Pro (Wireless)",
    0x009A: "Viper 8KHz",
    0x009C: "DeathAdder V2 X HyperSpeed",
    0x00A1: "Naga X",
    0x00A3: "Basilisk V3",
    0x00A5: "Mamba V2",
    0x00AA: "Cobra",
    0x00AD: "Deathadder Essential 2021",
    0x00AF: "Viper V2 Pro (Wired)",
    0x00B2: "DeathAdder V3",
    0x00B3: "Viper V2 Pro (Wireless)",
    0x00B6: "DeathAdder V3 Pro (Wired)",
    0x00B7: "DeathAdder V3 Pro (Wireless)",
    0x00B9: "DeathAdder V3 HyperSpeed",
    0x00BA: "Basilisk V3 Pro (Wired)",
    0x00BB: "Basilisk V3 Pro (Wireless)",
    0x00BC: "Pro Click Mini (Receiver)",
    0x00C0: "Cobra Pro (Wired)",
    0x00C1: "Cobra Pro (Wireless)",
    0x00C2: "DeathAdder V3 Pro (C) Wired",
    0x00C3: "DeathAdder V3 Pro (C) Wireless",
    0x00C4: "Naga V2 HyperSpeed (Receiver)",
    0x00C5: "Viper V3 HyperSpeed",
    0x00C6: "Basilisk V3 X HyperSpeed",
    0x00CD: "Viper V3 Pro (Wired)",
    0x00CE: "Viper V3 Pro (Wireless)",
}

# Known Logitech Gaming Mice (HID++ 2.0 compatible)
KNOWN_LOGITECH_MICE = {
    0xC24A: "G600",
    0xC246: "G700",
    0xC531: "G700s (Receiver)",
    0xC07C: "G700s (Wired)",
    0xC332: "G502 Hero",
    0xC08B: "G502 Hero (Wired)",
    0xC33C: "G502 Lightspeed (Receiver)",
    0xC08D: "G502 Lightspeed (Wired)",
    0xC539: "G502 X (Receiver)",
    0xC098: "G502 X (Wired)",
    0xC53A: "G502 X Plus (Receiver)",
    0xC099: "G502 X Plus (Wired)",
    0xC07D: "G402",
    0xC046: "G403 Hero (Wired)",
    0xC082: "G403 Hero",
    0xC083: "G403 Lightspeed",
    0xC08F: "G403 (Wired)",
    0xC084: "G203",
    0xC092: "G203 Lightsync",
    0xC334: "G Pro (Receiver)",
    0xC088: "G Pro (Wired)",
    0xC539: "G Pro X Superlight (Receiver)",
    0xC094: "G Pro X Superlight (Wired)",
    0xC53D: "G Pro X Superlight 2 (Receiver)",
    0xC09B: "G Pro X Superlight 2 (Wired)",
    0xC335: "G Pro Wireless (Receiver)",
    0xC08A: "G Pro Wireless (Wired)",
    0xC07E: "G102",
    0xC08E: "G102 Lightsync",
    0xC087: "G703 Hero (Wired)",
    0xC336: "G703 Hero (Receiver)",
    0xC337: "G903 Hero (Receiver)",
    0xC086: "G903 Hero (Wired)",
    0xC091: "G304/G305 (Receiver)",
    0xC085: "G304/G305 (Wired)",
    0xC547: "G309 (Receiver)",
    0xC093: "G309 (Wired)",
    0xC095: "G604 (Receiver)",
    0xC096: "MX Master 3",
    0xC548: "MX Master 3S (Receiver)",
    0xB034: "MX Master 3S (Bluetooth)",
    0xC52B: "Unifying Receiver",
    0xC534: "Nano Receiver",
    0xC545: "Lightspeed Receiver",
}

# EQ Presets (embedded - no external files needed)
EQ_PRESETS = {
    "fps": """# FPS Mode - Optimized for footsteps
Preamp: -3 dB
Filter 1: ON LP Fc 60 Hz Gain -6 dB
Filter 2: ON PK Fc 100 Hz Gain -4 dB Q 1.0
Filter 3: ON PK Fc 250 Hz Gain -2 dB Q 1.0
Filter 4: ON PK Fc 1000 Hz Gain 3 dB Q 1.5
Filter 5: ON PK Fc 2000 Hz Gain 4 dB Q 1.5
Filter 6: ON PK Fc 3500 Hz Gain 3 dB Q 1.5
Filter 7: ON PK Fc 6000 Hz Gain 2 dB Q 1.0
Filter 8: ON HP Fc 10000 Hz Gain -2 dB
""",
    "flat": """# Flat Mode - No EQ
Preamp: 0 dB
""",
    "music": """# Music Mode - Balanced
Preamp: -2 dB
Filter 1: ON PK Fc 60 Hz Gain 3 dB Q 1.0
Filter 2: ON PK Fc 150 Hz Gain 2 dB Q 1.0
Filter 3: ON PK Fc 3000 Hz Gain 1 dB Q 1.0
Filter 4: ON PK Fc 8000 Hz Gain 2 dB Q 1.0
Filter 5: ON PK Fc 12000 Hz Gain 2 dB Q 1.0
""",
}

# HID++ Constants
HIDPP_SHORT_MESSAGE = 0x10
HIDPP_LONG_MESSAGE = 0x11
HIDPP_DEVICE_RECEIVER = 0xFF
HIDPP_FEATURE_ROOT = 0x0000
HIDPP_FEATURE_ADJUSTABLE_DPI = 0x2201

# HID++ Function IDs
HIDPP_ROOT_GET_FEATURE = 0x00
HIDPP_DPI_GET_SENSOR_COUNT = 0x00
HIDPP_DPI_GET_SENSOR_DPI_LIST = 0x10
HIDPP_DPI_GET_SENSOR_DPI = 0x20
HIDPP_DPI_SET_SENSOR_DPI = 0x30


def calculate_razer_crc(data: bytes) -> int:
    """Calculate XOR checksum for Razer report (bytes 2-87)"""
    crc = 0
    for i in range(2, 88):
        crc ^= data[i]
    return crc


def build_razer_dpi_report(dpi_x: int, dpi_y: int, transaction_id: int = 0x1f) -> bytes:
    """Build a 90-byte Razer report to set DPI"""
    report = bytearray(90)
    report[0] = 0x00  # Status
    report[1] = transaction_id  # Transaction ID
    report[2] = 0x00  # Remaining packets high
    report[3] = 0x00  # Remaining packets low
    report[4] = 0x00  # Protocol type
    report[5] = 0x07  # Data size
    report[6] = 0x04  # Command class (MISC)
    report[7] = 0x05  # Command ID (SET_DPI)
    report[8] = 0x00  # Variable storage (NOSTORE)
    report[9] = (dpi_x >> 8) & 0xFF
    report[10] = dpi_x & 0xFF
    report[11] = (dpi_y >> 8) & 0xFF
    report[12] = dpi_y & 0xFF
    report[13] = 0x00
    report[14] = 0x00
    report[88] = calculate_razer_crc(report)
    report[89] = 0x00
    return bytes(report)


class LogitechHIDPP:
    """HID++ 2.0 protocol handler for Logitech mice"""

    def __init__(self, device_path):
        self.device = hid.device()
        self.device.open_path(device_path)
        self.device.set_nonblocking(False)
        self.dpi_feature_index = None

    def close(self):
        self.device.close()

    def send_short(self, device_idx, feature_idx, func_id, params=None):
        """Send a short HID++ message (7 bytes)"""
        if params is None:
            params = []
        msg = [HIDPP_SHORT_MESSAGE, device_idx, feature_idx, (func_id << 4) | 0x00]
        msg.extend(params[:3])
        while len(msg) < 7:
            msg.append(0x00)
        self.device.write(msg)
        return self._read_response()

    def send_long(self, device_idx, feature_idx, func_id, params=None):
        """Send a long HID++ message (20 bytes)"""
        if params is None:
            params = []
        msg = [HIDPP_LONG_MESSAGE, device_idx, feature_idx, (func_id << 4) | 0x00]
        msg.extend(params[:16])
        while len(msg) < 20:
            msg.append(0x00)
        self.device.write(msg)
        return self._read_response()

    def _read_response(self, timeout_ms=1000):
        """Read HID++ response"""
        data = self.device.read(64, timeout_ms)
        if not data:
            raise Exception("No response from device")
        # Check for error response (0x8F)
        if len(data) >= 4 and data[2] == 0x8F:
            error_code = data[4] if len(data) > 4 else 0
            raise Exception(f"HID++ error: 0x{error_code:02X}")
        return data

    def get_feature_index(self, feature_id):
        """Get the index of a feature from the root feature table"""
        params = [(feature_id >> 8) & 0xFF, feature_id & 0xFF]
        response = self.send_short(HIDPP_DEVICE_RECEIVER, 0x00, HIDPP_ROOT_GET_FEATURE, params)
        if len(response) >= 5:
            return response[4]
        return None

    def init_dpi_feature(self):
        """Initialize DPI feature index"""
        self.dpi_feature_index = self.get_feature_index(HIDPP_FEATURE_ADJUSTABLE_DPI)
        if self.dpi_feature_index is None or self.dpi_feature_index == 0:
            raise Exception("Device does not support DPI adjustment")
        return self.dpi_feature_index

    def set_dpi(self, dpi, sensor_idx=0):
        """Set DPI on the device"""
        if self.dpi_feature_index is None:
            self.init_dpi_feature()

        # DPI is big-endian 16-bit
        params = [sensor_idx, (dpi >> 8) & 0xFF, dpi & 0xFF]
        self.send_short(HIDPP_DEVICE_RECEIVER, self.dpi_feature_index,
                       HIDPP_DPI_SET_SENSOR_DPI >> 4, params)
        return True

    def get_dpi(self, sensor_idx=0):
        """Get current DPI from the device"""
        if self.dpi_feature_index is None:
            self.init_dpi_feature()

        params = [sensor_idx]
        response = self.send_short(HIDPP_DEVICE_RECEIVER, self.dpi_feature_index,
                                  HIDPP_DPI_GET_SENSOR_DPI >> 4, params)
        if len(response) >= 6:
            return (response[4] << 8) | response[5]
        return None


def is_equalizer_apo_installed():
    """Check if EqualizerAPO is installed"""
    return EQUALIZER_APO_PATH.exists()


def set_eq_preset(preset_name: str) -> bool:
    """Set EqualizerAPO preset"""
    if preset_name not in EQ_PRESETS:
        raise Exception(f"Unknown preset: {preset_name}")

    if not is_equalizer_apo_installed():
        raise Exception("EqualizerAPO not installed!\n\nPlease install from:\nhttps://sourceforge.net/projects/equalizerapo/")

    try:
        # Write the preset to config.txt
        with open(EQUALIZER_APO_CONFIG, 'w') as f:
            f.write(EQ_PRESETS[preset_name])
        return True
    except PermissionError:
        raise Exception("Cannot write to EqualizerAPO config.\nTry running as Administrator.")
    except Exception as e:
        raise Exception(f"Failed to set EQ: {e}")


def get_current_eq_preset() -> str:
    """Try to detect current EQ preset"""
    if not is_equalizer_apo_installed():
        return "unknown"

    try:
        with open(EQUALIZER_APO_CONFIG, 'r') as f:
            content = f.read()
            if "FPS Mode" in content:
                return "fps"
            elif "Music Mode" in content:
                return "music"
            elif "Flat Mode" in content or content.strip() == "" or "Preamp: 0" in content:
                return "flat"
    except:
        pass
    return "custom"


def find_razer_mouse():
    """Find connected Razer mouse"""
    for device in hid.enumerate(RAZER_VENDOR_ID):
        pid = device['product_id']
        iface = device.get('interface_number', -1)
        usage = device.get('usage', 0)

        # Interface 0 with mouse usage (0x0002) is the control interface
        if iface == 0 and usage == 0x0002:
            name = KNOWN_RAZER_MICE.get(pid, f"Razer Mouse (0x{pid:04X})")
            return device, name, "razer"

    # Fallback: any interface 0
    for device in hid.enumerate(RAZER_VENDOR_ID):
        pid = device['product_id']
        if device.get('interface_number', -1) == 0:
            name = KNOWN_RAZER_MICE.get(pid, f"Razer Mouse (0x{pid:04X})")
            return device, name, "razer"

    return None, None, None


def find_logitech_mouse():
    """Find connected Logitech mouse with HID++ support"""
    for device in hid.enumerate(LOGITECH_VENDOR_ID):
        pid = device['product_id']
        usage_page = device.get('usage_page', 0)

        # HID++ uses vendor-specific usage page (0xFF00) or sometimes 0x0001
        # Also check for known gaming mice
        if pid in KNOWN_LOGITECH_MICE:
            # Prefer the interface with usage_page 0xFF00 (vendor specific) for HID++
            if usage_page == 0xFF00 or usage_page == 0x0001:
                name = KNOWN_LOGITECH_MICE.get(pid, f"Logitech Mouse (0x{pid:04X})")
                return device, name, "logitech"

    # Fallback: try any Logitech device with vendor usage page
    for device in hid.enumerate(LOGITECH_VENDOR_ID):
        pid = device['product_id']
        usage_page = device.get('usage_page', 0)
        if usage_page == 0xFF00:
            name = KNOWN_LOGITECH_MICE.get(pid, f"Logitech Device (0x{pid:04X})")
            return device, name, "logitech"

    return None, None, None


def find_any_mouse():
    """Find any supported mouse (Razer or Logitech)"""
    # Try Razer first
    device, name, brand = find_razer_mouse()
    if device:
        return device, name, brand

    # Try Logitech
    device, name, brand = find_logitech_mouse()
    if device:
        return device, name, brand

    return None, None, None


def find_all_mice():
    """Find all supported mice"""
    mice = []

    # Find Razer mice
    for device in hid.enumerate(RAZER_VENDOR_ID):
        pid = device['product_id']
        iface = device.get('interface_number', -1)
        usage = device.get('usage', 0)

        if iface == 0 and usage == 0x0002:
            name = KNOWN_RAZER_MICE.get(pid, f"Razer Mouse (0x{pid:04X})")
            mice.append((device, name, "razer"))
            break  # Only add one Razer device

    if not any(m[2] == "razer" for m in mice):
        for device in hid.enumerate(RAZER_VENDOR_ID):
            pid = device['product_id']
            if device.get('interface_number', -1) == 0:
                name = KNOWN_RAZER_MICE.get(pid, f"Razer Mouse (0x{pid:04X})")
                mice.append((device, name, "razer"))
                break

    # Find Logitech mice
    found_logitech = set()
    for device in hid.enumerate(LOGITECH_VENDOR_ID):
        pid = device['product_id']
        usage_page = device.get('usage_page', 0)

        if pid in KNOWN_LOGITECH_MICE and pid not in found_logitech:
            if usage_page == 0xFF00 or usage_page == 0x0001:
                name = KNOWN_LOGITECH_MICE.get(pid, f"Logitech Mouse (0x{pid:04X})")
                mice.append((device, name, "logitech"))
                found_logitech.add(pid)

    return mice


def set_razer_dpi(device_info, dpi: int) -> bool:
    """Set DPI on Razer mouse"""
    try:
        device = hid.device()
        device.open_path(device_info['path'])
        report = build_razer_dpi_report(dpi, dpi)
        result = device.send_feature_report(b'\x00' + report)
        device.close()

        if result < 0:
            raise Exception("send_feature_report failed")
        return True
    except Exception as e:
        raise Exception(f"Failed to set DPI: {e}")


def set_logitech_dpi(device_info, dpi: int) -> bool:
    """Set DPI on Logitech mouse using HID++"""
    try:
        hidpp = LogitechHIDPP(device_info['path'])
        hidpp.set_dpi(dpi)
        hidpp.close()
        return True
    except Exception as e:
        raise Exception(f"Failed to set DPI: {e}")


def set_dpi(device_info, brand: str, dpi: int) -> bool:
    """Set DPI on any supported mouse"""
    if brand == "razer":
        return set_razer_dpi(device_info, dpi)
    elif brand == "logitech":
        return set_logitech_dpi(device_info, dpi)
    else:
        raise Exception(f"Unknown brand: {brand}")


class DPIApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Gaming Tool")
        self.root.geometry("380x480")
        self.root.resizable(False, False)

        self.current_device = None
        self.current_brand = None
        self.mice_list = []
        self.current_eq = "unknown"

        # Device selection
        device_frame = ttk.LabelFrame(self.root, text="Mouse", padding=10)
        device_frame.pack(fill='x', padx=20, pady=10)

        self.device_combo = ttk.Combobox(device_frame, state='readonly', width=32)
        self.device_combo.pack(side='left', padx=5)
        self.device_combo.bind('<<ComboboxSelected>>', self.on_device_select)

        refresh_btn = ttk.Button(device_frame, text="Refresh", command=self.refresh)
        refresh_btn.pack(side='left', padx=5)

        # Status label
        self.status_var = tk.StringVar()
        self.status_var.set("Click Refresh to scan")
        status_label = ttk.Label(self.root, textvariable=self.status_var)
        status_label.pack(pady=5)

        # Preset DPI buttons frame
        preset_frame = ttk.LabelFrame(self.root, text="Preset DPI", padding=10)
        preset_frame.pack(fill='x', padx=20, pady=5)

        presets = [400, 800, 1200, 1600, 3200]
        for dpi in presets:
            btn = ttk.Button(
                preset_frame,
                text=str(dpi),
                command=lambda d=dpi: self.apply_dpi(d),
                width=6
            )
            btn.pack(side='left', padx=2, expand=True)

        # Custom DPI frame
        custom_frame = ttk.LabelFrame(self.root, text="Custom DPI", padding=10)
        custom_frame.pack(fill='x', padx=20, pady=5)

        self.custom_entry = ttk.Entry(custom_frame, width=10)
        self.custom_entry.pack(side='left', padx=5)
        self.custom_entry.insert(0, "800")

        apply_btn = ttk.Button(
            custom_frame,
            text="Apply",
            command=self.apply_custom_dpi
        )
        apply_btn.pack(side='left', padx=5)

        # DPI Result label
        self.result_var = tk.StringVar()
        result_label = ttk.Label(self.root, textvariable=self.result_var)
        result_label.pack(pady=5)

        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', padx=20, pady=10)

        # Audio EQ Frame
        eq_frame = ttk.LabelFrame(self.root, text="Audio EQ (EqualizerAPO)", padding=10)
        eq_frame.pack(fill='x', padx=20, pady=5)

        # EQ Status
        self.eq_status_var = tk.StringVar()
        self.update_eq_status()
        eq_status_label = ttk.Label(eq_frame, textvariable=self.eq_status_var)
        eq_status_label.pack(pady=5)

        # EQ Buttons
        eq_btn_frame = ttk.Frame(eq_frame)
        eq_btn_frame.pack(fill='x', pady=5)

        self.fps_btn = ttk.Button(
            eq_btn_frame,
            text="FPS Mode",
            command=lambda: self.apply_eq("fps"),
            width=12
        )
        self.fps_btn.pack(side='left', padx=5, expand=True)

        self.music_btn = ttk.Button(
            eq_btn_frame,
            text="Music Mode",
            command=lambda: self.apply_eq("music"),
            width=12
        )
        self.music_btn.pack(side='left', padx=5, expand=True)

        self.flat_btn = ttk.Button(
            eq_btn_frame,
            text="Flat (Off)",
            command=lambda: self.apply_eq("flat"),
            width=12
        )
        self.flat_btn.pack(side='left', padx=5, expand=True)

        # EQ Result
        self.eq_result_var = tk.StringVar()
        eq_result_label = ttk.Label(eq_frame, textvariable=self.eq_result_var)
        eq_result_label.pack(pady=5)

        # Info
        info_label = ttk.Label(
            self.root,
            text="Razer/Logitech DPI + EqualizerAPO EQ",
            foreground="gray"
        )
        info_label.pack(pady=10)

        # Auto-refresh on start
        self.root.after(100, self.refresh)

    def update_eq_status(self):
        if not is_equalizer_apo_installed():
            self.eq_status_var.set("EqualizerAPO: Not installed")
            self.current_eq = "not_installed"
        else:
            self.current_eq = get_current_eq_preset()
            preset_names = {"fps": "FPS Mode", "music": "Music Mode", "flat": "Flat", "custom": "Custom", "unknown": "Unknown"}
            self.eq_status_var.set(f"Current: {preset_names.get(self.current_eq, self.current_eq)}")

    def refresh(self):
        self.mice_list = find_all_mice()
        self.device_combo['values'] = []

        if self.mice_list:
            names = [f"[{m[2].upper()}] {m[1]}" for m in self.mice_list]
            self.device_combo['values'] = names
            self.device_combo.current(0)
            self.on_device_select(None)
        else:
            self.current_device = None
            self.current_brand = None
            self.status_var.set("No mice found!")
            self.result_var.set("")

        self.update_eq_status()

    def on_device_select(self, event):
        idx = self.device_combo.current()
        if idx >= 0 and idx < len(self.mice_list):
            self.current_device = self.mice_list[idx][0]
            self.current_brand = self.mice_list[idx][2]
            name = self.mice_list[idx][1]
            self.status_var.set(f"Selected: {name}")
            self.result_var.set("")

    def apply_dpi(self, dpi: int):
        if not self.current_device:
            messagebox.showerror("Error", "No device selected!\nClick Refresh to scan.")
            return

        try:
            set_dpi(self.current_device, self.current_brand, dpi)
            self.result_var.set(f"DPI set to {dpi}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.result_var.set("Failed!")

    def apply_custom_dpi(self):
        try:
            dpi = int(self.custom_entry.get())
            if dpi < 100 or dpi > 25600:
                messagebox.showwarning("Warning", "DPI must be between 100 and 25600")
                return
            self.apply_dpi(dpi)
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a valid number")

    def apply_eq(self, preset: str):
        try:
            set_eq_preset(preset)
            preset_names = {"fps": "FPS Mode", "music": "Music Mode", "flat": "Flat"}
            self.eq_result_var.set(f"EQ: {preset_names.get(preset, preset)} applied!")
            self.update_eq_status()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.eq_result_var.set("Failed!")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = DPIApp()
    app.run()
