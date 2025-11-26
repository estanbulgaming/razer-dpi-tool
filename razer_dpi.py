"""
Razer DPI Tool
Simple tool to change DPI without Razer Synapse
"""

import hid
import tkinter as tk
from tkinter import ttk, messagebox

# Razer USB Vendor ID
RAZER_VENDOR_ID = 0x1532

# Known Razer Mouse Product IDs (from OpenRazer)
KNOWN_MICE = {
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


def calculate_crc(data: bytes) -> int:
    """Calculate XOR checksum for Razer report (bytes 2-87)"""
    crc = 0
    for i in range(2, 88):
        crc ^= data[i]
    return crc


def build_dpi_report(dpi_x: int, dpi_y: int, transaction_id: int = 0x1f) -> bytes:
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
    report[88] = calculate_crc(report)
    report[89] = 0x00
    return bytes(report)


def find_razer_mouse():
    """Find connected Razer mouse"""
    for device in hid.enumerate(RAZER_VENDOR_ID):
        pid = device['product_id']
        iface = device.get('interface_number', -1)
        usage = device.get('usage', 0)

        # Interface 0 with mouse usage (0x0002) is the control interface
        if iface == 0 and usage == 0x0002:
            name = KNOWN_MICE.get(pid, f"Razer Mouse (0x{pid:04X})")
            return device, name

    # Fallback: any interface 0
    for device in hid.enumerate(RAZER_VENDOR_ID):
        pid = device['product_id']
        if device.get('interface_number', -1) == 0:
            name = KNOWN_MICE.get(pid, f"Razer Mouse (0x{pid:04X})")
            return device, name

    return None, None


def set_dpi(dpi: int) -> bool:
    """Set DPI on connected Razer mouse"""
    device_info, name = find_razer_mouse()

    if not device_info:
        raise Exception("Razer mouse not found!\nMake sure mouse is connected.")

    try:
        device = hid.device()
        device.open_path(device_info['path'])
        report = build_dpi_report(dpi, dpi)
        # Prepend report ID 0x00
        result = device.send_feature_report(b'\x00' + report)
        device.close()

        if result < 0:
            raise Exception("send_feature_report failed")
        return True

    except Exception as e:
        raise Exception(f"Failed to set DPI: {e}")


class DPIApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Razer DPI Tool")
        self.root.geometry("320x300")
        self.root.resizable(False, False)

        # Check mouse on startup
        device_info, name = find_razer_mouse()

        # Status label
        self.status_var = tk.StringVar()
        if name:
            self.status_var.set(f"Connected: {name}")
        else:
            self.status_var.set("Mouse not found! Click Refresh")

        status_label = ttk.Label(self.root, textvariable=self.status_var)
        status_label.pack(pady=10)

        # Refresh button
        refresh_btn = ttk.Button(self.root, text="Refresh", command=self.refresh)
        refresh_btn.pack(pady=5)

        # Preset buttons frame
        preset_frame = ttk.LabelFrame(self.root, text="Preset DPI", padding=10)
        preset_frame.pack(fill='x', padx=20, pady=5)

        presets = [400, 600, 800, 1600, 3200]
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
        custom_frame.pack(fill='x', padx=20, pady=10)

        self.custom_entry = ttk.Entry(custom_frame, width=10)
        self.custom_entry.pack(side='left', padx=5)
        self.custom_entry.insert(0, "800")

        apply_btn = ttk.Button(
            custom_frame,
            text="Apply",
            command=self.apply_custom_dpi
        )
        apply_btn.pack(side='left', padx=5)

        # Result label
        self.result_var = tk.StringVar()
        result_label = ttk.Label(self.root, textvariable=self.result_var)
        result_label.pack(pady=10)

        # Info
        info_label = ttk.Label(
            self.root,
            text="DPI range: 100 - 30000",
            foreground="gray"
        )
        info_label.pack(pady=5)

    def refresh(self):
        device_info, name = find_razer_mouse()
        if name:
            self.status_var.set(f"Connected: {name}")
            self.result_var.set("")
        else:
            self.status_var.set("Mouse not found!")
            self.result_var.set("")

    def apply_dpi(self, dpi: int):
        try:
            set_dpi(dpi)
            self.result_var.set(f"DPI set to {dpi}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.result_var.set("Failed!")

    def apply_custom_dpi(self):
        try:
            dpi = int(self.custom_entry.get())
            if dpi < 100 or dpi > 30000:
                messagebox.showwarning("Warning", "DPI must be between 100 and 30000")
                return
            self.apply_dpi(dpi)
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a valid number")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = DPIApp()
    app.run()
