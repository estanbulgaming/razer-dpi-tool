"""
Razer DeathAdder V3 DPI Tool
Simple tool to change DPI without Razer Synapse
"""

import hid
import struct
import tkinter as tk
from tkinter import ttk, messagebox

# Razer USB Vendor ID
RAZER_VENDOR_ID = 0x1532

# DeathAdder V3 variants
DEATHADDER_V3_PRODUCTS = {
    0x00B6: "DeathAdder V3",
    0x00B7: "DeathAdder V3 Pro (Wired)",
    0x00B8: "DeathAdder V3 Pro (Wireless)",
    0x00B9: "DeathAdder V3 HyperSpeed",
}

# Razer protocol constants
RAZER_CMD_SET_DPI = 0x05
RAZER_CMD_CLASS_MISC = 0x04
RAZER_VARSTORE_NOSTORE = 0x00
RAZER_REPORT_LEN = 90


def calculate_crc(data: bytes) -> int:
    """Calculate XOR checksum for Razer report (bytes 2-87)"""
    crc = 0
    for i in range(2, 88):
        crc ^= data[i]
    return crc


def build_dpi_report(dpi_x: int, dpi_y: int, transaction_id: int = 0x1f) -> bytes:
    """Build a 90-byte Razer report to set DPI"""
    report = bytearray(RAZER_REPORT_LEN)

    # Status byte
    report[0] = 0x00
    # Transaction ID (0x1f for DeathAdder V3)
    report[1] = transaction_id
    # Remaining packets (big-endian short)
    report[2] = 0x00
    report[3] = 0x00
    # Protocol type
    report[4] = 0x00
    # Data size
    report[5] = 0x07
    # Command class
    report[6] = RAZER_CMD_CLASS_MISC
    # Command ID
    report[7] = RAZER_CMD_SET_DPI

    # Arguments (DPI data)
    report[8] = RAZER_VARSTORE_NOSTORE  # Variable storage
    report[9] = (dpi_x >> 8) & 0xFF     # DPI X high byte
    report[10] = dpi_x & 0xFF           # DPI X low byte
    report[11] = (dpi_y >> 8) & 0xFF    # DPI Y high byte
    report[12] = dpi_y & 0xFF           # DPI Y low byte
    report[13] = 0x00
    report[14] = 0x00

    # Calculate and set CRC
    report[88] = calculate_crc(report)
    report[89] = 0x00  # Reserved

    return bytes(report)


def find_razer_mouse():
    """Find connected Razer DeathAdder V3"""
    for device in hid.enumerate(RAZER_VENDOR_ID):
        pid = device['product_id']
        if pid in DEATHADDER_V3_PRODUCTS:
            # Prefer interface 0 for control
            if device.get('interface_number', 0) == 0:
                return device, DEATHADDER_V3_PRODUCTS[pid]
    return None, None


def set_dpi(dpi: int) -> bool:
    """Set DPI on connected Razer mouse"""
    device_info, name = find_razer_mouse()

    if not device_info:
        raise Exception("Razer DeathAdder V3 not found!\nMake sure mouse is connected.")

    try:
        device = hid.device()
        device.open_path(device_info['path'])

        report = build_dpi_report(dpi, dpi)

        # Send as feature report (report ID 0x00)
        # Prepend report ID for Windows
        device.send_feature_report(b'\x00' + report)

        device.close()
        return True

    except Exception as e:
        raise Exception(f"Failed to set DPI: {e}")


class DPIApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Razer DPI Tool")
        self.root.geometry("300x280")
        self.root.resizable(False, False)

        # Check mouse on startup
        device_info, name = find_razer_mouse()

        # Status label
        self.status_var = tk.StringVar()
        if name:
            self.status_var.set(f"Connected: {name}")
        else:
            self.status_var.set("Mouse not found!")

        status_label = ttk.Label(self.root, textvariable=self.status_var)
        status_label.pack(pady=10)

        # Preset buttons frame
        preset_frame = ttk.LabelFrame(self.root, text="Preset DPI", padding=10)
        preset_frame.pack(fill='x', padx=20, pady=5)

        presets = [400, 600, 800, 1600, 3200]
        for dpi in presets:
            btn = ttk.Button(
                preset_frame,
                text=str(dpi),
                command=lambda d=dpi: self.apply_dpi(d),
                width=8
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
