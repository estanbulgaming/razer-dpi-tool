"""Test script for Logitech HID enumeration"""
import hid

print("=== All Logitech HID devices ===\n")

LOGITECH_VENDOR_ID = 0x046D

devices = hid.enumerate(LOGITECH_VENDOR_ID)
print(f"Found {len(devices)} Logitech HID interfaces:\n")

for i, d in enumerate(devices):
    print(f"[{i}] PID=0x{d['product_id']:04X}")
    print(f"    usage_page=0x{d['usage_page']:04X}, usage=0x{d['usage']:04X}")
    print(f"    interface={d['interface_number']}")
    print(f"    product: {d['product_string']}")
    print(f"    path: {d['path']}")
    print()

# Try to find gaming mice
KNOWN_MICE = {
    0xC332: "G502 Hero",
    0xC08B: "G502 Hero (Wired)",
    0xC33C: "G502 Lightspeed",
}

print("\n=== Known gaming mice ===\n")
for d in devices:
    pid = d['product_id']
    if pid in KNOWN_MICE:
        print(f"Found: {KNOWN_MICE[pid]} (0x{pid:04X})")
        print(f"  usage_page=0x{d['usage_page']:04X}, usage=0x{d['usage']:04X}, iface={d['interface_number']}")

input("\nPress Enter to exit...")
