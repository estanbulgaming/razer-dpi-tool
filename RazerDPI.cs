using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;

class RazerDPI : Form
{
    // Razer USB IDs
    const ushort RAZER_VENDOR_ID = 0x1532;
    static readonly ushort[] DEATHADDER_V3_PIDS = { 0x00B6, 0x00B7, 0x00B8, 0x00B9 };

    // HID API
    [DllImport("hid.dll", SetLastError = true)]
    static extern void HidD_GetHidGuid(out Guid hidGuid);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern IntPtr SetupDiGetClassDevs(ref Guid classGuid, IntPtr enumerator, IntPtr hwndParent, uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiEnumDeviceInterfaces(IntPtr deviceInfoSet, IntPtr deviceInfoData, ref Guid interfaceClassGuid, uint memberIndex, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr deviceInfoSet, ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData, IntPtr deviceInterfaceDetailData, uint deviceInterfaceDetailDataSize, out uint requiredSize, IntPtr deviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern IntPtr CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool CloseHandle(IntPtr hObject);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_GetAttributes(IntPtr hidDeviceObject, ref HIDD_ATTRIBUTES attributes);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_SetFeature(IntPtr hidDeviceObject, byte[] reportBuffer, uint reportBufferLength);

    [StructLayout(LayoutKind.Sequential)]
    struct SP_DEVICE_INTERFACE_DATA
    {
        public int cbSize;
        public Guid InterfaceClassGuid;
        public int Flags;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct HIDD_ATTRIBUTES
    {
        public int Size;
        public ushort VendorID;
        public ushort ProductID;
        public ushort VersionNumber;
    }

    const uint DIGCF_PRESENT = 0x02;
    const uint DIGCF_DEVICEINTERFACE = 0x10;
    const uint GENERIC_READ = 0x80000000;
    const uint GENERIC_WRITE = 0x40000000;
    const uint FILE_SHARE_READ = 0x01;
    const uint FILE_SHARE_WRITE = 0x02;
    const uint OPEN_EXISTING = 3;
    static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    Label statusLabel;
    Label resultLabel;

    public RazerDPI()
    {
        Text = "Razer DPI Tool";
        Size = new Size(320, 300);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;

        // Status
        statusLabel = new Label
        {
            Text = FindMouse() ? "Mouse connected" : "Mouse not found!",
            Location = new Point(10, 10),
            Size = new Size(280, 20),
            TextAlign = ContentAlignment.MiddleCenter
        };
        Controls.Add(statusLabel);

        // Preset buttons
        var presetGroup = new GroupBox
        {
            Text = "Preset DPI",
            Location = new Point(10, 40),
            Size = new Size(285, 60)
        };

        int[] presets = { 400, 600, 800, 1600, 3200 };
        int btnX = 10;
        foreach (int dpi in presets)
        {
            var btn = new Button
            {
                Text = dpi.ToString(),
                Location = new Point(btnX, 20),
                Size = new Size(50, 28),
                Tag = dpi
            };
            btn.Click += PresetButton_Click;
            presetGroup.Controls.Add(btn);
            btnX += 54;
        }
        Controls.Add(presetGroup);

        // Custom DPI
        var customGroup = new GroupBox
        {
            Text = "Custom DPI",
            Location = new Point(10, 110),
            Size = new Size(285, 60)
        };

        var customInput = new TextBox
        {
            Name = "customInput",
            Text = "800",
            Location = new Point(10, 24),
            Size = new Size(80, 24)
        };
        customGroup.Controls.Add(customInput);

        var applyBtn = new Button
        {
            Text = "Apply",
            Location = new Point(100, 22),
            Size = new Size(70, 28)
        };
        applyBtn.Click += ApplyCustom_Click;
        customGroup.Controls.Add(applyBtn);
        Controls.Add(customGroup);

        // Result
        resultLabel = new Label
        {
            Location = new Point(10, 180),
            Size = new Size(280, 25),
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font(Font.FontFamily, 10, FontStyle.Bold)
        };
        Controls.Add(resultLabel);

        // Info
        var infoLabel = new Label
        {
            Text = "DPI range: 100 - 30000",
            Location = new Point(10, 210),
            Size = new Size(280, 20),
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.Gray
        };
        Controls.Add(infoLabel);
    }

    void PresetButton_Click(object sender, EventArgs e)
    {
        var btn = (Button)sender;
        int dpi = (int)btn.Tag;
        SetDPI(dpi);
    }

    void ApplyCustom_Click(object sender, EventArgs e)
    {
        var input = Controls.Find("customInput", true)[0] as TextBox;
        if (int.TryParse(input.Text, out int dpi))
        {
            if (dpi < 100 || dpi > 30000)
            {
                MessageBox.Show("DPI must be between 100 and 30000", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SetDPI(dpi);
        }
        else
        {
            MessageBox.Show("Please enter a valid number", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    bool FindMouse()
    {
        string path = FindRazerDevicePath();
        return path != null;
    }

    string FindRazerDevicePath()
    {
        Guid hidGuid;
        HidD_GetHidGuid(out hidGuid);

        IntPtr deviceInfoSet = SetupDiGetClassDevs(ref hidGuid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
        if (deviceInfoSet == INVALID_HANDLE_VALUE) return null;

        try
        {
            SP_DEVICE_INTERFACE_DATA deviceInterfaceData = new SP_DEVICE_INTERFACE_DATA();
            deviceInterfaceData.cbSize = Marshal.SizeOf(deviceInterfaceData);

            uint index = 0;
            while (SetupDiEnumDeviceInterfaces(deviceInfoSet, IntPtr.Zero, ref hidGuid, index, ref deviceInterfaceData))
            {
                uint requiredSize;
                SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref deviceInterfaceData, IntPtr.Zero, 0, out requiredSize, IntPtr.Zero);

                IntPtr detailDataBuffer = Marshal.AllocHGlobal((int)requiredSize);
                try
                {
                    Marshal.WriteInt32(detailDataBuffer, IntPtr.Size == 8 ? 8 : 6);

                    if (SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref deviceInterfaceData, detailDataBuffer, requiredSize, out requiredSize, IntPtr.Zero))
                    {
                        string devicePath = Marshal.PtrToStringAuto(detailDataBuffer + 4);

                        IntPtr handle = CreateFile(devicePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                        if (handle != INVALID_HANDLE_VALUE)
                        {
                            HIDD_ATTRIBUTES attrs = new HIDD_ATTRIBUTES();
                            attrs.Size = Marshal.SizeOf(attrs);

                            if (HidD_GetAttributes(handle, ref attrs))
                            {
                                if (attrs.VendorID == RAZER_VENDOR_ID && Array.IndexOf(DEATHADDER_V3_PIDS, attrs.ProductID) >= 0)
                                {
                                    CloseHandle(handle);
                                    return devicePath;
                                }
                            }
                            CloseHandle(handle);
                        }
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(detailDataBuffer);
                }
                index++;
            }
        }
        finally
        {
            SetupDiDestroyDeviceInfoList(deviceInfoSet);
        }
        return null;
    }

    byte CalculateCRC(byte[] data)
    {
        byte crc = 0;
        for (int i = 2; i < 88; i++)
            crc ^= data[i];
        return crc;
    }

    byte[] BuildDPIReport(int dpi)
    {
        byte[] report = new byte[91]; // 1 byte report ID + 90 bytes data

        report[0] = 0x00;  // Report ID
        report[1] = 0x00;  // Status
        report[2] = 0x1f;  // Transaction ID for DeathAdder V3
        report[3] = 0x00;  // Remaining packets high
        report[4] = 0x00;  // Remaining packets low
        report[5] = 0x00;  // Protocol type
        report[6] = 0x07;  // Data size
        report[7] = 0x04;  // Command class (misc)
        report[8] = 0x05;  // Command ID (set DPI)

        // Arguments
        report[9] = 0x00;  // Variable storage (NOSTORE)
        report[10] = (byte)((dpi >> 8) & 0xFF);  // DPI X high
        report[11] = (byte)(dpi & 0xFF);         // DPI X low
        report[12] = (byte)((dpi >> 8) & 0xFF);  // DPI Y high
        report[13] = (byte)(dpi & 0xFF);         // DPI Y low
        report[14] = 0x00;
        report[15] = 0x00;

        // CRC (calculated on bytes 2-88 of the 90-byte payload, which is indices 3-89 in our array)
        byte crc = 0;
        for (int i = 3; i < 89; i++)
            crc ^= report[i];
        report[89] = crc;
        report[90] = 0x00;  // Reserved

        return report;
    }

    void SetDPI(int dpi)
    {
        string devicePath = FindRazerDevicePath();
        if (devicePath == null)
        {
            MessageBox.Show("Razer DeathAdder V3 not found!\nMake sure mouse is connected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            resultLabel.Text = "Failed!";
            return;
        }

        IntPtr handle = CreateFile(devicePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
        if (handle == INVALID_HANDLE_VALUE)
        {
            MessageBox.Show("Cannot open device. Try running as Administrator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            resultLabel.Text = "Failed!";
            return;
        }

        try
        {
            byte[] report = BuildDPIReport(dpi);
            if (HidD_SetFeature(handle, report, (uint)report.Length))
            {
                resultLabel.Text = $"DPI set to {dpi}";
                resultLabel.ForeColor = Color.Green;
            }
            else
            {
                int error = Marshal.GetLastWin32Error();
                MessageBox.Show($"Failed to set DPI. Error code: {error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                resultLabel.Text = "Failed!";
                resultLabel.ForeColor = Color.Red;
            }
        }
        finally
        {
            CloseHandle(handle);
        }
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new RazerDPI());
    }
}
