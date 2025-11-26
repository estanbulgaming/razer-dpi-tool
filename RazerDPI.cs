using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Text;

class RazerDevice
{
    public string Path;
    public ushort ProductID;
    public string Name;
    public bool IsKnownMouse;
    public ushort FeatureReportLength;
    public ushort OutputReportLength;
    public ushort Usage;
    public ushort UsagePage;
    public bool ReadOnly;
}

class RazerDPI : Form
{
    const ushort RAZER_VENDOR_ID = 0x1532;

    // Known Razer Mouse Product IDs (from OpenRazer)
    static readonly Dictionary<ushort, string> KNOWN_MICE = new Dictionary<ushort, string>
    {
        { 0x0013, "Orochi 2011" },
        { 0x0015, "Naga" },
        { 0x0016, "DeathAdder 3.5G" },
        { 0x001F, "Naga Epic" },
        { 0x0020, "Abyssus 1800" },
        { 0x0024, "Mamba 2012 Wired" },
        { 0x0025, "Mamba 2012 Wireless" },
        { 0x0029, "DeathAdder 3.5G Black" },
        { 0x002E, "Naga 2012" },
        { 0x002F, "Imperator" },
        { 0x0032, "Ouroboros" },
        { 0x0034, "Taipan" },
        { 0x0036, "Naga Hex Red" },
        { 0x0037, "DeathAdder 2013" },
        { 0x0038, "DeathAdder 1800" },
        { 0x0039, "Orochi 2013" },
        { 0x003E, "Naga Epic Chroma" },
        { 0x003F, "Naga Epic Chroma Dock" },
        { 0x0040, "Naga 2014" },
        { 0x0041, "Naga Hex" },
        { 0x0042, "Abyssus" },
        { 0x0043, "DeathAdder Chroma" },
        { 0x0044, "Mamba Wired" },
        { 0x0045, "Mamba Wireless" },
        { 0x0046, "Mamba TE Wired" },
        { 0x0048, "Orochi Chroma" },
        { 0x004C, "Diamondback Chroma" },
        { 0x004F, "DeathAdder 2000" },
        { 0x0050, "Naga Hex V2" },
        { 0x0053, "Naga Chroma" },
        { 0x0054, "DeathAdder 3500" },
        { 0x0059, "Lancehead Wired" },
        { 0x005A, "Lancehead Wireless" },
        { 0x005B, "Abyssus V2" },
        { 0x005C, "DeathAdder Elite" },
        { 0x005E, "Abyssus 2000" },
        { 0x0060, "Lancehead TE Wired" },
        { 0x0062, "Atheris Receiver" },
        { 0x0064, "Basilisk" },
        { 0x0065, "Basilisk Essential" },
        { 0x0067, "Naga Trinity" },
        { 0x006A, "Abyssus Elite DVa Edition" },
        { 0x006B, "Abyssus Essential" },
        { 0x006C, "Mamba Elite" },
        { 0x006E, "DeathAdder Essential" },
        { 0x006F, "Lancehead Wireless Receiver" },
        { 0x0070, "Lancehead Wireless Wired" },
        { 0x0071, "DeathAdder Essential White Edition" },
        { 0x0072, "Mamba Wireless Receiver" },
        { 0x0073, "Mamba Wireless Wired" },
        { 0x0077, "Pro Click Receiver" },
        { 0x0078, "Viper" },
        { 0x007A, "Viper Ultimate Wired" },
        { 0x007B, "Viper Ultimate Wireless" },
        { 0x007C, "DeathAdder V2 Pro Wired" },
        { 0x007D, "DeathAdder V2 Pro Wireless" },
        { 0x0080, "Pro Click Wired" },
        { 0x0083, "Basilisk X HyperSpeed" },
        { 0x0084, "DeathAdder V2" },
        { 0x0085, "Basilisk V2" },
        { 0x0086, "Basilisk Ultimate Wired" },
        { 0x0088, "Basilisk Ultimate Receiver" },
        { 0x008A, "Viper Mini" },
        { 0x008C, "DeathAdder V2 Mini" },
        { 0x008D, "Naga Left Handed 2020" },
        { 0x008F, "Naga Pro Wired" },
        { 0x0090, "Naga Pro Wireless" },
        { 0x0091, "Viper 8K" },
        { 0x0094, "Orochi V2 Receiver" },
        { 0x0095, "Orochi V2 Bluetooth" },
        { 0x0096, "Naga X" },
        { 0x0098, "DeathAdder Essential 2021" },
        { 0x0099, "Basilisk V3" },
        { 0x009A, "Pro Click Mini Receiver" },
        { 0x009C, "DeathAdder V2 X HyperSpeed" },
        { 0x009E, "Viper Mini SE Wired" },
        { 0x009F, "Viper Mini SE Wireless" },
        { 0x00A1, "DeathAdder V2 Lite" },
        { 0x00A3, "Cobra" },
        { 0x00A5, "Viper V2 Pro Wired" },
        { 0x00A6, "Viper V2 Pro Wireless" },
        { 0x00A7, "Naga V2 Pro Wired" },
        { 0x00A8, "Naga V2 Pro Wireless" },
        { 0x00AA, "Basilisk V3 Pro Wired" },
        { 0x00AB, "Basilisk V3 Pro Wireless" },
        { 0x00AF, "Cobra Pro Wired" },
        { 0x00B0, "Cobra Pro Wireless" },
        { 0x00B2, "DeathAdder V3" },
        { 0x00B3, "HyperPolling Wireless Dongle" },
        { 0x00B4, "Naga V2 HyperSpeed Receiver" },
        { 0x00B6, "DeathAdder V3 Pro Wired" },
        { 0x00B7, "DeathAdder V3 Pro Wireless" },
        { 0x00B8, "Viper V3 HyperSpeed" },
        { 0x00B9, "Basilisk V3 X HyperSpeed" },
        { 0x00BE, "DeathAdder V4 Pro Wired" },
        { 0x00BF, "DeathAdder V4 Pro Wireless" },
        { 0x00C0, "Viper V3 Pro Wired" },
        { 0x00C1, "Viper V3 Pro Wireless" },
        { 0x00C2, "DeathAdder V3 Pro Wired Alt" },
        { 0x00C3, "DeathAdder V3 Pro Wireless Alt" },
        { 0x00C4, "DeathAdder V3 HyperSpeed Wired" },
        { 0x00C5, "DeathAdder V3 HyperSpeed Wireless" },
        { 0x00C7, "Pro Click V2 Vertical Edition Wired" },
        { 0x00C8, "Pro Click V2 Vertical Edition Wireless" },
        { 0x00CB, "Basilisk V3 35K" },
        { 0x00CC, "Basilisk V3 Pro 35K Wired" },
        { 0x00CD, "Basilisk V3 Pro 35K Wireless" },
        { 0x00D0, "Pro Click V2 Wired" },
        { 0x00D1, "Pro Click V2 Wireless" },
        { 0x00D6, "Basilisk V3 Pro 35K Phantom Green Edition Wired" },
        { 0x00D7, "Basilisk V3 Pro 35K Phantom Green Edition Wireless" },
    };

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

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_SetOutputReport(IntPtr hidDeviceObject, byte[] reportBuffer, uint reportBufferLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool WriteFile(IntPtr hFile, byte[] lpBuffer, uint nNumberOfBytesToWrite, out uint lpNumberOfBytesWritten, IntPtr lpOverlapped);

    [DllImport("hid.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern bool HidD_GetProductString(IntPtr hidDeviceObject, byte[] buffer, uint bufferLength);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_GetPreparsedData(IntPtr hidDeviceObject, out IntPtr preparsedData);

    [DllImport("hid.dll", SetLastError = true)]
    static extern bool HidD_FreePreparsedData(IntPtr preparsedData);

    [DllImport("hid.dll", SetLastError = true)]
    static extern int HidP_GetCaps(IntPtr preparsedData, out HIDP_CAPS capabilities);

    [StructLayout(LayoutKind.Sequential)]
    struct HIDP_CAPS
    {
        public ushort Usage;
        public ushort UsagePage;
        public ushort InputReportByteLength;
        public ushort OutputReportByteLength;
        public ushort FeatureReportByteLength;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
        public ushort[] Reserved;
        public ushort NumberLinkCollectionNodes;
        public ushort NumberInputButtonCaps;
        public ushort NumberInputValueCaps;
        public ushort NumberInputDataIndices;
        public ushort NumberOutputButtonCaps;
        public ushort NumberOutputValueCaps;
        public ushort NumberOutputDataIndices;
        public ushort NumberFeatureButtonCaps;
        public ushort NumberFeatureValueCaps;
        public ushort NumberFeatureDataIndices;
    }

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
    ComboBox deviceCombo;
    List<RazerDevice> foundDevices;
    RazerDevice selectedDevice;
    StringBuilder log;
    string logPath;

    void Log(string message)
    {
        string line = DateTime.Now.ToString("HH:mm:ss") + " " + message;
        log.AppendLine(line);
    }

    void SaveLog()
    {
        try
        {
            File.WriteAllText(logPath, log.ToString());
        }
        catch { }
    }

    public RazerDPI()
    {
        log = new StringBuilder();
        logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "razer_dpi_log.txt");
        Log("=== Razer DPI Tool Started ===");
        Log("OS: " + Environment.OSVersion.ToString());
        Log("64-bit: " + Environment.Is64BitOperatingSystem.ToString());

        Text = "Razer DPI Tool";
        Size = new Size(340, 360);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;

        // Status
        statusLabel = new Label
        {
            Text = "Searching for Razer devices...",
            Location = new Point(10, 10),
            Size = new Size(300, 20),
            TextAlign = ContentAlignment.MiddleCenter
        };
        Controls.Add(statusLabel);

        // Device dropdown
        deviceCombo = new ComboBox
        {
            Location = new Point(10, 35),
            Size = new Size(230, 24),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        deviceCombo.SelectedIndexChanged += DeviceCombo_Changed;
        Controls.Add(deviceCombo);

        // Refresh button
        var refreshBtn = new Button
        {
            Text = "Refresh",
            Location = new Point(250, 34),
            Size = new Size(65, 24)
        };
        refreshBtn.Click += RefreshButton_Click;
        Controls.Add(refreshBtn);

        // Preset buttons
        var presetGroup = new GroupBox
        {
            Text = "Preset DPI",
            Location = new Point(10, 70),
            Size = new Size(305, 60)
        };

        int[] presets = { 400, 600, 800, 1600, 3200 };
        int btnX = 10;
        foreach (int dpi in presets)
        {
            var btn = new Button
            {
                Text = dpi.ToString(),
                Location = new Point(btnX, 22),
                Size = new Size(54, 28),
                Tag = dpi
            };
            btn.Click += PresetButton_Click;
            presetGroup.Controls.Add(btn);
            btnX += 58;
        }
        Controls.Add(presetGroup);

        // Custom DPI
        var customGroup = new GroupBox
        {
            Text = "Custom DPI",
            Location = new Point(10, 140),
            Size = new Size(305, 60)
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
            Location = new Point(10, 210),
            Size = new Size(305, 25),
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font(Font.FontFamily, 10, FontStyle.Bold)
        };
        Controls.Add(resultLabel);

        // Info
        var infoLabel = new Label
        {
            Text = "DPI range: 100 - 30000\nGreen = Known mouse, Yellow = Unknown Razer device",
            Location = new Point(10, 245),
            Size = new Size(305, 40),
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.Gray
        };
        Controls.Add(infoLabel);

        // Initial scan
        RefreshDevices();
    }

    void DeviceCombo_Changed(object sender, EventArgs e)
    {
        if (deviceCombo.SelectedIndex >= 0 && deviceCombo.SelectedIndex < foundDevices.Count)
        {
            selectedDevice = foundDevices[deviceCombo.SelectedIndex];
            resultLabel.Text = "";
        }
    }

    void RefreshButton_Click(object sender, EventArgs e)
    {
        RefreshDevices();
    }

    void RefreshDevices()
    {
        foundDevices = FindAllRazerDevices();
        deviceCombo.Items.Clear();
        selectedDevice = null;

        if (foundDevices.Count == 0)
        {
            statusLabel.Text = "No Razer devices found!";
            return;
        }

        // Prioritize known mice
        foundDevices.Sort((a, b) => b.IsKnownMouse.CompareTo(a.IsKnownMouse));

        foreach (var device in foundDevices)
        {
            string prefix = device.IsKnownMouse ? "[Mouse] " : "[?] ";
            deviceCombo.Items.Add(prefix + device.Name);
        }

        deviceCombo.SelectedIndex = 0;
        selectedDevice = foundDevices[0];

        int knownCount = 0;
        foreach (var d in foundDevices) if (d.IsKnownMouse) knownCount++;

        statusLabel.Text = "Found " + foundDevices.Count.ToString() + " device(s), " + knownCount.ToString() + " known mouse(s)";
    }

    void PresetButton_Click(object sender, EventArgs e)
    {
        var btn = (Button)sender;
        int dpi = (int)btn.Tag;
        SetDPI(dpi);
    }

    void ApplyCustom_Click(object sender, EventArgs e)
    {
        TextBox input = (TextBox)Controls.Find("customInput", true)[0];
        int dpi;
        if (int.TryParse(input.Text, out dpi))
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

    List<RazerDevice> FindAllRazerDevices()
    {
        Log("--- Scanning for devices ---");
        List<RazerDevice> devices = new List<RazerDevice>();
        Dictionary<string, int> productIdCount = new Dictionary<string, int>();

        Guid hidGuid;
        HidD_GetHidGuid(out hidGuid);
        Log("HID GUID: " + hidGuid.ToString());

        IntPtr deviceInfoSet = SetupDiGetClassDevs(ref hidGuid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
        if (deviceInfoSet == INVALID_HANDLE_VALUE)
        {
            Log("ERROR: SetupDiGetClassDevs failed");
            return devices;
        }

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

                        // Try read/write first, then read-only if that fails
                        IntPtr handle = CreateFile(devicePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                        bool readOnly = false;
                        if (handle == INVALID_HANDLE_VALUE)
                        {
                            // Try read-only for enumeration
                            handle = CreateFile(devicePath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                            readOnly = true;
                        }
                        if (handle != INVALID_HANDLE_VALUE)
                        {
                            HIDD_ATTRIBUTES attrs = new HIDD_ATTRIBUTES();
                            attrs.Size = Marshal.SizeOf(attrs);

                            if (HidD_GetAttributes(handle, ref attrs))
                            {
                                if (attrs.VendorID == RAZER_VENDOR_ID)
                                {
                                    string productName;
                                    bool isKnown = KNOWN_MICE.TryGetValue(attrs.ProductID, out productName);

                                    if (!isKnown)
                                    {
                                        byte[] productBuffer = new byte[256];
                                        if (HidD_GetProductString(handle, productBuffer, (uint)productBuffer.Length))
                                        {
                                            productName = System.Text.Encoding.Unicode.GetString(productBuffer).TrimEnd('\0');
                                        }
                                        else
                                        {
                                            productName = "Unknown Device";
                                        }
                                    }

                                    // Track interface number for this product
                                    string pidKey = attrs.ProductID.ToString("X4");
                                    int ifaceNum = 0;
                                    if (productIdCount.ContainsKey(pidKey))
                                    {
                                        ifaceNum = productIdCount[pidKey];
                                        productIdCount[pidKey] = ifaceNum + 1;
                                    }
                                    else
                                    {
                                        productIdCount[pidKey] = 1;
                                    }

                                    // Extract MI (interface) number from path
                                    string miInfo = "";
                                    int miIdx = devicePath.IndexOf("&mi_");
                                    if (miIdx >= 0 && miIdx + 7 < devicePath.Length)
                                    {
                                        miInfo = " MI" + devicePath.Substring(miIdx + 4, 2);
                                    }

                                    // Extract collection number from path
                                    string colInfo = "";
                                    int colIdx = devicePath.IndexOf("&col");
                                    if (colIdx >= 0 && colIdx + 6 < devicePath.Length)
                                    {
                                        colInfo = " Col" + devicePath.Substring(colIdx + 4, 2);
                                    }

                                    // Get HID capabilities
                                    ushort featureLen = 0, outputLen = 0, usage = 0, usagePage = 0;
                                    IntPtr preparsedData;
                                    if (HidD_GetPreparsedData(handle, out preparsedData))
                                    {
                                        HIDP_CAPS caps;
                                        if (HidP_GetCaps(preparsedData, out caps) == 0x00110000)
                                        {
                                            featureLen = caps.FeatureReportByteLength;
                                            outputLen = caps.OutputReportByteLength;
                                            usage = caps.Usage;
                                            usagePage = caps.UsagePage;
                                        }
                                        HidD_FreePreparsedData(preparsedData);
                                    }

                                    string capsInfo = " Feat=" + featureLen.ToString() + " Out=" + outputLen.ToString() + " Usage=" + usagePage.ToString("X2") + ":" + usage.ToString("X2");

                                    Log("Found Razer device: VID=0x" + attrs.VendorID.ToString("X4") + " PID=0x" + pidKey + miInfo + colInfo);
                                    Log("  Path: " + devicePath);
                                    Log("  Name: " + productName);
                                    Log("  IsKnownMouse: " + isKnown.ToString() + ", ReadOnly: " + readOnly.ToString());
                                    Log("  FeatureReportLen: " + featureLen.ToString() + ", OutputReportLen: " + outputLen.ToString());
                                    Log("  UsagePage: 0x" + usagePage.ToString("X4") + ", Usage: 0x" + usage.ToString("X4"));

                                    devices.Add(new RazerDevice
                                    {
                                        Path = devicePath,
                                        ProductID = attrs.ProductID,
                                        Name = productName + " (0x" + pidKey + miInfo + colInfo + capsInfo + (readOnly ? " RO" : "") + ")",
                                        IsKnownMouse = isKnown,
                                        FeatureReportLength = featureLen,
                                        OutputReportLength = outputLen,
                                        Usage = usage,
                                        UsagePage = usagePage,
                                        ReadOnly = readOnly
                                    });
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
        Log("Total Razer devices found: " + devices.Count.ToString());

        // Sort: prioritize mi_00 (control interface) for known mice
        devices.Sort((a, b) => {
            // Known mice first
            if (a.IsKnownMouse != b.IsKnownMouse)
                return b.IsKnownMouse.CompareTo(a.IsKnownMouse);
            // Then by MI number (mi_00 first)
            bool aHasMi00 = a.Path.Contains("&mi_00");
            bool bHasMi00 = b.Path.Contains("&mi_00");
            if (aHasMi00 != bHasMi00)
                return bHasMi00.CompareTo(aHasMi00);
            return 0;
        });

        SaveLog();
        return devices;
    }

    byte[] BuildDPIReport(int dpi)
    {
        byte[] report = new byte[91];

        byte transactionId = 0x1f;

        report[0] = 0x00;
        report[1] = 0x00;
        report[2] = transactionId;
        report[3] = 0x00;
        report[4] = 0x00;
        report[5] = 0x00;
        report[6] = 0x07;
        report[7] = 0x04;
        report[8] = 0x05;

        report[9] = 0x00;
        report[10] = (byte)((dpi >> 8) & 0xFF);
        report[11] = (byte)(dpi & 0xFF);
        report[12] = (byte)((dpi >> 8) & 0xFF);
        report[13] = (byte)(dpi & 0xFF);
        report[14] = 0x00;
        report[15] = 0x00;

        byte crc = 0;
        for (int i = 3; i < 89; i++)
            crc ^= report[i];
        report[89] = crc;
        report[90] = 0x00;

        return report;
    }

    void SetDPI(int dpi)
    {
        Log("--- SetDPI called: " + dpi.ToString() + " ---");

        if (selectedDevice == null)
        {
            Log("ERROR: No device selected");
            MessageBox.Show("No device selected!\nClick Refresh to scan.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            resultLabel.Text = "Failed!";
            resultLabel.ForeColor = Color.Red;
            SaveLog();
            return;
        }

        Log("Selected device: " + selectedDevice.Name);
        Log("Device path: " + selectedDevice.Path);
        Log("Device ReadOnly: " + selectedDevice.ReadOnly.ToString());

        if (selectedDevice.ReadOnly)
        {
            Log("WARNING: Device was opened read-only during enumeration");
        }

        IntPtr handle = CreateFile(selectedDevice.Path, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
        if (handle == INVALID_HANDLE_VALUE)
        {
            int openError = Marshal.GetLastWin32Error();
            Log("ERROR: CreateFile failed with error: " + openError.ToString());
            MessageBox.Show("Cannot open device. Try running as Administrator.\nError: " + openError.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            resultLabel.Text = "Failed!";
            resultLabel.ForeColor = Color.Red;
            SaveLog();
            return;
        }
        Log("CreateFile succeeded, handle: " + handle.ToString());

        try
        {
            Log("Device FeatureReportLength: " + selectedDevice.FeatureReportLength.ToString());
            Log("Device OutputReportLength: " + selectedDevice.OutputReportLength.ToString());
            Log("Device UsagePage: 0x" + selectedDevice.UsagePage.ToString("X4") + ", Usage: 0x" + selectedDevice.Usage.ToString("X4"));

            byte[] report = BuildDPIReport(dpi);
            Log("Report built, length: " + report.Length.ToString());
            Log("Report hex: " + BitConverter.ToString(report, 0, 20));

            // Try multiple methods to send the report
            bool result = false;
            int error = 0;

            // Method 1: HidD_SetFeature with device's feature report length
            if (selectedDevice.FeatureReportLength > 0)
            {
                Log("Trying HidD_SetFeature with device length " + selectedDevice.FeatureReportLength.ToString() + "...");
                byte[] sizedReport = new byte[selectedDevice.FeatureReportLength];
                Array.Copy(report, 0, sizedReport, 0, Math.Min(report.Length, sizedReport.Length));
                result = HidD_SetFeature(handle, sizedReport, (uint)sizedReport.Length);
                error = Marshal.GetLastWin32Error();
                Log("HidD_SetFeature (sized) result: " + result.ToString() + ", error: " + error.ToString());
            }

            if (!result)
            {
                // Method 2: HidD_SetFeature with standard 91 bytes
                Log("Trying HidD_SetFeature with 91 bytes...");
                result = HidD_SetFeature(handle, report, (uint)report.Length);
                error = Marshal.GetLastWin32Error();
                Log("HidD_SetFeature result: " + result.ToString() + ", error: " + error.ToString());
            }

            if (!result)
            {
                // Method 3: HidD_SetOutputReport
                Log("Trying HidD_SetOutputReport...");
                result = HidD_SetOutputReport(handle, report, (uint)report.Length);
                error = Marshal.GetLastWin32Error();
                Log("HidD_SetOutputReport result: " + result.ToString() + ", error: " + error.ToString());
            }

            if (!result && selectedDevice.OutputReportLength > 0)
            {
                // Method 4: WriteFile with device's output report length
                Log("Trying WriteFile with device length " + selectedDevice.OutputReportLength.ToString() + "...");
                byte[] outputReport = new byte[selectedDevice.OutputReportLength];
                Array.Copy(report, 0, outputReport, 0, Math.Min(report.Length, outputReport.Length));
                uint bytesWritten;
                result = WriteFile(handle, outputReport, (uint)outputReport.Length, out bytesWritten, IntPtr.Zero);
                error = Marshal.GetLastWin32Error();
                Log("WriteFile (sized) result: " + result.ToString() + ", bytesWritten: " + bytesWritten.ToString() + ", error: " + error.ToString());
            }

            if (!result)
            {
                // Method 5: WriteFile with standard report
                Log("Trying WriteFile...");
                uint bytesWritten;
                result = WriteFile(handle, report, (uint)report.Length, out bytesWritten, IntPtr.Zero);
                error = Marshal.GetLastWin32Error();
                Log("WriteFile result: " + result.ToString() + ", bytesWritten: " + bytesWritten.ToString() + ", error: " + error.ToString());
            }

            if (result)
            {
                Log("SUCCESS: DPI set to " + dpi.ToString());
                resultLabel.Text = "DPI set to " + dpi.ToString();
                resultLabel.ForeColor = Color.Green;
            }
            else
            {
                Log("FAILED: All methods failed, last error: " + error.ToString());
                MessageBox.Show("Failed to set DPI. Error code: " + error.ToString() + "\nTry selecting a different device.\nCheck razer_dpi_log.txt for details.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                resultLabel.Text = "Failed!";
                resultLabel.ForeColor = Color.Red;
            }
        }
        finally
        {
            CloseHandle(handle);
            Log("Handle closed");
            SaveLog();
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
