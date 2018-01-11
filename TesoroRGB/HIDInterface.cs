using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.IO;

namespace TesoroRGB
{
    internal class HIDDevice
    {
        #region constants
        private const int DIGCF_DEFAULT = 0x1;
        private const int DIGCF_PRESENT = 0x2;
        private const int DIGCF_ALLCLASSES = 0x4;
        private const int DIGCF_PROFILE = 0x8;
        private const int DIGCF_DEVICEINTERFACE = 0x10;

        private const short FILE_ATTRIBUTE_NORMAL = 0x80;
        private const short INVALID_HANDLE_VALUE = -1;
        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint CREATE_NEW = 1;
        private const uint CREATE_ALWAYS = 2;
        private const uint OPEN_EXISTING = 3;

        #endregion

        #region win32_API_declarations
        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr SetupDiGetClassDevs(ref Guid ClassGuid,
                                                      IntPtr Enumerator,
                                                      IntPtr hwndParent,
                                                      uint Flags);

        [DllImport("hid.dll", SetLastError = true)]
        private static extern void HidD_GetHidGuid(ref Guid hidGuid);

        [DllImport(@"setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern Boolean SetupDiEnumDeviceInterfaces(
           IntPtr hDevInfo,
           //ref SP_DEVINFO_DATA devInfo,
           IntPtr devInfo,
           ref Guid interfaceClassGuid,
           UInt32 memberIndex,
           ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData
        );

        [DllImport(@"setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern Boolean SetupDiGetDeviceInterfaceDetail(
           IntPtr hDevInfo,
           ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
           ref SP_DEVICE_INTERFACE_DETAIL_DATA deviceInterfaceDetailData,
           UInt32 deviceInterfaceDetailDataSize,
           out UInt32 requiredSize,
           ref SP_DEVINFO_DATA deviceInfoData
        );

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string lpFileName, uint dwDesiredAccess,
            uint dwShareMode, IntPtr lpSecurityAttributes, uint dwCreationDisposition,
            uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool ReadFile(SafeFileHandle hFile, byte[] lpBuffer,
           uint nNumberOfBytesToRead, ref uint lpNumberOfBytesRead, IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool WriteFile(SafeFileHandle hFile, byte[] lpBuffer,
           uint nNumberOfBytesToWrite, ref uint lpNumberOfBytesWritten, IntPtr lpOverlapped);

        [DllImport("hid.dll", SetLastError = true)]
        private static extern bool HidD_GetPreparsedData(
            SafeFileHandle hObject,
            ref IntPtr PreparsedData);

        [DllImport("hid.dll", SetLastError = true)]
        private static extern Boolean HidD_FreePreparsedData(ref IntPtr PreparsedData);

        [DllImport("hid.dll", SetLastError = true)]
        private static extern int HidP_GetCaps(
            IntPtr pPHIDP_PREPARSED_DATA,					// IN PHIDP_PREPARSED_DATA  PreparsedData,
            ref HIDP_CAPS myPHIDP_CAPS);				// OUT PHIDP_CAPS  Capabilities

        [DllImport("hid.dll", SetLastError = true)]
        private static extern Boolean HidD_GetAttributes(SafeFileHandle hObject, ref HIDD_ATTRIBUTES Attributes);

        [DllImport("hid.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool HidD_GetFeature(
           SafeFileHandle hDevice,
           IntPtr hReportBuffer,
           uint ReportBufferLength);

        [DllImport("hid.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool HidD_SetFeature(
           SafeFileHandle hDevice,
           IntPtr ReportBuffer,
           uint ReportBufferLength);

        [DllImport("hid.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool HidD_GetProductString(
           SafeFileHandle hDevice,
           IntPtr Buffer,
           uint BufferLength);

        [DllImport("hid.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool HidD_GetSerialNumberString(
           SafeFileHandle hDevice,
           IntPtr Buffer,
           uint BufferLength);

        [DllImport("hid.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern Boolean HidD_GetManufacturerString(
            SafeFileHandle hDevice,
            IntPtr Buffer,
            uint BufferLength);

        #endregion

        #region structs

        public struct interfaceDetails
        {
            public string manufacturer;
            public string product;
            public int serialNumber;
            public ushort VID;
            public ushort PID;
            public string devicePath;
            public int IN_reportByteLength;
            public int OUT_reportByteLength;
            public int FEATURE_reportByteLength;
            public ushort versionNumber;
        }

        // HIDP_CAPS
        [StructLayout(LayoutKind.Sequential)]
        private struct HIDP_CAPS
        {
            public System.UInt16 Usage;					// USHORT
            public System.UInt16 UsagePage;				// USHORT
            public System.UInt16 InputReportByteLength;
            public System.UInt16 OutputReportByteLength;
            public System.UInt16 FeatureReportByteLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
            public System.UInt16[] Reserved;				// USHORT  Reserved[17];			
            public System.UInt16 NumberLinkCollectionNodes;
            public System.UInt16 NumberInputButtonCaps;
            public System.UInt16 NumberInputValueCaps;
            public System.UInt16 NumberInputDataIndices;
            public System.UInt16 NumberOutputButtonCaps;
            public System.UInt16 NumberOutputValueCaps;
            public System.UInt16 NumberOutputDataIndices;
            public System.UInt16 NumberFeatureButtonCaps;
            public System.UInt16 NumberFeatureValueCaps;
            public System.UInt16 NumberFeatureDataIndices;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public IntPtr Reserved;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVICE_INTERFACE_DATA
        {
            public uint cbSize;
            public Guid InterfaceClassGuid;
            public uint Flags;
            public IntPtr Reserved;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SP_DEVICE_INTERFACE_DETAIL_DATA
        {
            public int cbSize;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string DevicePath;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HIDD_ATTRIBUTES
        {
            public Int32 Size;
            public Int16 VendorID;
            public Int16 ProductID;
            public Int16 VersionNumber;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct COMMTIMEOUTS
        {
            public UInt32 ReadIntervalTimeout;
            public UInt32 ReadTotalTimeoutMultiplier;
            public UInt32 ReadTotalTimeoutConstant;
            public UInt32 WriteTotalTimeoutMultiplier;
            public UInt32 WriteTotalTimeoutConstant;
        }

        #endregion

        public bool deviceConnected { get; set; }
        private SafeFileHandle handle;
        private FileStream FS;
        private HIDP_CAPS capabilities;
        public delegate void dataReceivedEvent(byte[] message);


        /// <summary>
        /// Creates an object to handle read/write functionality for a USB HID device
        /// asnychronous read
        /// </summary>
        public static interfaceDetails[] getConnectedDevices()
        {
            interfaceDetails[] devices = new interfaceDetails[0];

            //Create structs to hold interface information
            SP_DEVINFO_DATA devInfo = new SP_DEVINFO_DATA();
            SP_DEVICE_INTERFACE_DATA devIface = new SP_DEVICE_INTERFACE_DATA();
            devInfo.cbSize = (uint)Marshal.SizeOf(devInfo);
            devIface.cbSize = (uint)(Marshal.SizeOf(devIface));

            Guid G = new Guid();
            HidD_GetHidGuid(ref G); //Get the guid of the HID device class

            IntPtr i = SetupDiGetClassDevs(ref G, IntPtr.Zero, IntPtr.Zero, DIGCF_DEVICEINTERFACE | DIGCF_PRESENT);

            //Loop through all available entries in the device list, until false
            SP_DEVICE_INTERFACE_DETAIL_DATA didd = new SP_DEVICE_INTERFACE_DETAIL_DATA();
            if (IntPtr.Size == 8) // for 64 bit operating systems
                didd.cbSize = 8;
            else
                didd.cbSize = 4 + Marshal.SystemDefaultCharSize; // for 32 bit systems

            int j = -1;
            bool b = true;
            int error;
            SafeFileHandle tempHandle;

            while (b)
            {
                j++;

                b = SetupDiEnumDeviceInterfaces(i, IntPtr.Zero, ref G, (uint)j, ref devIface);
                error = Marshal.GetLastWin32Error();
                if (b == false)
                    break;

                uint requiredSize = 0;
                bool b1 = SetupDiGetDeviceInterfaceDetail(i, ref devIface, ref didd, 256, out requiredSize, ref devInfo);
                string devicePath = didd.DevicePath;

                //create file handles using CT_CreateFile
                tempHandle = CreateFile(devicePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                    IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);

                //get capabilites - use getPreParsedData, and getCaps
                //store the reportlengths
                IntPtr ptrToPreParsedData = new IntPtr();
                bool ppdSucsess = HidD_GetPreparsedData(tempHandle, ref ptrToPreParsedData);
                if (ppdSucsess == false)
                    continue;

                HIDP_CAPS capabilities = new HIDP_CAPS();
                int hidCapsSucsess = HidP_GetCaps(ptrToPreParsedData, ref capabilities);

                HIDD_ATTRIBUTES attributes = new HIDD_ATTRIBUTES();
                bool hidAttribSucsess = HidD_GetAttributes(tempHandle, ref attributes);

                string productName = "";
                string manfString = "";
                IntPtr buffer = Marshal.AllocHGlobal(126);//max alloc for string; 
                if (HidD_GetProductString(tempHandle, buffer, 126)) productName = Marshal.PtrToStringAuto(buffer);
                if (HidD_GetManufacturerString(tempHandle, buffer, 126)) manfString = Marshal.PtrToStringAuto(buffer);
                Marshal.FreeHGlobal(buffer);

                //Call freePreParsedData to release some stuff
                HidD_FreePreparsedData(ref ptrToPreParsedData);

                //If connection was sucsessful, record the values in a global struct
                interfaceDetails productInfo = new interfaceDetails();
                productInfo.devicePath = devicePath;
                productInfo.manufacturer = manfString;
                productInfo.product = productName;
                productInfo.PID = (ushort)attributes.ProductID;
                productInfo.VID = (ushort)attributes.VendorID;
                productInfo.versionNumber = (ushort)attributes.VersionNumber;
                productInfo.IN_reportByteLength = (int)capabilities.InputReportByteLength;
                productInfo.OUT_reportByteLength = (int)capabilities.OutputReportByteLength;
                productInfo.FEATURE_reportByteLength = (int)capabilities.FeatureReportByteLength;

                int newSize = devices.Length + 1;
                Array.Resize(ref devices, newSize);
                devices[newSize - 1] = productInfo;
            }
            SetupDiDestroyDeviceInfoList(i);

            return devices;
        }

        /// <summary>
        /// Creates an object to handle read/write functionality for a USB HID device
        /// Uses one filestream to read/write
        /// </summary>
        /// <param name="devicePath">The USB device path - from getConnectedDevices</param>
        public HIDDevice(string devicePath)
        {
            initDevice(devicePath);

            if (!deviceConnected)
            {
                throw new Exception("Device could not be found");
            }
        }

        /// <summary>
        /// Opens and requests information from USB device
        /// Creates file stream to carry data.
        /// </summary>
        /// <param name="devicePath">The USB device path - from getConnectedDevices</param>
        private void initDevice(string devicePath)
        {
            deviceConnected = false;

            //create file handle using CT_CreateFile
            handle = CreateFile(devicePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);

            //get capabilites - use getPreParsedData, and getCaps
            //store the reportlengths
            IntPtr ptrToPreParsedData = new IntPtr();
            bool ppdSucsess = HidD_GetPreparsedData(handle, ref ptrToPreParsedData);

            capabilities = new HIDP_CAPS();
            int hidCapsSucsess = HidP_GetCaps(ptrToPreParsedData, ref capabilities);

            HIDD_ATTRIBUTES attributes = new HIDD_ATTRIBUTES();
            bool hidAttribSucsess = HidD_GetAttributes(handle, ref attributes);

            string productName = "";
            string SN = "";
            string manfString = "";
            IntPtr buffer = Marshal.AllocHGlobal(126);//max alloc for string; 
            if (HidD_GetProductString(handle, buffer, 126)) productName = Marshal.PtrToStringAuto(buffer);
            if (HidD_GetSerialNumberString(handle, buffer, 126)) SN = Marshal.PtrToStringAuto(buffer);
            if (HidD_GetManufacturerString(handle, buffer, 126)) manfString = Marshal.PtrToStringAuto(buffer);
            Marshal.FreeHGlobal(buffer);

            //Call freePreParsedData to release some stuff
            HidD_FreePreparsedData(ref ptrToPreParsedData);
            //SetupDiDestroyDeviceInfoList(i);

            if (handle.IsInvalid)
                return;

            deviceConnected = true;

            //use a filestream object to bring this stuff into .NET
            FS = new FileStream(handle, FileAccess.ReadWrite, capabilities.FeatureReportByteLength, false);
        }

        /// <summary>
        /// Closes filestream and handle of this device
        /// </summary>
        public void close()
        {
            if (FS != null)
                FS.Close();

            if ((handle != null) && (!(handle.IsInvalid)))
                handle.Close();

            deviceConnected = false;
        }

        /// <summary>
        /// Sends a feature report containing the provided data
        /// </summary>
        /// <param name="data">An array of bytes to be sent. Must be shorter than the devices feature report length</param>
        public void writeFeature(byte[] data)
        {
            if (data.Length > capabilities.FeatureReportByteLength)
                throw new Exception("Feature report must not exceed " + (capabilities.FeatureReportByteLength - 1).ToString() + " bytes");


            int size = Marshal.SizeOf(data[0]) * data.Length;

            IntPtr pnt = Marshal.AllocHGlobal(size);
            Marshal.Copy(data, 0, pnt, data.Length);

            HidD_SetFeature(handle, pnt, (uint)size);
        }
    }
}