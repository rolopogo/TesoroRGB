using System;
using System.Drawing; // for Image support
using System.Threading; // for sleep
using System.Diagnostics; // For stopwatch

namespace TesoroRGB
{
    public enum TesoroProfile : byte
    {
        Profile1 = 0x01,
        Profile2 = 0x02,
        Profile3 = 0x03,
        Profile4 = 0x04,
        Profile5 = 0x05,
        PC = 0x06
    }

    public enum LightingMode : byte
    {
        Shine = 0x00,
        Trigger = 0x01,
        Ripple = 0x02,
        Fireworks = 0x03,
        Radiation = 0x04,
        Breathing = 0x05,
        RainbowWave = 0x06,
        SpectrumColors = 0x08
    }

    public enum SpectrumMode : byte
    {
        SpectrumShine = 0x00,
        SpectrumBreathing = 0x01,
        SpectrumTrigger = 0x02
    }

    public enum TesoroLedID : byte
    {
        Escape = 0x0B,
        F1 = 0x16,
        F2 = 0x1E,
        F3 = 0x19,
        F4 = 0x1B,
        F5 = 0X07,
        F6 = 0X33,
        F7 = 0X39,
        F8 = 0X3e,
        F9 = 0X56,
        F10 = 0X57,
        F11 = 0X53,
        F12 = 0X55,
        PrintScreen = 0x4F,
        ScrollLock = 0x48,
        Pause = 0x00,
        Oemtilde = 0x0E, // `¬
        D1 = 0x0F,
        D2 = 0x17,
        D3 = 0x1f,
        D4 = 0x27,
        D5 = 0x26,
        D6 = 0x2E,
        D7 = 0x2f,
        D8 = 0x37,
        D9 = 0x3F,
        D0 = 0x47,
        OemMinus = 0x46,
        OemPlus = 0x36,
        Back = 0x51,
        Insert = 0x66,
        Home = 0x76,
        PageUp = 0x6E,
        Tab = 0x09,
        Q = 0x08,
        W = 0x10,
        E = 0x18,
        R = 0x20,
        T = 0x21,
        Y = 0x29,
        U = 0x28,
        I = 0x30,
        O = 0x38,
        P = 0x40,
        OemOpenBrackets = 0x41,
        OemCloseBrackets = 0x31,
        OemPipe = 0x52,
        Delete = 0x5E,
        End = 0x77,
        PageDown = 0x6F,
        CapsLock = 0x11,
        A = 0x0A,
        S = 0x12,
        D = 0x1A,
        F = 0x22,
        G = 0x23,
        H = 0x2B,
        J = 0x2A,
        K = 0x32,
        L = 0x3A,
        OemSemicolon = 0x42,
        Apostrophe = 0x43,
        Enter = 0x54,
        LeftShift = 0x79,
        Z = 0x0C,
        X = 0x14,
        C = 0x1C,
        V = 0x24,
        B = 0x25,
        N = 0x2D,
        M = 0x2C,
        Comma = 0x34,
        Period = 0x3C,
        Slash = 0x45,
        RightShift = 0x7A,
        Up = 0x73,
        LeftControl = 0x06,
        Windows = 0x7C,
        Alt = 0x4B,
        Space = 0x5B,
        AltGr = 0x4D,
        TesoroFn = 0x7D,
        Apps = 0x3D,
        RightControl = 0x04,
        Left = 0x75,
        Down = 0x5D,
        Right = 0x65,
        NumLock = 0x5C,
        Divide = 0x64,
        Multiply = 0x6C,
        Subtract = 0x6D,
        NumPad7 = 0x58,
        NumPad8 = 0x60,
        NumPad9 = 0x68,
        NumPad4 = 0x59,
        NumPad5 = 0x61,
        NumPad6 = 0x69,
        Add = 0x70,
        NumPad1 = 0x5A,
        NumPad2 = 0x62,
        NumPad3 = 0x6A,
        NumPad0 = 0x63,
        Decimal = 0x6B,
        NumPadEnter = 0x72,

        None = 0xFF
    }

    public class Keyboard
    {
        HIDDevice device;
        public static int width = 22;
        public static int height = 6;
        TesoroLedID[,] keyPositions = new TesoroLedID[width, height];

        // Attempt to find and open communications with a Tesoro Keyboard
        // Returns true if successful, false if not
        /// <summary>
        /// Close communications with the USB device.
        /// </summary>
        /// <returns>True if successful, False if the device couldn't be found</returns>
        public bool Initialize()
        {
            //Get the details of all connected USB HID devices
            HIDDevice.interfaceDetails[] devices = HIDDevice.getConnectedDevices();

            string devicePath = "";

            // Loop through these to find our device
            foreach (HIDDevice.interfaceDetails dev in devices)
            {
                // check vendor ID to find tesoro devices
                if (dev.devicePath.Contains("hid#vid_195d"))
                {
                    // find this particular device - Tested on Gram Spectrum Only!
                    // TODO: find a cleaner way to find Tesoro Keyboards
                    if (dev.devicePath.Contains("&mi_01&col05"))
                    {
                        // found correct device
                        devicePath = dev.devicePath;
                    }
                }
            }

            if (devicePath == "")
            {
                // no device found
                return false;
            }

            // prepare the key positions array
            SetKeyPositions();

            //open the device
            device = new HIDDevice(devicePath);
            // report success
            return true;
        }

        /// <summary>
        /// Close communications with the USB device.
        /// </summary>
        public void Uninitialize()
        {
            if (device.deviceConnected) device.close();
        }

        /// <summary>
        /// Set the keyboard to display the designated profile
        /// </summary>
        /// <param name="profile">The profile to change to.</param>
        public void SetProfile(TesoroProfile profile)
        {
            // Prepare the command data
            byte[] data = { 0x07, 0x03, (byte)profile, 0x00, 0x00, 0x00, 0x00, 0x00 };

            // Send the data
            device.writeFeature(data);
        }

        /// <summary>
        /// Sets the main background Color for standard effects
        /// </summary>
        /// <param name="mode">Lighting mode</param>
        /// <param name="profile">The profile to modify.</param>
        public void SetLightingMode(LightingMode mode, TesoroProfile profile)
        {
            SetLightingMode(mode, 0x00, profile);
        }

        /// <summary>
        /// Sets the main background Color for standard effects
        /// </summary>
        /// <param name="mode">Lighting mode</param>
        /// <param name="spectrumMode">Lighting sub-mode for spectrum Color. Should be 0x00 for non-spectrum modes</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void SetLightingMode(LightingMode mode, SpectrumMode spectrumMode, TesoroProfile profile)
        {
            // Prepare the command data
            byte[] data = { 0x07, 0x0A, (byte)profile, (byte)mode, (byte)spectrumMode, 0x00, 0x00, 0x00 };

            // Send the data
            device.writeFeature(data);
        }

        /// <summary>
        /// Sets the main background Color for standard effects
        /// </summary>
        /// <param name="r">Red exponent 0-255</param>
        /// <param name="g">Green exponent 0-255</param>
        /// <param name="b">Blue exponent 0-255</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void SetProfileColor(int r, int g, int b, TesoroProfile profile)
        {
            // Prepare the command data
            byte[] data = { 0x07, 0x0B, (byte)profile, IntToByte(r), IntToByte(g), IntToByte(b), 0x00, 0x00 };

            // Send the data
            device.writeFeature(data);
        }

        /// <summary>
        /// Sets the LED of a single key using 0-255 integers.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="r">Red exponent 0-255</param>
        /// <param name="g">Green exponent 0-255</param>
        /// <param name="b">Blue exponent 0-255</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void SetKeyColor(TesoroLedID key, int r, int g, int b, TesoroProfile profile)
        {
            if (key == TesoroLedID.None) return;

            // Prepare the command data
            byte[] data = { 0x07, 0x0D, (byte)profile, (byte)key, IntToByte(r), IntToByte(g), IntToByte(b), 0x00 };

            // Send the data
            device.writeFeature(data);
        }

        /// <summary>
        /// Set all LEDs using a Bitmap object. Image will be scaled to fit the keyboard
        /// </summary>
        /// <param name="bitmap">A Bitmap object describing the desired pattern.</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        /// <param name="fast">True - speeds up execution at the risk of missed pixels. Recomended true for animated effects, false for static effects.</param>
        void SetKeysColor(Bitmap bitmap, TesoroProfile profile, bool fast)
        {
            if (bitmap.Width > width || bitmap.Height > height) bitmap = new Bitmap(bitmap, width, height);
            for (int x = 0; x < width; x++)
            {
                for (int y = 0; y < height; y++)
                {
                    // get pixel (using tiling for small bitmaps)
                    Color col = bitmap.GetPixel(x, y);

                    // set Color
                    SetKeyColor(keyPositions[x, y], col.R, col.G, col.B, profile);
                    // delay between sends
                    Wait(fast);
                }
            }
        }

        /// <summary>
        /// Set the LEDs for all using an image file. Image will be scaled to fit the keyboard
        /// </summary>
        /// <param name="file">Path to an image file describing the desired pattern.</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        /// <param name="fast">True - speeds up execution at the risk of missed pixels. Recomended true for animated effects, false for static effects.</param>
        public void SetKeysColor(string file, TesoroProfile profile, bool fast)
        {
            Image image = Image.FromFile(file);

            SetKeysColor(new Bitmap(image), profile, fast);
        }

        /// <summary>
        /// Turn off all LEDs.
        /// </summary>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void ClearSpectrumColors(TesoroProfile profile)
        {
            // Prepare the command data
            byte[] data = { 0x07, 0x0D, (byte)profile, 0xFE, 0x00, 0x00, 0x00, 0x00 };

            // Send the data
            device.writeFeature(data);
        }

        /// <summary>
        /// Save the current Spectrum Color layout to the keyboard.
        /// Changing profile without saving first will lose all changes made to the profile.
        /// </summary>
        /// <param name="profile">The profile to save. This should be the active profile.</param>
        public void SaveSpectrumColors(TesoroProfile profile)
        {
            // Prepare the command data
            byte[] data = { 0x07, 0x0D, (byte)profile, 0xFF, 0x00, 0x00, 0x00, 0x00 };

            // Send the data
            device.writeFeature(data);
        }

        /// <summary>
        /// Perform a short wait. For use in between successive calls.
        /// </summary>
        /// <param name="fast">True - Shorter delay for use in animation effect. When keys are set with fast delay, some leds may not be set. False - Longer delay for safely assigning</param>
        void Wait(bool fast)
        {
            if (fast)
            {
                Stopwatch watch = new Stopwatch();
                watch.Start();
                while (watch.Elapsed.TotalMilliseconds < 0.8) { } // waiting for .8ms
            }
            else
            {
                Thread.Sleep(1);
            }
        }
        
        /// <summary>
        /// Initialises the key positions array.
        /// </summary>
        void SetKeyPositions()
        {
            for (int x = 0; x < width; x++)
            {
                for (int y = 0; y < height; y++)
                {
                    SetKeyPosition(TesoroLedID.None, x, y);
                }
            }
            // Row 1
            SetKeyPosition(TesoroLedID.Escape, 0, 0);
            SetKeyPosition(TesoroLedID.F1, 2, 0);
            SetKeyPosition(TesoroLedID.F2, 3, 0);
            SetKeyPosition(TesoroLedID.F3, 4, 0);
            SetKeyPosition(TesoroLedID.F4, 5, 0);
            SetKeyPosition(TesoroLedID.F5, 6, 0);
            SetKeyPosition(TesoroLedID.F6, 7, 0);
            SetKeyPosition(TesoroLedID.F7, 8, 0);
            SetKeyPosition(TesoroLedID.F8, 9, 0);
            SetKeyPosition(TesoroLedID.F9, 11, 0);
            SetKeyPosition(TesoroLedID.F10, 12, 0);
            SetKeyPosition(TesoroLedID.F11, 13, 0);
            SetKeyPosition(TesoroLedID.F12, 14, 0);
            SetKeyPosition(TesoroLedID.PrintScreen, 15, 0);
            SetKeyPosition(TesoroLedID.ScrollLock, 16, 0);
            SetKeyPosition(TesoroLedID.Pause, 17, 0);

            // Row 2
            SetKeyPosition(TesoroLedID.Oemtilde, 0, 1);
            SetKeyPosition(TesoroLedID.D1, 1, 1);
            SetKeyPosition(TesoroLedID.D2, 2, 1);
            SetKeyPosition(TesoroLedID.D3, 3, 1);
            SetKeyPosition(TesoroLedID.D4, 4, 1);
            SetKeyPosition(TesoroLedID.D5, 5, 1);
            SetKeyPosition(TesoroLedID.D6, 6, 1);
            SetKeyPosition(TesoroLedID.D7, 7, 1);
            SetKeyPosition(TesoroLedID.D8, 8, 1);
            SetKeyPosition(TesoroLedID.D9, 9, 1);
            SetKeyPosition(TesoroLedID.D0, 10, 1);
            SetKeyPosition(TesoroLedID.OemMinus, 11, 1);
            SetKeyPosition(TesoroLedID.OemPlus, 12, 1);
            SetKeyPosition(TesoroLedID.Back, 13, 1);
            SetKeyPosition(TesoroLedID.Insert, 15, 1);
            SetKeyPosition(TesoroLedID.Home, 16, 1);
            SetKeyPosition(TesoroLedID.PageUp, 17, 1);
            SetKeyPosition(TesoroLedID.NumLock, 18, 1);
            SetKeyPosition(TesoroLedID.Divide, 19, 1);
            SetKeyPosition(TesoroLedID.Multiply, 20, 1);
            SetKeyPosition(TesoroLedID.Subtract, 21, 1);

            // Row 3
            SetKeyPosition(TesoroLedID.Tab, 0, 2);
            SetKeyPosition(TesoroLedID.Q, 1, 2);
            SetKeyPosition(TesoroLedID.W, 2, 2);
            SetKeyPosition(TesoroLedID.E, 3, 2);
            SetKeyPosition(TesoroLedID.R, 5, 2);
            SetKeyPosition(TesoroLedID.T, 6, 2);
            SetKeyPosition(TesoroLedID.Y, 7, 2);
            SetKeyPosition(TesoroLedID.U, 8, 2);
            SetKeyPosition(TesoroLedID.I, 9, 2);
            SetKeyPosition(TesoroLedID.O, 10, 2);
            SetKeyPosition(TesoroLedID.P, 11, 2);
            SetKeyPosition(TesoroLedID.OemOpenBrackets, 12, 2);
            SetKeyPosition(TesoroLedID.OemCloseBrackets, 13, 2);
            SetKeyPosition(TesoroLedID.OemPipe, 14, 2);
            SetKeyPosition(TesoroLedID.Delete, 15, 2);
            SetKeyPosition(TesoroLedID.End, 16, 2);
            SetKeyPosition(TesoroLedID.PageDown, 17, 2);
            SetKeyPosition(TesoroLedID.NumPad7, 18, 2);
            SetKeyPosition(TesoroLedID.NumPad8, 19, 2);
            SetKeyPosition(TesoroLedID.NumPad9, 20, 2);
            SetKeyPosition(TesoroLedID.Add, 21, 2);

            // Row 4
            SetKeyPosition(TesoroLedID.CapsLock, 0, 3);
            SetKeyPosition(TesoroLedID.A, 1, 3);
            SetKeyPosition(TesoroLedID.S, 3, 3);
            SetKeyPosition(TesoroLedID.D, 4, 3);
            SetKeyPosition(TesoroLedID.F, 5, 3);
            SetKeyPosition(TesoroLedID.G, 6, 3);
            SetKeyPosition(TesoroLedID.H, 7, 3);
            SetKeyPosition(TesoroLedID.J, 8, 3);
            SetKeyPosition(TesoroLedID.K, 9, 3);
            SetKeyPosition(TesoroLedID.L, 10, 3);
            SetKeyPosition(TesoroLedID.OemSemicolon, 11, 3);
            SetKeyPosition(TesoroLedID.Apostrophe, 12, 3);
            SetKeyPosition(TesoroLedID.Enter, 14, 3);
            SetKeyPosition(TesoroLedID.NumPad4, 17, 3);
            SetKeyPosition(TesoroLedID.NumPad5, 18, 3);
            SetKeyPosition(TesoroLedID.NumPad6, 19, 3);

            // Row 5
            SetKeyPosition(TesoroLedID.LeftShift, 1, 4);
            SetKeyPosition(TesoroLedID.Z, 2, 4);
            SetKeyPosition(TesoroLedID.X, 3, 4);
            SetKeyPosition(TesoroLedID.C, 4, 4);
            SetKeyPosition(TesoroLedID.V, 5, 4);
            SetKeyPosition(TesoroLedID.B, 6, 4);
            SetKeyPosition(TesoroLedID.N, 7, 4);
            SetKeyPosition(TesoroLedID.M, 8, 4);
            SetKeyPosition(TesoroLedID.Comma, 9, 4);
            SetKeyPosition(TesoroLedID.Period, 10, 4);
            SetKeyPosition(TesoroLedID.Slash, 11, 4);
            SetKeyPosition(TesoroLedID.RightShift, 13, 4);
            SetKeyPosition(TesoroLedID.Up, 16, 4);
            SetKeyPosition(TesoroLedID.NumPad1, 18, 4);
            SetKeyPosition(TesoroLedID.NumPad2, 19, 4);
            SetKeyPosition(TesoroLedID.NumPad3, 20, 4);
            SetKeyPosition(TesoroLedID.NumPadEnter, 21, 4);

            // Row 6
            SetKeyPosition(TesoroLedID.LeftControl, 0, 5);
            SetKeyPosition(TesoroLedID.Windows, 1, 5);
            SetKeyPosition(TesoroLedID.Alt, 3, 5);
            SetKeyPosition(TesoroLedID.Space, 6, 5);
            SetKeyPosition(TesoroLedID.AltGr, 11, 5);
            SetKeyPosition(TesoroLedID.TesoroFn, 12, 5);
            SetKeyPosition(TesoroLedID.Apps, 13, 5);
            SetKeyPosition(TesoroLedID.RightControl, 14, 5);
            SetKeyPosition(TesoroLedID.Left, 15, 5);
            SetKeyPosition(TesoroLedID.Down, 16, 5);
            SetKeyPosition(TesoroLedID.Right, 17, 5);
            SetKeyPosition(TesoroLedID.NumPad0, 18, 5);
            SetKeyPosition(TesoroLedID.Decimal, 20, 5);
        }

        void SetKeyPosition(TesoroLedID key, int x, int y)
        {
            if ((x >= width) || (y >= height))
            {
                // out of bounds
                return;
            }
            keyPositions[x, y] = key;
        }

        // Returns the integer as a byte, truncates to one byte
        byte IntToByte(int i)
        {
            return BitConverter.GetBytes(i % 256)[0];
        }
    }
}
