
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing; // Für System.Drawing.Rectangle

namespace CopagoAutomation.Automation
{
    public class WindowAutomation
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int DefaultWindowTextCapacity = 512;
        private const int SW_RESTORE = 9;

        // Win32 API Konstanten
        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
        private const int INPUT_MOUSE = 0;
        private const int INPUT_KEYBOARD = 1;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;
        private const ushort VK_CONTROL = 0x11;
        private const ushort VK_A = 0x41;

        public const uint GA_ROOT = 2;
        public const uint MOD_NONE = 0x0000;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_WIN = 0x0008;

        // Win32 API Strukturen
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public int type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        // Win32 API Imports
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern void SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // Wrapper-Methoden
        public void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public void LeftClick()
        {
            var inputs = new INPUT[2];

            inputs[0] = new INPUT
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MOUSEEVENTF_LEFTDOWN
                    }
                }
            };

            inputs[1] = new INPUT
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MOUSEEVENTF_LEFTUP
                    }
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        }

        public void KeyPress(ushort virtualKey)
        {
            KeyDown(virtualKey);
            KeyUp(virtualKey);
        }

        public void KeyDown(ushort virtualKey)
        {
            var input = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey,
                        dwFlags = 0
                    }
                }
            };

            SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
        }

        public void KeyUp(ushort virtualKey)
        {
            var input = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey,
                        dwFlags = KEYEVENTF_KEYUP
                    }
                }
            };

            SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
        }

        public void SendUnicodeChar(char ch)
        {
            var keyDown = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = ch,
                        dwFlags = KEYEVENTF_UNICODE
                    }
                }
            };

            var keyUp = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = ch,
                        dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
                    }
                }
            };

            var inputs = new[] { keyDown, keyUp };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        }

        public void SelectAll()
        {
            KeyDown(VK_CONTROL);
            KeyPress(VK_A);
            KeyUp(VK_CONTROL);
        }

        public void TypeText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            foreach (char ch in text)
            {
                SendUnicodeChar(ch);
            }
        }

        public float GetDpiScaleFactor()
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            if (hdc == IntPtr.Zero)
                return 1.0f; // Fallback

            int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            ReleaseDC(IntPtr.Zero, hdc);

            return (float)dpiX / 96.0f;
        }

        public IntPtr GetActiveWindowHandle()
        {
            return GetForegroundWindow();
        }

        public System.Drawing.Point GetCursorScreenPosition()
        {
            POINT p;
            GetCursorPos(out p);
            return new System.Drawing.Point(p.X, p.Y);
        }

        public bool HasActiveWindow()
        {
            return GetForegroundWindow() != IntPtr.Zero;
        }

        public bool TryBindWindowByHandle(IntPtr hWnd, out BoundWindowInfo? boundWindow)
        {
            boundWindow = null;
            if (!IsWindow(hWnd))
                return false;

            var windowText = GetWindowText(hWnd);
            if (string.IsNullOrEmpty(windowText) || !windowText.Contains("Copago"))
                return false;

            boundWindow = new BoundWindowInfo(hWnd, windowText);
            return true;
        }

        public string GetWindowText(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0) return string.Empty;

            var sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public class BoundWindowInfo
        {
            public IntPtr Handle { get; }
            public string Title { get; }
            public Rectangle ClientRectangle { get; }

            public BoundWindowInfo(IntPtr handle, string title)
            {
                Handle = handle;
                Title = title;

                RECT clientRect;
                GetClientRect(handle, out clientRect);
                ClientRectangle = new Rectangle(clientRect.Left, clientRect.Top, clientRect.Right - clientRect.Left, clientRect.Bottom - clientRect.Top);
            }
        }
    }
}
