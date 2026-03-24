using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing; // Für System.Drawing.Rectangle

// WPF-Projekte benötigen normalerweise keinen direkten Verweis auf System.Windows.Forms.
// Wenn Screen.PrimaryScreen.Bounds benötigt wird, muss der Verweis im Projekt hinzugefügt werden.
// Für GetDpiScaleFactor und andere Win32-Aufrufe ist System.Windows.Forms nicht direkt erforderlich.
// Wenn es nur für die RECT-Struktur war, wird jetzt System.Drawing.Rectangle verwendet.

namespace CopagoAutomation.Automation
{
    public class WindowAutomation
    {
        // Delegat muss public sein, da er in einer public Methode verwendet wird
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int DefaultWindowTextCapacity = 512;
        private const int SW_RESTORE = 9;

        // Win32 API Imports
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

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
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern void SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int INPUT_MOUSE = 0;
        private const int INPUT_KEYBOARD = 1;

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        private const ushort VK_CONTROL = 0x11;
        private const ushort VK_A = 0x41;

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

        // INPUT struct muss public sein, da sie in einer public Methode verwendet wird
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public int type;
            public InputUnion U;
        }

        // InputUnion muss public sein
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

        // MOUSEINPUT muss public sein
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

        // KEYBDINPUT muss public sein
        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        // HARDWAREINPUT muss public sein
        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

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

            // Standard-DPI ist 96
            return (float)dpiX / 96.0f;
        }

        public IntPtr GetActiveWindowHandle()
        {
            return GetForegroundWindow();
        }

        public bool HasActiveWindow()
        {
            return IsValidHandle(GetActiveWindowHandle());
        }

        public string GetActiveWindowTitle()
        {
            IntPtr handle = GetActiveWindowHandle();

            if (!IsValidHandle(handle))
                return string.Empty;

            return GetWindowTitle(handle);
        }

        public bool IsActiveWindowTitle(string expectedTitle, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            if (string.IsNullOrWhiteSpace(expectedTitle))
                return false;

            string activeTitle = GetActiveWindowTitle();
            if (string.IsNullOrWhiteSpace(activeTitle))
                return false;

            return string.Equals(activeTitle, expectedTitle, comparison);
        }

        public bool ActiveWindowTitleContains(string text, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string activeTitle = GetActiveWindowTitle();
            if (string.IsNullOrWhiteSpace(activeTitle))
                return false;

            return activeTitle.IndexOf(text, comparison) >= 0;
        }

        public bool TryActivateWindow(IntPtr handle)
        {
            if (!IsValidHandle(handle))
                return false;

            if (IsIconic(handle))
                ShowWindow(handle, SW_RESTORE);

            if (SetForegroundWindow(handle))
                return true;

            return GetForegroundWindow() == handle;
        }

        public bool TryActivateBoundWindow(BoundWindowInfo boundWindow)
        {
            if (!IsWindowValid(boundWindow))
                return false;

            return TryActivateWindow(boundWindow.Handle);
        }

        public bool IsBoundWindowActive(BoundWindowInfo boundWindow)
        {
            if (!IsWindowValid(boundWindow))
                return false;

            return GetForegroundWindow() == boundWindow.Handle;
        }

        public bool TryFindWindowByTitleContains(string titlePart, out IntPtr handle)
        {
            handle = IntPtr.Zero;
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd) && GetWindowTitle(hWnd).Contains(titlePart, StringComparison.OrdinalIgnoreCase))
                {
                    handle = hWnd;
                    return false; // Stop enumeration
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);
            return handle != IntPtr.Zero;
        }

        public bool TryBindWindowByTitleContains(string titlePart, out BoundWindowInfo boundWindow)
        {
            boundWindow = BoundWindowInfo.Empty;
            if (TryFindWindowByTitleContains(titlePart, out var handle))
            {
                boundWindow = new BoundWindowInfo(handle, GetWindowTitle(handle), GetWindowRect(handle), GetClientRect(handle));
                return true;
            }
            return false;
        }

        public bool TryBindWindowByHandle(IntPtr handle, out BoundWindowInfo boundWindow)
        {
            boundWindow = BoundWindowInfo.Empty;
            if (IsValidHandle(handle))
            {
                boundWindow = new BoundWindowInfo(handle, GetWindowTitle(handle), GetWindowRect(handle), GetClientRect(handle));
                return true;
            }
            return false;
        }

        public string GetWindowTitle(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0)
                return string.Empty;

            StringBuilder sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public Rectangle GetWindowRect(IntPtr hWnd)
        {
            RECT rect = new RECT();
            GetWindowRect(hWnd, out rect);
            return new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
        }

        public Rectangle GetClientRect(IntPtr hWnd)
        {
            RECT rect = new RECT();
            GetClientRect(hWnd, out rect);
            // ClientRect ist relativ zum Client-Bereich, muss in Bildschirmkoordinaten umgerechnet werden
            POINT topLeft = new POINT { X = rect.Left, Y = rect.Top };
            ClientToScreen(hWnd, ref topLeft);
            return new Rectangle(topLeft.X, topLeft.Y, rect.Right - rect.Left, rect.Bottom - rect.Top);
        }

        public bool IsValidHandle(IntPtr handle)
        {
            return handle != IntPtr.Zero && IsWindow(handle);
        }

        public bool IsWindowValid(BoundWindowInfo boundWindow)
        {
            return IsValidHandle(boundWindow.Handle);
        }

        // BoundWindowInfo als struct, da es nur Daten hält und keine komplexen Operationen hat
        // Properties müssen get; set; sein, wenn die struct nicht readonly ist, oder readonly, wenn die struct readonly ist.
        // Da wir die Werte im Konstruktor setzen, kann die struct readonly sein und die Properties auch.
        public readonly struct BoundWindowInfo
        {
            public IntPtr Handle { get; }
            public string Title { get; }
            public Rectangle WindowRect { get; }
            public Rectangle ClientRect { get; }

            public BoundWindowInfo(IntPtr handle, string title, Rectangle windowRect, Rectangle clientRect)
            {
                Handle = handle;
                Title = title;
                WindowRect = windowRect;
                ClientRect = clientRect;
            }

            public static BoundWindowInfo Empty => new BoundWindowInfo(IntPtr.Zero, string.Empty, Rectangle.Empty, Rectangle.Empty);

            public bool IsEmpty => Handle == IntPtr.Zero;
        }
    }
}
