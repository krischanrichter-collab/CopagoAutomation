
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
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

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

        // For Save Dialog Automation
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_SETTEXT = 0x000C;
        private const uint BM_CLICK = 0x00F5;
        private const int GW_CHILD = 5;
        private const int WM_GETTEXT = 0x000D;
        private const int WM_GETTEXTLENGTH = 0x000E;

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

        public void RightClick()
        {
            // Implementierung für RightClick, da sie in MouseAutomation.cs verwendet wird
            // MOUSEEVENTF_RIGHTDOWN = 0x0008, MOUSEEVENTF_RIGHTUP = 0x0010
            var inputs = new INPUT[2];

            inputs[0] = new INPUT
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = 0x0008 // MOUSEEVENTF_RIGHTDOWN
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
                        dwFlags = 0x0010 // MOUSEEVENTF_RIGHTUP
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

        public bool AutomateSaveDialog(string filePath, string dialogTitlePart, out string logMessage)
        {
            logMessage = string.Empty;
            IntPtr saveDialogHwnd = IntPtr.Zero;

            // Find the "Save As" dialog window
            EnumWindowsProc enumProc = (hWnd, lParam) =>
            {
                StringBuilder sb = new StringBuilder(DefaultWindowTextCapacity);
                GetWindowText(hWnd, sb, DefaultWindowTextCapacity);
                string windowTitle = sb.ToString();

                if (IsWindowVisible(hWnd) && windowTitle.Contains(dialogTitlePart))
                {
                    saveDialogHwnd = hWnd;
                    return false; // Stop enumeration
                }
                return true;
            };

            EnumWindows(enumProc, IntPtr.Zero);

            if (saveDialogHwnd == IntPtr.Zero)
            {
                logMessage = $"Save Dialog mit Titelteil '{dialogTitlePart}' nicht gefunden.";
                return false;
            }

            // Activate the dialog window
            SetForegroundWindow(saveDialogHwnd);

            // Find the file name text box (often a ComboBoxEx32 or Edit control)
            // This part might need adjustment based on the exact dialog structure
            IntPtr fileNameTextBoxHwnd = FindWindowEx(saveDialogHwnd, IntPtr.Zero, "DUIViewWndClassName", null); // Modern dialogs
            if (fileNameTextBoxHwnd == IntPtr.Zero)
            {
                fileNameTextBoxHwnd = FindWindowEx(saveDialogHwnd, IntPtr.Zero, "ComboBoxEx32", null);
                if (fileNameTextBoxHwnd != IntPtr.Zero)
                {
                    fileNameTextBoxHwnd = FindWindowEx(fileNameTextBoxHwnd, IntPtr.Zero, "Edit", null);
                }
            }
            if (fileNameTextBoxHwnd == IntPtr.Zero)
            {
                fileNameTextBoxHwnd = FindWindowEx(saveDialogHwnd, IntPtr.Zero, "Edit", null); // Older dialogs
            }

            if (fileNameTextBoxHwnd == IntPtr.Zero)
            {
                logMessage = "Dateiname-Textfeld im Save Dialog nicht gefunden.";
                return false;
            }

            // Set the file path
            SendMessage(fileNameTextBoxHwnd, WM_SETTEXT, IntPtr.Zero, filePath);
            // Find and click the "Save" button
            IntPtr saveButtonHwnd = FindWindowEx(saveDialogHwnd, IntPtr.Zero, "Button", "Speichern");
            if (saveButtonHwnd == IntPtr.Zero)
            {
                // Fallback for different button text or if it's a default button
                saveButtonHwnd = FindWindowEx(saveDialogHwnd, IntPtr.Zero, "Button", null); // Try to find any button
            }

            if (saveButtonHwnd == IntPtr.Zero)
            {
                logMessage = "'Speichern'-Button im Save Dialog nicht gefunden.";
                return false;
            }

            // Click the "Save" button
            SendMessage(saveButtonHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);

            logMessage = "Save Dialog erfolgreich automatisiert.";
            return true;
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
            IntPtr foundHandle = IntPtr.Zero;
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd) && GetWindowTitle(hWnd).Contains(titlePart, StringComparison.OrdinalIgnoreCase))
                {
                    foundHandle = hWnd;
                    return false; // Stop enumeration
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);
            handle = foundHandle;
            return handle != IntPtr.Zero;
        }

        public bool TryBindWindowByTitleContains(string titlePart, out BoundWindowInfo boundWindow)
        {
            boundWindow = new BoundWindowInfo(IntPtr.Zero, string.Empty);
            if (TryFindWindowByTitleContains(titlePart, out var handle))
            {
                boundWindow = new BoundWindowInfo(handle, GetWindowTitle(handle), GetClientRect(handle));
                return true;
            }
            return false;
        }

        public bool TryBindWindowByHandle(IntPtr handle, out BoundWindowInfo boundWindow)
        {
            boundWindow = new BoundWindowInfo(IntPtr.Zero, string.Empty);
            if (IsValidHandle(handle))
            {
                boundWindow = new BoundWindowInfo(handle, GetWindowTitle(handle), GetClientRect(handle));
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

        public bool TryGetClientRect(IntPtr hWnd, out Rectangle clientRect)
        {
            clientRect = Rectangle.Empty;
            if (!IsValidHandle(hWnd))
                return false;

            clientRect = GetClientRect(hWnd);
            return true;
        }

        public bool IsValidHandle(IntPtr handle)
        {
            return handle != IntPtr.Zero && IsWindow(handle);
        }

        public bool IsWindowValid(BoundWindowInfo boundWindow)
        {
            return IsValidHandle(boundWindow.Handle);
        }


    }
}
