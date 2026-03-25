using System;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Interop;

namespace CopagoAutomation.Services
{
    public class GlobalHotkeyService : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9000;

        // Modifiers
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_WIN = 0x0008;

        // Virtual Keys
        private const uint VK_0 = 0x30;

        private IntPtr _windowHandle;
        private HwndSource? _hwndSource;

        public event EventHandler? HotkeyPressed;

        public GlobalHotkeyService(IntPtr windowHandle)
        {
            _windowHandle = windowHandle;
            _hwndSource = HwndSource.FromHwnd(_windowHandle);
            _hwndSource?.AddHook(WndProc);
        }

        public bool RegisterHotkey()
        {
            // Register Ctrl + Alt + 0
            return RegisterHotKey(_windowHandle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_0);
        }

        public void UnregisterHotkey()
        {
            UnregisterHotKey(_windowHandle, HOTKEY_ID);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;

            if (msg == WM_HOTKEY)
            {
                if (wParam.ToInt32() == HOTKEY_ID)
                {
                    HotkeyPressed?.Invoke(this, EventArgs.Empty);
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }

        public void Dispose()
        {
            UnregisterHotkey();
            if (_hwndSource != null)
            {
                _hwndSource.RemoveHook(WndProc);
                _hwndSource.Dispose();
            }
        }
    }
}
