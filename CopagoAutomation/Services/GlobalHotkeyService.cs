using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace CopagoAutomation.Services
{
    public class GlobalHotkeyService : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const uint MOD_ALT     = 0x0001;
        private const uint MOD_CONTROL = 0x0002;

        // IDs: 9001 = Ctrl+Alt+1, 9002 = Ctrl+Alt+2, ..., 9009 = Ctrl+Alt+9
        private const int HOTKEY_ID_BASE = 9000;

        private readonly IntPtr _windowHandle;
        private HwndSource? _hwndSource;

        private int _registeredFrom;
        private int _registeredTo;

        /// <summary>Fired when a registered hotkey is pressed. The int argument is the digit (1–9).</summary>
        public event EventHandler<int>? HotkeyPressed;

        public GlobalHotkeyService(IntPtr windowHandle)
        {
            _windowHandle = windowHandle;
            _hwndSource = HwndSource.FromHwnd(_windowHandle);
            _hwndSource?.AddHook(WndProc);
        }

        /// <summary>Registers Ctrl+Alt+<fromDigit> through Ctrl+Alt+<toDigit>.</summary>
        public bool RegisterHotkeys(int fromDigit = 1, int toDigit = 9)
        {
            _registeredFrom = fromDigit;
            _registeredTo   = toDigit;

            bool allOk = true;
            for (int digit = fromDigit; digit <= toDigit; digit++)
            {
                uint vk = (uint)(0x30 + digit); // VK_1 = 0x31 … VK_9 = 0x39
                if (!RegisterHotKey(_windowHandle, HOTKEY_ID_BASE + digit, MOD_CONTROL | MOD_ALT, vk))
                    allOk = false;
            }
            return allOk;
        }

        private void UnregisterAll()
        {
            for (int digit = _registeredFrom; digit <= _registeredTo; digit++)
                UnregisterHotKey(_windowHandle, HOTKEY_ID_BASE + digit);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            if (msg == WM_HOTKEY)
            {
                int digit = wParam.ToInt32() - HOTKEY_ID_BASE;
                if (digit >= 1 && digit <= 9)
                {
                    HotkeyPressed?.Invoke(this, digit);
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }

        public void Dispose()
        {
            UnregisterAll();
            if (_hwndSource != null)
            {
                _hwndSource.RemoveHook(WndProc);
                // Do NOT dispose _hwndSource — it was obtained via HwndSource.FromHwnd()
                // which returns the shared WPF source owned by the window, not by us.
                _hwndSource = null;
            }
        }
    }
}
