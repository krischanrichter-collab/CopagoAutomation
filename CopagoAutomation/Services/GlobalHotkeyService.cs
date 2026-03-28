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

        // IDs: 9001 = Ctrl+Alt+Q, 9002 = Ctrl+Alt+W, ..., 9010 = Ctrl+Alt+P
        private const int HOTKEY_ID_BASE = 9000;

        // Deutsche Tastatur, obere Reihe: Q W E R T Z U I O P (Index 1–10)
        private static readonly uint[] HotkeyVkCodes = { 0, 0x51, 0x57, 0x45, 0x52, 0x54, 0x5A, 0x55, 0x49, 0x4F, 0x50 };

        private readonly IntPtr _windowHandle;
        private HwndSource? _hwndSource;

        private int _registeredFrom;
        private int _registeredTo;

        /// <summary>Fired when a registered hotkey is pressed. The int argument is the 1-based step index.</summary>
        public event EventHandler<int>? HotkeyPressed;

        public GlobalHotkeyService(IntPtr windowHandle)
        {
            _windowHandle = windowHandle;
            _hwndSource = HwndSource.FromHwnd(_windowHandle);
            _hwndSource?.AddHook(WndProc);
        }

        /// <summary>Registers Ctrl+Alt+Q through Ctrl+Alt+P (indices 1–10).</summary>
        public bool RegisterHotkeys(int fromIndex = 1, int toIndex = 10)
        {
            _registeredFrom = fromIndex;
            _registeredTo   = toIndex;

            bool allOk = true;
            for (int i = fromIndex; i <= toIndex && i < HotkeyVkCodes.Length; i++)
            {
                if (!RegisterHotKey(_windowHandle, HOTKEY_ID_BASE + i, MOD_CONTROL | MOD_ALT, HotkeyVkCodes[i]))
                    allOk = false;
            }
            return allOk;
        }

        private void UnregisterAll()
        {
            for (int i = _registeredFrom; i <= _registeredTo; i++)
                UnregisterHotKey(_windowHandle, HOTKEY_ID_BASE + i);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            if (msg == WM_HOTKEY)
            {
                int index = wParam.ToInt32() - HOTKEY_ID_BASE;
                if (index >= 1 && index <= 10)
                {
                    HotkeyPressed?.Invoke(this, index);
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
