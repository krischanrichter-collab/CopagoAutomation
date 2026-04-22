using System;
using System.Runtime.InteropServices;

namespace CopagoAutomation.Services
{
    public sealed class EscapeWatcher : IDisposable
    {
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(string? lpModuleName);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN     = 0x0100;
        private const int VK_ESCAPE      = 0x1B;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private readonly LowLevelKeyboardProc _proc;
        private IntPtr _hookId = IntPtr.Zero;

        public event Action? EscapePressed;

        public EscapeWatcher()
        {
            _proc = HookCallback;
        }

        public void Start()
        {
            if (_hookId != IntPtr.Zero) return;
            var module = System.Diagnostics.Process.GetCurrentProcess().MainModule;
            _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(module?.ModuleName), 0);
        }

        public void Stop()
        {
            if (_hookId == IntPtr.Zero) return;
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                if (Marshal.ReadInt32(lParam) == VK_ESCAPE)
                    EscapePressed?.Invoke();
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        public void Dispose() => Stop();
    }
}
