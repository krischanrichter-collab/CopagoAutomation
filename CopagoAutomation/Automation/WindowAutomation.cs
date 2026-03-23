using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CopagoAutomation.Automation
{
	public class WindowAutomation
	{
		private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

		private const int DefaultWindowTextCapacity = 512;
		private const int SW_RESTORE = 9;

		[DllImport("user32.dll")]
		private static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetWindowTextLength(IntPtr hWnd);

		[DllImport("user32.dll")]
		private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

		[DllImport("user32.dll")]
		private static extern bool SetForegroundWindow(IntPtr hWnd);

		[DllImport("user32.dll")]
		private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

		[DllImport("user32.dll")]
		private static extern bool IsWindowVisible(IntPtr hWnd);

		[DllImport("user32.dll")]
		private static extern bool IsWindow(IntPtr hWnd);

		[DllImport("user32.dll")]
		private static extern bool IsIconic(IntPtr hWnd);

		[DllImport("user32.dll")]
		private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

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

		public bool ActivateActiveWindow()
		{
			IntPtr handle = GetActiveWindowHandle();
			if (!IsValidHandle(handle))
				return false;

			return TryActivateWindow(handle);
		}

		public bool TryFindWindowByTitleContains(
			string titlePart,
			out IntPtr windowHandle,
			StringComparison comparison = StringComparison.OrdinalIgnoreCase)
		{
			windowHandle = IntPtr.Zero;

			if (string.IsNullOrWhiteSpace(titlePart))
				return false;

			IntPtr foundHandle = IntPtr.Zero;

			EnumWindows((hWnd, lParam) =>
			{
				if (!IsValidHandle(hWnd))
					return true;

				if (!IsWindowVisible(hWnd))
					return true;

				string title = GetWindowTitle(hWnd);
				if (string.IsNullOrWhiteSpace(title))
					return true;

				if (title.IndexOf(titlePart, comparison) >= 0)
				{
					foundHandle = hWnd;
					return false;
				}

				return true;
			}, IntPtr.Zero);

			windowHandle = foundHandle;
			return IsValidHandle(windowHandle);
		}

		public bool TryBindWindowByTitleContains(
			string titlePart,
			out BoundWindowInfo boundWindow,
			StringComparison comparison = StringComparison.OrdinalIgnoreCase)
		{
			boundWindow = default;

			if (!TryFindWindowByTitleContains(titlePart, out IntPtr handle, comparison))
				return false;

			string title = GetWindowTitle(handle);
			if (string.IsNullOrWhiteSpace(title))
				return false;

			boundWindow = new BoundWindowInfo(handle, title);
			return true;
		}

		public bool TryActivateWindowByTitleContains(
			string titlePart,
			StringComparison comparison = StringComparison.OrdinalIgnoreCase)
		{
			if (!TryFindWindowByTitleContains(titlePart, out var handle, comparison))
				return false;

			return TryActivateWindow(handle);
		}

		public bool IsWindowValid(IntPtr handle)
		{
			return IsValidHandle(handle);
		}

		public bool IsWindowValid(BoundWindowInfo boundWindow)
		{
			return boundWindow.HasHandle && IsValidHandle(boundWindow.Handle);
		}

		public bool IsBoundWindowActive(BoundWindowInfo boundWindow)
		{
			if (!IsWindowValid(boundWindow))
				return false;

			return GetForegroundWindow() == boundWindow.Handle;
		}

		public bool IsBoundWindowTitleAvailable(BoundWindowInfo boundWindow)
		{
			if (!IsWindowValid(boundWindow))
				return false;

			string currentTitle = GetWindowTitle(boundWindow.Handle);
			return !string.IsNullOrWhiteSpace(currentTitle);
		}

		public string GetBoundWindowTitle(BoundWindowInfo boundWindow)
		{
			if (!IsWindowValid(boundWindow))
				return string.Empty;

			return GetWindowTitle(boundWindow.Handle);
		}

		public bool TryGetWindowRect(IntPtr handle, out WindowRect windowRect)
		{
			windowRect = default;

			if (!IsValidHandle(handle))
				return false;

			if (!GetWindowRect(handle, out RECT rect))
				return false;

			var result = new WindowRect
			{
				Left = rect.Left,
				Top = rect.Top,
				Right = rect.Right,
				Bottom = rect.Bottom
			};

			if (!IsValidWindowRect(result))
				return false;

			windowRect = result;
			return true;
		}

		public bool TryGetWindowRect(BoundWindowInfo boundWindow, out WindowRect windowRect)
		{
			windowRect = default;

			if (!IsWindowValid(boundWindow))
				return false;

			return TryGetWindowRect(boundWindow.Handle, out windowRect);
		}

		public bool TryGetActiveWindowRect(out WindowRect windowRect)
		{
			IntPtr handle = GetActiveWindowHandle();
			return TryGetWindowRect(handle, out windowRect);
		}

		private static bool IsValidHandle(IntPtr handle)
		{
			return handle != IntPtr.Zero && IsWindow(handle);
		}

		private static bool IsValidWindowRect(WindowRect rect)
		{
			return rect.Right > rect.Left && rect.Bottom > rect.Top;
		}

		private string GetWindowTitle(IntPtr handle)
		{
			if (!IsValidHandle(handle))
				return string.Empty;

			int textLength = GetWindowTextLength(handle);
			int capacity = Math.Max(textLength + 1, DefaultWindowTextCapacity);

			var builder = new StringBuilder(capacity);
			int copiedLength = GetWindowText(handle, builder, builder.Capacity);

			if (copiedLength <= 0)
				return string.Empty;

			return builder.ToString();
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct RECT
		{
			public int Left;
			public int Top;
			public int Right;
			public int Bottom;
		}
	}

	public struct WindowRect
	{
		public int Left { get; set; }
		public int Top { get; set; }
		public int Right { get; set; }
		public int Bottom { get; set; }

		public int Width => Right - Left;
		public int Height => Bottom - Top;
	}
}