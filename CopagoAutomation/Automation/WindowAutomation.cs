using System;
using System.Runtime.InteropServices;
using System.Text;

using System.Windows.Forms;

namespace CopagoAutomation.Automation
{
	public class WindowAutomation
	{
		private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

		private const int DefaultWindowTextCapacity = 512;
		private const int SW_RESTORE = 9;

			[DllImport("user32.dll")]
			private static extern IntPtr GetDC(IntPtr hWnd);

			[DllImport("user32.dll")]
			private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

			[DllImport("gdi32.dll")]
			private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

			private const int LOGPIXELSX = 88;
			private const int LOGPIXELSY = 90;

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

			[DllImport("user32.dll")]
			private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

			[DllImport("user32.dll")]
			private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

			[DllImport("user32.dll")]
			private static extern void SetCursorPos(int x, int y);

			[DllImport("user32.dll")]
			private static extern void SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

			private const int INPUT_MOUSE = 0;
			private const int INPUT_KEYBOARD = 1;

			private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
			private const uint MOUSEEVENTF_LEFTUP = 0x0004;

			private const uint KEYEVENTF_KEYUP = 0x0002;
			private const uint KEYEVENTF_UNICODE = 0x0004;

			private const ushort VK_CONTROL = 0x11;
			private const ushort VK_A = 0x41;

			[StructLayout(LayoutKind.Sequential)]
			private struct INPUT
			{
				public int type;
				public InputUnion U;
			}

			[StructLayout(LayoutKind.Explicit)]
			private struct InputUnion
			{
				[FieldOffset(0)]
				public MOUSEINPUT mi;
				[FieldOffset(0)]
				public KEYBDINPUT ki;
				[FieldOffset(0)]
				public HARDWAREINPUT hi;
			}

			[StructLayout(LayoutKind.Sequential)]
			private struct MOUSEINPUT
			{
				public int dx;
				public int dy;
				public uint mouseData;
				public uint dwFlags;
				public uint time;
				public IntPtr dwExtraInfo;
			}

			[StructLayout(LayoutKind.Sequential)]
			private struct KEYBDINPUT
			{
				public ushort wVk;
				public ushort wScan;
				public uint dwFlags;
				public uint time;
				public IntPtr dwExtraInfo;
			}

			[StructLayout(LayoutKind.Sequential)]
			private struct HARDWAREINPUT
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
				// Sleep(20); // Sleep wird in Automation-Klasse gehandhabt
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
						mi = new MOUSEINPUT
						{
							dwFlags = MOUSEEVENTF_LEFTUP
						}
					}
				};

				var inputs = new[] { keyDown, keyUp };
				SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
			}
			{
				public int dx;
				public int dy;
				public uint mouseData;
				public uint dwFlags;
				public uint time;
				public IntPtr dwExtraInfo;
			}

			[StructLayout(LayoutKind.Sequential)]
			private struct KEYBDINPUT
			{
				public ushort wVk;
				public ushort wScan;
				public uint dwFlags;
				public uint time;
				public IntPtr dwExtraInfo;
			}

			[StructLayout(LayoutKind.Sequential)]
			private struct HARDWAREINPUT
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
				// Sleep(20); // Sleep wird in Automation-Klasse gehandhabt
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
				// Sleep(25); // Sleep wird in Automation-Klasse gehandhabt
				KeyPress(VK_A);
				// Sleep(25); // Sleep wird in Automation-Klasse gehandhabt
				KeyUp(VK_CONTROL);
				// Sleep(60); // Sleep wird in Automation-Klasse gehandhabt
			}

			public void TypeText(string text)
			{
				if (string.IsNullOrEmpty(text))
					return;

				foreach (char ch in text)
				{
					SendUnicodeChar(ch);
					// Sleep(DefaultTypingDelayMs); // Sleep wird in Automation-Klasse gehandhabt
				}
			}
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
						mi = new MOUSEINPUT
						{
							dwFlags = MOUSEEVENTF_LEFTUP
						}
					}
				};

				var inputs = new[] { keyDown, keyUp };
				SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
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

		public bool TryGetClientRect(IntPtr handle, out WindowRect clientRect)
			{
				clientRect = default;

				if (!IsValidHandle(handle))
					return false;

				if (!GetClientRect(handle, out RECT rect))
					return false;

				// Convert client coordinates to screen coordinates for the top-left corner
				POINT topLeft = new POINT { X = rect.Left, Y = rect.Top };
				ClientToScreen(handle, ref topLeft);

				clientRect = new WindowRect
				{
					Left = topLeft.X,
					Top = topLeft.Y,
					Right = topLeft.X + rect.Right - rect.Left,
					Bottom = topLeft.Y + rect.Bottom - rect.Top
				};

				return true;
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

			[StructLayout(LayoutKind.Sequential)]
			private struct POINT
			{
				public int X;
				public int Y;
			}
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