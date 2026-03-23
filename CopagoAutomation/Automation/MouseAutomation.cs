using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace CopagoAutomation.Automation
{
	public class MouseAutomation
	{
		private readonly WindowAutomation _windowAutomation;

		private const int DefaultMoveDelayMs = 50;
		private const int DefaultDoubleClickDelayMs = 100;
		private const int DefaultWindowActivateDelayMs = 300;

		private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
		private const uint MOUSEEVENTF_LEFTUP = 0x0004;
		private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
		private const uint MOUSEEVENTF_RIGHTUP = 0x0010;

		[DllImport("user32.dll")]
		private static extern bool SetCursorPos(int X, int Y);

		[DllImport("user32.dll")]
		private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

		public MouseAutomation()
		{
			_windowAutomation = new WindowAutomation();
		}

		public void MoveMouse(int x, int y)
		{
			if (!SetCursorPos(x, y))
				throw new InvalidOperationException($"Maus konnte nicht auf Position ({x}, {y}) bewegt werden.");
		}

		public void LeftClick()
		{
			mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
			mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
		}

		public void RightClick()
		{
			mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
			mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
		}

		public void DoubleLeftClick(int delayBetweenClicksMs = DefaultDoubleClickDelayMs)
		{
			ValidateDelay(delayBetweenClicksMs, nameof(delayBetweenClicksMs));

			LeftClick();
			SleepIfNeeded(delayBetweenClicksMs);
			LeftClick();
		}

		public void ClickAt(int x, int y, int delayAfterMs = 0)
		{
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			MoveMouse(x, y);
			SleepIfNeeded(DefaultMoveDelayMs);

			LeftClick();
			SleepIfNeeded(delayAfterMs);
		}

		public void RightClickAt(int x, int y, int delayAfterMs = 0)
		{
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			MoveMouse(x, y);
			SleepIfNeeded(DefaultMoveDelayMs);

			RightClick();
			SleepIfNeeded(delayAfterMs);
		}

		public void DoubleClickAt(
			int x,
			int y,
			int delayBetweenClicksMs = DefaultDoubleClickDelayMs,
			int delayAfterMs = 0)
		{
			ValidateDelay(delayBetweenClicksMs, nameof(delayBetweenClicksMs));
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			MoveMouse(x, y);
			SleepIfNeeded(DefaultMoveDelayMs);

			DoubleLeftClick(delayBetweenClicksMs);
			SleepIfNeeded(delayAfterMs);
		}

		public bool ClickInActiveWindow(int relativeX, int relativeY, int delayAfterMs = 0)
		{
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!_windowAutomation.TryGetActiveWindowRect(out var rect))
				return false;

			return TryClickInWindowRect(rect, relativeX, relativeY, delayAfterMs);
		}

		public bool RightClickInActiveWindow(int relativeX, int relativeY, int delayAfterMs = 0)
		{
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!_windowAutomation.TryGetActiveWindowRect(out var rect))
				return false;

			return TryRightClickInWindowRect(rect, relativeX, relativeY, delayAfterMs);
		}

		public bool DoubleClickInActiveWindow(
			int relativeX,
			int relativeY,
			int delayBetweenClicksMs = DefaultDoubleClickDelayMs,
			int delayAfterMs = 0)
		{
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayBetweenClicksMs, nameof(delayBetweenClicksMs));
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!_windowAutomation.TryGetActiveWindowRect(out var rect))
				return false;

			return TryDoubleClickInWindowRect(rect, relativeX, relativeY, delayBetweenClicksMs, delayAfterMs);
		}

		public bool ClickInActiveWindowIfTitleContains(
			string expectedTitlePart,
			int relativeX,
			int relativeY,
			int delayAfterMs = 0)
		{
			ValidateTitlePart(expectedTitlePart);
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!_windowAutomation.ActiveWindowTitleContains(expectedTitlePart))
				return false;

			return ClickInActiveWindow(relativeX, relativeY, delayAfterMs);
		}

		public bool ActivateWindowAndClick(
			string titlePart,
			int relativeX,
			int relativeY,
			int activateDelayMs = DefaultWindowActivateDelayMs,
			int delayAfterClickMs = 0)
		{
			ValidateTitlePart(titlePart);
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(activateDelayMs, nameof(activateDelayMs));
			ValidateDelay(delayAfterClickMs, nameof(delayAfterClickMs));

			if (!_windowAutomation.TryFindWindowByTitleContains(titlePart, out var handle))
				return false;

			if (!_windowAutomation.TryActivateWindow(handle))
				return false;

			SleepIfNeeded(activateDelayMs);

			if (!_windowAutomation.TryGetWindowRect(handle, out var rect))
				return false;

			return TryClickInWindowRect(rect, relativeX, relativeY, delayAfterClickMs);
		}

		public bool ActivateWindowAndRightClick(
			string titlePart,
			int relativeX,
			int relativeY,
			int activateDelayMs = DefaultWindowActivateDelayMs,
			int delayAfterRightClickMs = 0)
		{
			ValidateTitlePart(titlePart);
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(activateDelayMs, nameof(activateDelayMs));
			ValidateDelay(delayAfterRightClickMs, nameof(delayAfterRightClickMs));

			if (!_windowAutomation.TryFindWindowByTitleContains(titlePart, out var handle))
				return false;

			if (!_windowAutomation.TryActivateWindow(handle))
				return false;

			SleepIfNeeded(activateDelayMs);

			if (!_windowAutomation.TryGetWindowRect(handle, out var rect))
				return false;

			return TryRightClickInWindowRect(rect, relativeX, relativeY, delayAfterRightClickMs);
		}

		public bool ActivateWindowAndDoubleClick(
			string titlePart,
			int relativeX,
			int relativeY,
			int activateDelayMs = DefaultWindowActivateDelayMs,
			int delayBetweenClicksMs = DefaultDoubleClickDelayMs,
			int delayAfterDoubleClickMs = 0)
		{
			ValidateTitlePart(titlePart);
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(activateDelayMs, nameof(activateDelayMs));
			ValidateDelay(delayBetweenClicksMs, nameof(delayBetweenClicksMs));
			ValidateDelay(delayAfterDoubleClickMs, nameof(delayAfterDoubleClickMs));

			if (!_windowAutomation.TryFindWindowByTitleContains(titlePart, out var handle))
				return false;

			if (!_windowAutomation.TryActivateWindow(handle))
				return false;

			SleepIfNeeded(activateDelayMs);

			if (!_windowAutomation.TryGetWindowRect(handle, out var rect))
				return false;

			return TryDoubleClickInWindowRect(rect, relativeX, relativeY, delayBetweenClicksMs, delayAfterDoubleClickMs);
		}

		private bool TryClickInWindowRect(WindowRect rect, int relativeX, int relativeY, int delayAfterMs = 0)
		{
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!TryGetAbsoluteCoordinates(rect, relativeX, relativeY, out int absoluteX, out int absoluteY))
				return false;

			ClickAt(absoluteX, absoluteY, delayAfterMs);
			return true;
		}

		private bool TryRightClickInWindowRect(WindowRect rect, int relativeX, int relativeY, int delayAfterMs = 0)
		{
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!TryGetAbsoluteCoordinates(rect, relativeX, relativeY, out int absoluteX, out int absoluteY))
				return false;

			RightClickAt(absoluteX, absoluteY, delayAfterMs);
			return true;
		}

		private bool TryDoubleClickInWindowRect(
			WindowRect rect,
			int relativeX,
			int relativeY,
			int delayBetweenClicksMs = DefaultDoubleClickDelayMs,
			int delayAfterMs = 0)
		{
			ValidateRelativeCoordinates(relativeX, relativeY);
			ValidateDelay(delayBetweenClicksMs, nameof(delayBetweenClicksMs));
			ValidateDelay(delayAfterMs, nameof(delayAfterMs));

			if (!TryGetAbsoluteCoordinates(rect, relativeX, relativeY, out int absoluteX, out int absoluteY))
				return false;

			DoubleClickAt(absoluteX, absoluteY, delayBetweenClicksMs, delayAfterMs);
			return true;
		}

		private bool TryGetAbsoluteCoordinates(WindowRect rect, int relativeX, int relativeY, out int absoluteX, out int absoluteY)
		{
			absoluteX = 0;
			absoluteY = 0;

			if (!IsValidWindowRect(rect))
				return false;

			if (relativeX >= rect.Width || relativeY >= rect.Height)
				return false;

			absoluteX = rect.Left + relativeX;
			absoluteY = rect.Top + relativeY;
			return true;
		}

		private void SleepIfNeeded(int delayMs)
		{
			if (delayMs > 0)
				Thread.Sleep(delayMs);
		}

		private static void ValidateTitlePart(string titlePart)
		{
			if (string.IsNullOrWhiteSpace(titlePart))
				throw new ArgumentException("titlePart darf nicht leer sein.", nameof(titlePart));
		}

		private static void ValidateDelay(int delayMs, string parameterName)
		{
			if (delayMs < 0)
				throw new ArgumentOutOfRangeException(parameterName, "Delay darf nicht negativ sein.");
		}

		private static void ValidateRelativeCoordinates(int relativeX, int relativeY)
		{
			if (relativeX < 0)
				throw new ArgumentOutOfRangeException(nameof(relativeX), "relativeX darf nicht negativ sein.");

			if (relativeY < 0)
				throw new ArgumentOutOfRangeException(nameof(relativeY), "relativeY darf nicht negativ sein.");
		}

		private static bool IsValidWindowRect(WindowRect rect)
		{
			return rect.Right > rect.Left && rect.Bottom > rect.Top;
		}
	}
}