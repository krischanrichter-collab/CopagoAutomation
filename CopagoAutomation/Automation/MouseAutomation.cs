using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing; // For System.Drawing.Rectangle

namespace CopagoAutomation.Automation
{
    public class MouseAutomation
    {
        private readonly WindowAutomation _windowAutomation;

        private const int DefaultMoveDelayMs = 50;
        private const int DefaultDoubleClickDelayMs = 100;
        private const int DefaultWindowActivateDelayMs = 300;

        // These DllImports are now handled by WindowAutomation, but if MouseAutomation needs direct access,
        // they should be here or passed from WindowAutomation.
        // For now, we assume WindowAutomation handles the low-level mouse events.

        public MouseAutomation()
        {
            _windowAutomation = new WindowAutomation();
        }

        public MouseAutomation(WindowAutomation windowAutomation)
        {
            _windowAutomation = windowAutomation ?? throw new ArgumentNullException(nameof(windowAutomation));
        }

        public void MoveMouse(int x, int y)
        {
            _windowAutomation.SetCursorPosition(x, y);
        }

        public void LeftClick()
        {
            _windowAutomation.LeftClick();
        }

        public void RightClick()
        {
            _windowAutomation.RightClick();
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

            if (!_windowAutomation.TryBindWindowByHandle(_windowAutomation.GetActiveWindowHandle(), out var boundWindow))
                return false;

            return TryClickInWindowRect(boundWindow.ClientRect, relativeX, relativeY, delayAfterMs);
        }

        public bool RightClickInActiveWindow(int relativeX, int relativeY, int delayAfterMs = 0)
        {
            ValidateRelativeCoordinates(relativeX, relativeY);
            ValidateDelay(delayAfterMs, nameof(delayAfterMs));

            if (!_windowAutomation.TryBindWindowByHandle(_windowAutomation.GetActiveWindowHandle(), out var boundWindow))
                return false;

            return TryRightClickInWindowRect(boundWindow.ClientRect, relativeX, relativeY, delayAfterMs);
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

            if (!_windowAutomation.TryBindWindowByHandle(_windowAutomation.GetActiveWindowHandle(), out var boundWindow))
                return false;

            return TryDoubleClickInWindowRect(boundWindow.ClientRect, relativeX, relativeY, delayBetweenClicksMs, delayAfterMs);
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

            if (!_windowAutomation.TryBindWindowByTitleContains(titlePart, out var boundWindow))
                return false;

            if (!_windowAutomation.TryActivateBoundWindow(boundWindow))
                return false;

            SleepIfNeeded(activateDelayMs);

            return TryClickInWindowRect(boundWindow.ClientRect, relativeX, relativeY, delayAfterClickMs);
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

            if (!_windowAutomation.TryBindWindowByTitleContains(titlePart, out var boundWindow))
                return false;

            if (!_windowAutomation.TryActivateBoundWindow(boundWindow))
                return false;

            SleepIfNeeded(activateDelayMs);

            return TryRightClickInWindowRect(boundWindow.ClientRect, relativeX, relativeY, delayAfterRightClickMs);
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

            if (!_windowAutomation.TryBindWindowByTitleContains(titlePart, out var boundWindow))
                return false;

            if (!_windowAutomation.TryActivateBoundWindow(boundWindow))
                return false;

            SleepIfNeeded(activateDelayMs);

            return TryDoubleClickInWindowRect(boundWindow.ClientRect, relativeX, relativeY, delayBetweenClicksMs, delayAfterDoubleClickMs);
        }

        private bool TryClickInWindowRect(Rectangle clientRect, int relativeX, int relativeY, int delayAfterMs = 0)
        {
            ValidateRelativeCoordinates(relativeX, relativeY);
            ValidateDelay(delayAfterMs, nameof(delayAfterMs));

            if (!IsValidClientRect(clientRect))
                return false;

            if (relativeX >= clientRect.Width || relativeY >= clientRect.Height)
                return false;

            int absoluteX = clientRect.Left + relativeX;
            int absoluteY = clientRect.Top + relativeY;

            ClickAt(absoluteX, absoluteY, delayAfterMs);
            return true;
        }

        private bool TryRightClickInWindowRect(Rectangle clientRect, int relativeX, int relativeY, int delayAfterMs = 0)
        {
            ValidateRelativeCoordinates(relativeX, relativeY);
            ValidateDelay(delayAfterMs, nameof(delayAfterMs));

            if (!IsValidClientRect(clientRect))
                return false;

            if (relativeX >= clientRect.Width || relativeY >= clientRect.Height)
                return false;

            int absoluteX = clientRect.Left + relativeX;
            int absoluteY = clientRect.Top + relativeY;

            RightClickAt(absoluteX, absoluteY, delayAfterMs);
            return true;
        }

        private bool TryDoubleClickInWindowRect(
            Rectangle clientRect,
            int relativeX,
            int relativeY,
            int delayBetweenClicksMs = DefaultDoubleClickDelayMs,
            int delayAfterMs = 0)
        {
            ValidateRelativeCoordinates(relativeX, relativeY);
            ValidateDelay(delayBetweenClicksMs, nameof(delayBetweenClicksMs));
            ValidateDelay(delayAfterMs, nameof(delayAfterMs));

            if (!IsValidClientRect(clientRect))
                return false;

            if (relativeX >= clientRect.Width || relativeY >= clientRect.Height)
                return false;

            int absoluteX = clientRect.Left + relativeX;
            int absoluteY = clientRect.Top + relativeY;

            DoubleClickAt(absoluteX, absoluteY, delayBetweenClicksMs, delayAfterMs);
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

        private static bool IsValidClientRect(Rectangle rect)
        {
            return rect.Width > 0 && rect.Height > 0;
        }
    }
}
