using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using CopagoAutomation.ViewModels;
using CopagoAutomation.Automation;

namespace CopagoAutomation.Windows
{
    public partial class CalibrationPromptWindow : Window, IDisposable
    {
        private readonly MainViewModel _viewModel;
        private HwndSource? _hwndSource;

        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;

        private const int HOTKEY_ID_DIGIT_0 = 1000;
        private const int HOTKEY_ID_NUMPAD_0 = 2000;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public string StepTitle
        {
            get => StepTitleText.Text;
            set => StepTitleText.Text = value;
        }

        public string HotkeyText
        {
            get => HotkeyTextBlock.Text;
            set => HotkeyTextBlock.Text = value;
        }

        public string InstructionText
        {
            get => InstructionTextBlock.Text;
            set => InstructionTextBlock.Text = value;
        }

        public bool WasConfirmed { get; private set; }

        public bool WasCancelled { get; private set; }

        public CalibrationPromptWindow(Window owner, MainViewModel viewModel)
        {
            InitializeComponent();
            Owner = owner;
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            // Ensure the window stays on top and doesn't appear in the taskbar
            Topmost = true;
            ShowInTaskbar = false;

            Loaded += CalibrationPromptWindow_Loaded;
            SourceInitialized += CalibrationPromptWindow_SourceInitialized;
            Closed += CalibrationPromptWindow_Closed;

            OkButton.Click += OkButton_Click;
            CancelButton.Click += CancelButton_Click;

            ResetWindowState();
        }

        private void CalibrationPromptWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Activate();
            _viewModel.ActiveCalibrationPrompt = this;
        }

        public void ResetWindowState()
        {
            WasConfirmed = false;
            WasCancelled = false;
            ResetCaptureState();
        }

        public void ResetCaptureState()
        {
            StatusTextBlock.Text = "Warte auf Tastenkombination...";
            OkButton.IsEnabled = false;
        }

        public void SetStepInfo(string stepTitle, string instructionText, string hotkeyText)
        {
            StepTitle = stepTitle;
            InstructionText = instructionText;
            HotkeyText = hotkeyText;

            ResetWindowState();
        }

        public void SetCapturedPosition(int x, int y)
        {
            StatusTextBlock.Text = $"Position erfasst: X={x}, Y={y}";
            OkButton.IsEnabled = true;
            OkButton.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.HasLastCapturedPosition)
            {
                MessageBox.Show(
                    "Bitte zuerst die angezeigte Tastenkombination drücken, um die Mausposition zu übernehmen.",
                    "Kalibrierung",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            WasConfirmed = true;
            WasCancelled = false;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            WasConfirmed = false;
            WasCancelled = true;
            DialogResult = false;
            Close();
        }

        public void Dispose()
        {
            UnregisterHotkeys();
            if (_hwndSource != null)
            {
                _hwndSource.RemoveHook(WndProc);
                _hwndSource = null;
            }
        }

        private void CalibrationPromptWindow_SourceInitialized(object? sender, EventArgs e)
        {
            _hwndSource = PresentationSource.FromVisual(this) as HwndSource;

            if (_hwndSource == null)
                return;

            _hwndSource.AddHook(WndProc);
            RegisterHotkeys();
        }

        private void CalibrationPromptWindow_Closed(object? sender, EventArgs e)
        {
            UnregisterHotkeys();
            if (_hwndSource != null)
            {
                _hwndSource.RemoveHook(WndProc);
                _hwndSource = null;
            }
            _viewModel.ActiveCalibrationPrompt = null;
        }

        private void RegisterHotkeys()
        {
            IntPtr handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero)
                return;

            for (int digit = 0; digit <= 9; digit++)
            {
                int normalId = HOTKEY_ID_DIGIT_0 + digit;
                int numpadId = HOTKEY_ID_NUMPAD_0 + digit;

                uint normalVk = (uint)(0x30 + digit);
                uint numpadVk = (uint)(0x60 + digit);
                    bool registeredNormal = RegisterHotKey(handle, normalId, MOD_CONTROL | MOD_ALT, normalVk);
                    bool registeredNumpad = RegisterHotKey(handle, numpadId, MOD_CONTROL | MOD_ALT, numpadVk);
                    System.Windows.MessageBox.Show($"Hotkey {digit} registriert: Normal={registeredNormal}, Numpad={registeredNumpad}", "Debug Hotkey Registration");
            }
        }

        private void UnregisterHotkeys()
        {
            IntPtr handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero)
                return;

            for (int digit = 0; digit <= 9; digit++)
            {
                int normalId = HOTKEY_ID_DIGIT_0 + digit;
                int numpadId = HOTKEY_ID_NUMPAD_0 + digit;
                UnregisterHotKey(handle, normalId);
                UnregisterHotKey(handle, numpadId);
            }
        }

        private IntPtr WndProc(
            IntPtr hwnd,
            int msg,
            IntPtr wParam,
            IntPtr lParam,
            ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int hotkeyId = wParam.ToInt32();
                HandleCalibrationHotkey(hotkeyId);
                handled = true;
            }

            return IntPtr.Zero;
        }

        private void HandleCalibrationHotkey(int hotkeyId)
        {
            if (_viewModel == null || _viewModel.CalibrationRunner == null)
                return;

                if (!TryGetDigitFromHotkeyId(hotkeyId, out int digit))
                    return;
                System.Windows.MessageBox.Show($"Hotkey {digit} in CalibrationPromptWindow empfangen.", "Debug Hotkey Reception");

            if (digit == 0)
            {
                if (_viewModel.TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundCopagoWindow))
                {
                    _viewModel.SetLastCapturedPosition(x, y, boundCopagoWindow);
                }
            }
            else if (digit >= 1 && digit <= 9)
            {
                _viewModel.SetLastCapturedPositionForDigit(digit);
            }
        }

        private bool TryGetDigitFromHotkeyId(int hotkeyId, out int digit)
        {
            digit = -1;

            if (hotkeyId >= HOTKEY_ID_DIGIT_0 && hotkeyId <= HOTKEY_ID_DIGIT_0 + 9)
            {
                digit = hotkeyId - HOTKEY_ID_DIGIT_0;
                return true;
            }

            if (hotkeyId >= HOTKEY_ID_NUMPAD_0 && hotkeyId <= HOTKEY_ID_NUMPAD_0 + 9)
            {
                digit = hotkeyId - HOTKEY_ID_NUMPAD_0;
                return true;
            }

            return false;
        }
    }
}
