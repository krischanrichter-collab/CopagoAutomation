using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using CopagoAutomation.Services;
using CopagoAutomation.ViewModels;
using CopagoAutomation.Automation;

namespace CopagoAutomation.Windows
{
    public partial class CalibrationPromptWindow : Window, IDisposable
    {
        private readonly MainViewModel _viewModel;
        private GlobalHotkeyService? _globalHotkeyService;

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
            Closed += CalibrationPromptWindow_Closed;

            OkButton.Click += OkButton_Click;
            CancelButton.Click += CancelButton_Click;

            ResetWindowState();
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
            _globalHotkeyService?.Dispose();
        }

        private void CalibrationPromptWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Activate();
            _viewModel.ActiveCalibrationPrompt = this;

            IntPtr handle = new WindowInteropHelper(this).Handle;
            if (handle != IntPtr.Zero)
            {
                _globalHotkeyService = new GlobalHotkeyService(handle);
                _globalHotkeyService.HotkeyPressed += GlobalHotkeyService_HotkeyPressed;
                if (!_globalHotkeyService.RegisterHotkey())
                {
                    MessageBox.Show("Fehler beim Registrieren des Hotkeys.", "Hotkey Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CalibrationPromptWindow_Closed(object? sender, EventArgs e)
        {
            _globalHotkeyService?.Dispose();
            _viewModel.ActiveCalibrationPrompt = null;
        }

        private void GlobalHotkeyService_HotkeyPressed(object? sender, EventArgs e)
        {
            if (_viewModel == null || _viewModel.CalibrationRunner == null)
                return;

            // Assuming the hotkey is always for capturing the current position (digit 0)
            if (_viewModel.TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundCopagoWindow))
            {
                _viewModel.SetLastCapturedPosition(x, y, boundCopagoWindow);
            }
        }
    }
}
