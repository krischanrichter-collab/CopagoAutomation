using System;
using System.Windows;
using System.Windows.Interop;
using CopagoAutomation.Automation;
using CopagoAutomation.Services;
using CopagoAutomation.ViewModels;

namespace CopagoAutomation.Windows
{
    public partial class CalibrationPromptWindow : Window, IDisposable
    {
        private readonly MainViewModel _viewModel;
        private GlobalHotkeyService? _globalHotkeyService;

        public bool WasConfirmed { get; private set; }

        public CalibrationPromptWindow(Window owner, MainViewModel viewModel)
        {
            InitializeComponent();
            Owner = owner;
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));

            Topmost = true;
            ShowInTaskbar = false;

            Loaded  += CalibrationPromptWindow_Loaded;
            Closed  += CalibrationPromptWindow_Closed;

            OkButton.Click     += OkButton_Click;
            CancelButton.Click += CancelButton_Click;
        }

        private void CalibrationPromptWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Activate();

            // Register Ctrl+Alt+1 through Ctrl+Alt+9 for this window
            IntPtr handle = new WindowInteropHelper(this).Handle;
            if (handle != IntPtr.Zero)
            {
                _globalHotkeyService = new GlobalHotkeyService(handle);
                _globalHotkeyService.HotkeyPressed += GlobalHotkeyService_HotkeyPressed;
                if (!_globalHotkeyService.RegisterHotkeys(1, 9))
                {
                    MessageBox.Show(
                        "Fehler beim Registrieren der Hotkeys. Möglicherweise sind Strg+Alt+1–9 bereits belegt.",
                        "Hotkey Fehler",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }

            // Show the first step
            ShowCurrentStep();
        }

        private void CalibrationPromptWindow_Closed(object? sender, EventArgs e)
        {
            _globalHotkeyService?.Dispose();
        }

        /// <summary>
        /// Fired when Ctrl+Alt+N is pressed. Captures the cursor position only if N
        /// matches the hotkey digit of the current calibration step.
        /// </summary>
        private void GlobalHotkeyService_HotkeyPressed(object? sender, int digit)
        {
            if (_viewModel.CalibrationRunner?.CurrentStep == null)
                return;

            if (digit != _viewModel.CalibrationRunner.CurrentStep.HotkeyDigit)
                return;

            if (_viewModel.TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundWindow))
            {
                _viewModel.SetLastCapturedPosition(x, y, boundWindow);
                SetCapturedStatus(x, y);
            }
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

            // Save the captured point; CalibrationRunner advances to the next step internally
            _viewModel.SaveCurrentCalibrationPoint(_viewModel.LastBoundWindow);

            if (_viewModel.IsCalibrationRunning)
            {
                // More steps remain — update the UI for the next step
                ShowCurrentStep();
                Activate(); // Keep calibration dialog in foreground, prevent main window from gaining focus
            }
            else
            {
                // All steps completed
                WasConfirmed = true;
                DialogResult = true;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.CancelCalibration();
            WasConfirmed = false;
            DialogResult = false;
            Close();
        }

        /// <summary>Updates the UI to reflect the ViewModel's current calibration step.</summary>
        private void ShowCurrentStep()
        {
            var step = _viewModel.CurrentCalibrationStep;
            if (step == null)
                return;

            StepTitleText.Text       = step.Title;
            HotkeyTextBlock.Text     = step.HotkeyText;
            InstructionTextBlock.Text = step.InstructionText;

            StatusTextBlock.Text      = "Warte auf Tastenkombination...";
            OkButton.IsEnabled        = false;
        }

        private void SetCapturedStatus(int x, int y)
        {
            StatusTextBlock.Text = $"Position erfasst: X={x}, Y={y}";
            OkButton.IsEnabled   = true;
            Activate(); // Fenster auf OS-Ebene in den Vordergrund bringen, damit Enter den OK-Button trifft
            OkButton.Focus();
        }

        public void Dispose()
        {
            _globalHotkeyService?.Dispose();
        }
    }
}
