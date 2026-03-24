using System;
using CopagoAutomation.Automation;
using System.Drawing; // For System.Drawing.Point
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Win32;
using CopagoAutomation.Calibration;
using CopagoAutomation.Data;
using CopagoAutomation.Models;
using CopagoAutomation.Services;
using CopagoAutomation.ViewModels;
using CopagoAutomation.Windows;
using System.Runtime.InteropServices; // Keep this for DllImport of RegisterHotKey/UnregisterHotKey if needed

namespace CopagoAutomation
{
    public partial class MainWindow : Window
    {
        private readonly SettingsStore _store;
        private readonly CalibrationStorage _calibrationStorage;

        private AutomationService? _automationService;
        private WindowAutomation? _windowAutomation;
        private CalibrationData? _calibrationData;
        private CalibrationService? _calibrationService;
        private MainViewModel? _mainViewModel;
        private AppSettings? _settings;

        private HwndSource? _hwndSource;
        private CalibrationPromptWindow? _activeCalibrationPrompt;

        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;

        private const int HOTKEY_ID_DIGIT_0 = 1000;
        private const int HOTKEY_ID_NUMPAD_0 = 2000;

        private const uint GA_ROOT = 2;

        public MainWindow()
        {
            InitializeComponent();

            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CopagoAutomation",
                "settings.json");

            _store = new SettingsStore(path);

            var calibrationPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CopagoAutomation",
                "calibration.json");

            _calibrationStorage = new CalibrationStorage(calibrationPath);

            AbcModeLaptop.Checked += AbcMode_Checked;
            AbcModeDocking.Checked += AbcMode_Checked;

            XModeLaptop.Checked += XMode_Checked;
            XModeDocking.Checked += XMode_Checked;

            AbcSammelordner.LostFocus += AbcStorageFields_LostFocus;
            XSammelordner.LostFocus += XStorageFields_LostFocus;

            LoadPosEntries();
            LoadSettings();

            Loaded += MainWindow_Loaded;
            SourceInitialized += MainWindow_SourceInitialized;
            Closed += MainWindow_Closed;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                _calibrationData = await _calibrationStorage.LoadAsync();
                _windowAutomation = new WindowAutomation();
                _calibrationService = new CalibrationService(_calibrationData, _windowAutomation);

                if (_settings == null)
                    _settings = new AppSettings();

                _automationService = new AutomationService(_calibrationService, new PathResolver(_settings));
                _mainViewModel = new MainViewModel(_calibrationService);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Fehler beim Laden der Kalibrierungsdaten: {ex.Message}",
                    "Kalibrierung",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void MainWindow_SourceInitialized(object? sender, EventArgs e)
        {
            _hwndSource = PresentationSource.FromVisual(this) as HwndSource;

            if (_hwndSource == null)
                return;

            _hwndSource.AddHook(WndProc);
            RegisterCalibrationHotkeys();
        }

        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            UnregisterHotkeys();

            if (_hwndSource != null)
            {
                _hwndSource.RemoveHook(WndProc);
                _hwndSource = null;
            }
        }

        private async void LoadSettings()
        {
            try
            {
                _settings = await _store.LoadAsync();
            }
            catch
            {
                _settings = new AppSettings();
            }

            ApplyModesToUi();
            ApplyStorageSettingsToUi();
            UpdateAbcSaveModeUi();
            ApplyXStorageSettingsToUi();
            UpdateXSaveModeUi();
        }

        private void ApplyXStorageSettingsToUi()
        {
            if (_settings == null || XSaveModeSemco == null || XSaveModeAlt == null || XSammelordner == null)
                return;

            if (_settings.XSaveMode == SaveMode.SemcoUpload)
                XSaveModeSemco.IsChecked = true;
            else
                XSaveModeAlt.IsChecked = true;

            XSammelordner.Text = _settings.XSammelordnerPath;

            UpdateXSaveModeUi();
        }

        private void LoadPosEntries()
        {
            var allPos = PosRepository.GetAll();

            AbcPosList.Items.Clear();
            XPosList.Items.Clear();

            foreach (string pos in allPos)
            {
                AbcPosList.Items.Add(pos);
                XPosList.Items.Add(pos);
            }
        }

        private async Task SaveSettingsAsync()
        {
            if (_settings != null)
                await _store.SaveAsync(_settings);
        }

        private async Task SaveCalibrationDataAsync()
        {
            if (_calibrationData != null)
                await _calibrationStorage.SaveAsync(_calibrationData);
        }

        private static string GetCalibrationModeName(MachineMode mode)
        {
            return mode == MachineMode.Laptop ? "laptop" : "dock";
        }

        private bool TryGetCurrentClientCursorPosition(out int x, out int y, out WindowAutomation.BoundWindowInfo? boundCopagoWindow)
        {
            x = 0;
            y = 0;
            boundCopagoWindow = null;

            if (_windowAutomation == null) return false;

            // Get current cursor position
            System.Drawing.Point screenPoint = _windowAutomation.GetCursorScreenPosition();

            // Try to get the window under the cursor and its root
            IntPtr childWindow = WindowAutomation.WindowFromPoint(new WindowAutomation.POINT { X = screenPoint.X, Y = screenPoint.Y });
            if (childWindow == IntPtr.Zero) return false;

            IntPtr rootWindow = WindowAutomation.GetAncestor(childWindow, GA_ROOT);
            if (rootWindow == IntPtr.Zero) return false;

            // Convert screen coordinates to client coordinates of the root window
            System.Drawing.Point clientPoint = screenPoint;
            WindowAutomation.POINT clientPointWin32 = new WindowAutomation.POINT { X = clientPoint.X, Y = clientPoint.Y };
            if (!WindowAutomation.ScreenToClient(rootWindow, ref clientPointWin32)) return false;
            clientPoint = new System.Drawing.Point(clientPointWin32.X, clientPointWin32.Y);

            // If the root window is Copago, bind it
            if (_windowAutomation.TryBindWindowByHandle(rootWindow, out var copagoWindow))
            {
                boundCopagoWindow = copagoWindow;
            }

            x = clientPoint.X;
            y = clientPoint.Y;
            return true;
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

        private void RegisterCalibrationHotkeys()
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

                WindowAutomation.RegisterHotKey(handle, normalId, MOD_CONTROL | MOD_ALT, normalVk);
                WindowAutomation.RegisterHotKey(handle, numpadId, MOD_CONTROL | MOD_ALT, numpadVk);
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

                WindowAutomation.UnregisterHotKey(handle, normalId);
                WindowAutomation.UnregisterHotKey(handle, numpadId);
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
            if (_mainViewModel == null || _mainViewModel.CalibrationRunner == null)
                return;

            if (!TryGetDigitFromHotkeyId(hotkeyId, out int digit))
                return;

            if (digit == 0)
            {
                if (TryGetCurrentClientCursorPosition(out int x, out int y, out WindowAutomation.BoundWindowInfo? boundWindow))
                {
                    _mainViewModel.SetLastCapturedPosition(x, y, boundWindow);
                }
            }
        }

        private void ApplyModesToUi()
        {
            if (_settings == null)
                return;

            if (_settings.AbcMode == MachineMode.Laptop)
                AbcModeLaptop.IsChecked = true;
            else
                AbcModeDocking.IsChecked = true;

            if (_settings.XMode == MachineMode.Laptop)
                XModeLaptop.IsChecked = true;
            else
                XModeDocking.IsChecked = true;
        }

        private void ApplyStorageSettingsToUi()
        {
            if (_settings == null) return;

            if (AbcSaveModeSemco != null && AbcSaveModeAlt != null)
            {
                if (_settings.AbcSaveMode == SaveMode.SemcoUpload)
                    AbcSaveModeSemco.IsChecked = true;
                else
                    AbcSaveModeAlt.IsChecked = true;
            }

            if (XSaveModeSemco != null && XSaveModeAlt != null)
            {
                if (_settings.XSaveMode == SaveMode.SemcoUpload)
                    XSaveModeSemco.IsChecked = true;
                else
                    XSaveModeAlt.IsChecked = true;
            }

            if (AbcSammelordner != null)
                AbcSammelordner.Text = _settings.AbcSammelordnerPath ?? string.Empty;
            if (XSammelordner != null)
                XSammelordner.Text = _settings.XSammelordnerPath ?? string.Empty;
        }

        private void UpdateAbcSaveModeUi()
        {
            if (_settings == null) return;

            bool isAlternativeMode = _settings.AbcSaveMode == SaveMode.Alternative;

            if (AbcSammelordner != null)
                AbcSammelordner.IsEnabled = isAlternativeMode;
            if (AbcBrowseSammelordner != null)
                AbcBrowseSammelordner.IsEnabled = isAlternativeMode;
        }

        private void UpdateXSaveModeUi()
        {
            if (_settings == null) return;

            bool isAlternativeMode = _settings.XSaveMode == SaveMode.Alternative;

            if (XSammelordner != null)
                XSammelordner.IsEnabled = isAlternativeMode;
            if (XBrowseSammelordner != null)
                XBrowseSammelordner.IsEnabled = isAlternativeMode;
        }

        private async void AbcMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (AbcModeLaptop != null && AbcModeLaptop.IsChecked == true)
                _settings.AbcMode = MachineMode.Laptop;
            else if (AbcModeDocking != null && AbcModeDocking.IsChecked == true)
                _settings.AbcMode = MachineMode.Docking;

            await SaveSettingsAsync();
        }

        private async void XMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (XModeLaptop != null && XModeLaptop.IsChecked == true)
                _settings.XMode = MachineMode.Laptop;
            else if (XModeDocking != null && XModeDocking.IsChecked == true)
                _settings.XMode = MachineMode.Docking;

            await SaveSettingsAsync();
        }

        private async void AbcSaveModeChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (AbcSaveModeSemco != null && AbcSaveModeSemco.IsChecked == true)
                _settings.AbcSaveMode = SaveMode.SemcoUpload;
            else if (AbcSaveModeAlt != null && AbcSaveModeAlt.IsChecked == true)
                _settings.AbcSaveMode = SaveMode.Alternative;

            UpdateAbcSaveModeUi();
            await SaveSettingsAsync();
        }

        private async void XSaveModeChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (XSaveModeSemco != null && XSaveModeSemco.IsChecked == true)
                _settings.XSaveMode = SaveMode.SemcoUpload;
            else if (XSaveModeAlt != null && XSaveModeAlt.IsChecked == true)
                _settings.XSaveMode = SaveMode.Alternative;

            UpdateXSaveModeUi();
            await SaveSettingsAsync();
        }

        private async void AbcBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            var dialog = new OpenFolderDialog
            {
                Title = "Sammelordner für ABC-Berichte auswählen",
                InitialDirectory = _settings.AbcSammelordnerPath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dialog.ShowDialog() == true)
            {
                _settings.AbcSammelordnerPath = dialog.FolderName;
                if (AbcSammelordner != null)
                    AbcSammelordner.Text = dialog.FolderName;
                await SaveSettingsAsync();
            }
        }

        private async void XBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            var dialog = new OpenFolderDialog
            {
                Title = "Sammelordner für X-Berichte auswählen",
                InitialDirectory = _settings.XSammelordnerPath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dialog.ShowDialog() == true)
            {
                _settings.XSammelordnerPath = dialog.FolderName;
                if (XSammelordner != null)
                    XSammelordner.Text = dialog.FolderName;
                await SaveSettingsAsync();
            }
        }

        private async void AbcStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            _settings.AbcSammelordnerPath = AbcSammelordner.Text;
            await SaveSettingsAsync();
        }

        private async void XStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            _settings.XSammelordnerPath = XSammelordner.Text;
            await SaveSettingsAsync();
        }

        private async void AbcStart_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null || _automationService == null || _settings == null) return;

            string? selectedPos = AbcPosList.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selectedPos))
            {
                MessageBox.Show("Bitte wählen Sie eine POS aus.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!AbcDateRuleSelector.DateFrom.HasValue || !AbcDateRuleSelector.DateTo.HasValue)
            {
                MessageBox.Show("Bitte wählen Sie einen Datumsbereich aus.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            LogAbc($"Starte ABC-Automatisierung für POS {selectedPos} von {AbcDateRuleSelector.DateFrom.Value:d} bis {AbcDateRuleSelector.DateTo.Value:d}...");

            var request = new AbcStartRequest
            {
                PosId = selectedPos,
                DateFrom = AbcDateRuleSelector.DateFrom.Value,
                DateTo = AbcDateRuleSelector.DateTo.Value,
                SaveMode = _settings.AbcSaveMode,
                SammelordnerPath = _settings.AbcSammelordnerPath
            };

            try
            {
                await _automationService.StartAbcAutomationAsync(request);
                LogAbc("ABC-Automatisierung abgeschlossen.");
            }
            catch (Exception ex)
            {
                LogAbc($"Fehler bei ABC-Automatisierung: {ex.Message}");
                MessageBox.Show($"Fehler bei ABC-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void XStart_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null || _automationService == null || _settings == null) return;

            string? selectedPos = XPosList.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selectedPos))
            {
                MessageBox.Show("Bitte wählen Sie eine POS aus.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!XDateRuleSelector.DateFrom.HasValue || !XDateRuleSelector.DateTo.HasValue)
            {
                MessageBox.Show("Bitte wählen Sie einen Datumsbereich aus.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            LogX($"Starte X-Automatisierung für POS {selectedPos} von {XDateRuleSelector.DateFrom.Value:d} bis {XDateRuleSelector.DateTo.Value:d}...");

            var request = new XStartRequest
            {
                PosId = selectedPos,
                DateFrom = XDateRuleSelector.DateFrom.Value,
                DateTo = XDateRuleSelector.DateTo.Value,
                SaveMode = _settings.XSaveMode,
                SammelordnerPath = _settings.XSammelordnerPath
            };

            try
            {
                await _automationService.StartXAutomationAsync(request);
                LogX("X-Automatisierung abgeschlossen.");
            }
            catch (Exception ex)
            {
                LogX($"Fehler bei X-Automatisierung: {ex.Message}");
                MessageBox.Show($"Fehler bei X-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LogAbc(string message)
        {
            AbcLogBox.AppendText($"{DateTime.Now:HH:mm:ss} {message}\n");
            AbcLogBox.ScrollToEnd();
        }

        private void LogX(string message)
        {
            XLogBox.AppendText($"{DateTime.Now:HH:mm:ss} {message}\n");
            XLogBox.ScrollToEnd();
        }

        private void AbcCalibrate_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;

            string modeName = GetCalibrationModeName(_settings?.AbcMode ?? MachineMode.Laptop);
            _mainViewModel.StartCalibration(modeName, "AbcReport");

            _activeCalibrationPrompt = new CalibrationPromptWindow(this, _mainViewModel);
            _activeCalibrationPrompt.Owner = this;
            _activeCalibrationPrompt.ShowDialog();

            if (_mainViewModel.IsCalibrationFinished)
            {
                await SaveCalibrationDataAsync();
                MessageBox.Show("Kalibrierung abgeschlossen und gespeichert.", "Kalibrierung", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Kalibrierung abgebrochen.", "Kalibrierung", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void XCalibrate_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;

            string modeName = GetCalibrationModeName(_settings?.XMode ?? MachineMode.Laptop);
            _mainViewModel.StartCalibration(modeName, "XReport");

            _activeCalibrationPrompt = new CalibrationPromptWindow(this, _mainViewModel);
            _activeCalibrationPrompt.Owner = this;
            _activeCalibrationPrompt.ShowDialog();

            if (_mainViewModel.IsCalibrationFinished)
            {
                await SaveCalibrationDataAsync();
                MessageBox.Show("Kalibrierung abgeschlossen und gespeichert.", "Kalibrierung", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Kalibrierung abgebrochen.", "Kalibrierung", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
