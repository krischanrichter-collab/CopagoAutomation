using System;
using CopagoAutomation.Automation;
using System.Drawing; // For System.Drawing.Point
using System.Windows.Forms; // For Cursor.Position
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

        private bool TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundCopagoWindow)
        {
            x = 0;
            y = 0;
            boundCopagoWindow = null;

            if (_windowAutomation == null) return false;

            // Get current cursor position
            System.Drawing.Point screenPoint = System.Windows.Forms.Cursor.Position;

            // Try to get the window under the cursor and its root
            IntPtr childWindow = _windowAutomation.WindowFromPoint(screenPoint);
            if (childWindow == IntPtr.Zero) return false;

            IntPtr rootWindow = _windowAutomation.GetAncestor(childWindow, GA_ROOT);
            if (rootWindow == IntPtr.Zero) return false;

            // Convert screen coordinates to client coordinates of the root window
            System.Drawing.Point clientPoint = screenPoint;
            if (!_windowAutomation.ScreenToClient(rootWindow, ref clientPoint)) return false;

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

                _windowAutomation?.RegisterHotKey(handle, normalId, MOD_CONTROL | MOD_ALT, normalVk);
                _windowAutomation?.RegisterHotKey(handle, numpadId, MOD_CONTROL | MOD_ALT, numpadVk);
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

                _windowAutomation?.UnregisterHotKey(handle, normalId);
                _windowAutomation?.UnregisterHotKey(handle, numpadId);
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
                if (TryGetCurrentClientCursorPosition(out int x, out int y, out var boundWindow))
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
            if (_settings == null)
                return;

            if (AbcSammelordner != null)
                AbcSammelordner.Text = _settings.AbcSammelordnerPath;

            if (_settings.AbcSaveMode == SaveMode.Semco)
                AbcSaveModeSemco.IsChecked = true;
            else
                AbcSaveModeAlt.IsChecked = true;
        }

        private void UpdateAbcSaveModeUi()
        {
            bool isAltMode = AbcSaveModeAlt.IsChecked == true;

            if (AbcSammelordner != null)
                AbcSammelordner.IsEnabled = isAltMode;
            if (AbcBrowseSammelordner != null)
                AbcBrowseSammelordner.IsEnabled = isAltMode;
        }

        private async void AbcMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null)
                return;

            if (AbcModeLaptop.IsChecked == true)
                _settings.AbcMode = MachineMode.Laptop;
            else
                _settings.AbcMode = MachineMode.Docking;

            await SaveSettingsAsync();
        }

        private async void AbcSaveModeChanged(object sender, RoutedEventArgs e)
        {
            await SaveAbcStorageSettingsFromUiAsync();
            UpdateAbcSaveModeUi();
        }

        private async void AbcStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            await SaveAbcStorageSettingsFromUiAsync();
        }

        private async Task SaveAbcStorageSettingsFromUiAsync()
        {
            if (_settings == null)
                return;

            if (AbcSaveModeSemco.IsChecked == true)
                _settings.AbcSaveMode = SaveMode.Semco;
            else
                _settings.AbcSaveMode = SaveMode.Alternative;

            if (AbcSammelordner != null)
                _settings.AbcSammelordnerPath = AbcSammelordner.Text;

            await SaveSettingsAsync();
        }

        private async void AbcBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                AbcSammelordner.Text = dialog.FolderName;
                await SaveAbcStorageSettingsFromUiAsync();
            }
        }

        private async void AbcStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null)
                return;

            var selectedPos = AbcPosList.SelectedItems.Cast<string>().ToList();

            var request = new AbcStartRequest
            {
                DateFrom = AbcDateRuleSelector.DateFrom,
                DateTo = AbcDateRuleSelector.DateTo,
                SelectedPosValues = selectedPos,
                BaseFolder = _settings.AbcBaseFolder, // This will be handled by PathResolver
                SaveMode = _settings.AbcSaveMode
            };

            var logs = await Task.Run(() => _automationService.StartAbcAutomation(request, GetCalibrationModeName(_settings.AbcMode)));

            LogAbc.Text = string.Join("\n", logs);
        }

        private async void AbcCalibrate_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null || _settings == null)
                return;

            string profileName = "abc";
            string modeName = GetCalibrationModeName(_settings.AbcMode);

            _mainViewModel.StartCalibration(profileName, modeName);

            using (_activeCalibrationPrompt = new CalibrationPromptWindow(_mainViewModel))
            {
                _activeCalibrationPrompt.Owner = this;
                _activeCalibrationPrompt.ShowDialog();
            }

            await SaveCalibrationDataAsync();
            LogAbc.Text = "Kalibrierung für ABC abgeschlossen.";
        }

        private void ApplyXStorageSettingsToUi()
        {
            if (_settings == null)
                return;

            if (XSammelordner != null)
                XSammelordner.Text = _settings.XSammelordnerPath;

            if (_settings.XSaveMode == SaveMode.Semco)
                XSaveModeSemco.IsChecked = true;
            else
                XSaveModeAlt.IsChecked = true;
        }

        private void UpdateXSaveModeUi()
        {
            bool isAltMode = XSaveModeAlt.IsChecked == true;

            if (XSammelordner != null)
                XSammelordner.IsEnabled = isAltMode;
            if (XBrowseSammelordner != null)
                XBrowseSammelordner.IsEnabled = isAltMode;
        }

        private async void XMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null)
                return;

            if (XModeLaptop.IsChecked == true)
                _settings.XMode = MachineMode.Laptop;
            else
                _settings.XMode = MachineMode.Docking;

            await SaveSettingsAsync();
        }

        private async void XSaveModeChanged(object sender, RoutedEventArgs e)
        {
            await SaveXStorageSettingsFromUiAsync();
            UpdateXSaveModeUi();
        }

        private async void XStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            await SaveXStorageSettingsFromUiAsync();
        }

        private async Task SaveXStorageSettingsFromUiAsync()
        {
            if (_settings == null)
                return;

            if (XSaveModeSemco.IsChecked == true)
                _settings.XSaveMode = SaveMode.Semco;
            else
                _settings.XSaveMode = SaveMode.Alternative;

            if (XSammelordner != null)
                _settings.XSammelordnerPath = XSammelordner.Text;

            await SaveSettingsAsync();
        }

        private async void XBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                XSammelordner.Text = dialog.FolderName;
                await SaveXStorageSettingsFromUiAsync();
            }
        }

        private async void XStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null)
                return;

            var selectedPos = XPosList.SelectedItems.Cast<string>().ToList();

            var request = new XStartRequest
            {
                Year = int.Parse(XYear.Text),
                ToWeek = int.Parse(XToWeek.Text),
                CumPercent = int.Parse(XCumPercent.Text),
                SelectedPosValues = selectedPos,
                BaseFolder = _settings.XBaseFolder, // This will be handled by PathResolver
                SaveMode = _settings.XSaveMode
            };

            var logs = await Task.Run(() => _automationService.StartXAutomation(request, GetCalibrationModeName(_settings.XMode)));

            LogX.Text = string.Join("\n", logs);
        }

        private async void XCalibrate_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null || _settings == null)
                return;

            string profileName = "xlist";
            string modeName = GetCalibrationModeName(_settings.XMode);

            _mainViewModel.StartCalibration(profileName, modeName);

            using (_activeCalibrationPrompt = new CalibrationPromptWindow(_mainViewModel))
            {
                _activeCalibrationPrompt.Owner = this;
                _activeCalibrationPrompt.ShowDialog();
            }

            await SaveCalibrationDataAsync();
            LogX.Text = "Kalibrierung für X-Liste abgeschlossen.";
        }
    }
}
