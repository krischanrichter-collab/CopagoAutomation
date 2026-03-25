
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
        }

        private void MainWindow_Closed(object? sender, EventArgs e)
        {


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

            // ApplyModesToUi();
            // ApplyStorageSettingsToUi();
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

        private bool TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundCopagoWindow)
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

            IntPtr rootWindow = WindowAutomation.GetAncestor(childWindow, WindowAutomation.GA_ROOT);
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




        }







        private void CalibrateButton_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;

            _activeCalibrationPrompt = new CalibrationPromptWindow(this, _mainViewModel);
            _activeCalibrationPrompt.Owner = this;
            _activeCalibrationPrompt.Show();

            _mainViewModel.StartCalibration("laptop", "ABC"); // Example default values
        }

        private async void AbcMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (AbcModeLaptop.IsChecked == true)
                _settings.AbcMode = MachineMode.Laptop;
            else if (AbcModeDocking.IsChecked == true)
                _settings.AbcMode = MachineMode.Docking;

            await SaveSettingsAsync();
        }

        private async void XMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (XModeLaptop.IsChecked == true)
                _settings.XMode = MachineMode.Laptop;
            else if (XModeDocking.IsChecked == true)
                _settings.XMode = MachineMode.Docking;

            await SaveSettingsAsync();
        }

        private async void AbcSaveModeChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (AbcSaveModeSemco.IsChecked == true)
                _settings.AbcSaveMode = SaveMode.SemcoUpload;
            else if (AbcSaveModeAlt.IsChecked == true)
                _settings.AbcSaveMode = SaveMode.Alternativ;

            UpdateAbcSaveModeUi();
            await SaveSettingsAsync();
        }

        private void UpdateAbcSaveModeUi()
        {
            if (_settings == null || AbcSammelordner == null) return;

            bool isSemco = _settings.AbcSaveMode == SaveMode.SemcoUpload;

            AbcSammelordner.IsEnabled = !isSemco;
            // AbcBrowseButton.IsEnabled = !isSemco;

            if (isSemco)
            {
                AbcSammelordner.Text = ""; // Clear path for Semco Upload
            }
            else
            {
                AbcSammelordner.Text = _settings.AbcSammelordnerPath; // Restore path for Alternativ
            }
        }

        private async void XSaveModeChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;

            if (XSaveModeSemco.IsChecked == true)
                _settings.XSaveMode = SaveMode.SemcoUpload;
            else if (XSaveModeAlt.IsChecked == true)
                _settings.XSaveMode = SaveMode.Alternativ;

            UpdateXSaveModeUi();
            await SaveSettingsAsync();
        }

        private void UpdateXSaveModeUi()
        {
            if (_settings == null || XSammelordner == null) return;

            bool isSemco = _settings.XSaveMode == SaveMode.SemcoUpload;

            XSammelordner.IsEnabled = !isSemco;
            // XBrowseButton.IsEnabled = !isSemco;

            if (isSemco)
            {
                XSammelordner.Text = ""; // Clear path for Semco Upload
            }
            else
            {
                XSammelordner.Text = _settings.XSammelordnerPath; // Restore path for Alternativ
            }
        }

        private async void AbcStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_settings == null || AbcSammelordner == null) return;

            _settings.AbcSammelordnerPath = AbcSammelordner.Text;
            await SaveSettingsAsync();
        }

        private async void XStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_settings == null || XSammelordner == null) return;

            _settings.XSammelordnerPath = XSammelordner.Text;
            await SaveSettingsAsync();
        }

        private async void AbcBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                AbcSammelordner.Text = dialog.FolderName;
                if (_settings != null)
                {
                    _settings.AbcSammelordnerPath = dialog.FolderName;
                    await SaveSettingsAsync();
                }
            }
        }

        private async void XBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                XSammelordner.Text = dialog.FolderName;
                if (_settings != null)
                {
                    _settings.XSammelordnerPath = dialog.FolderName;
                    await SaveSettingsAsync();
                }
            }
        }

        private async void AbcPosList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_settings == null || AbcPosList.SelectedItem == null) return;

            _settings.LastPosId = AbcPosList.SelectedItem.ToString();
            await SaveSettingsAsync();
        }

        private async void XPosList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_settings == null || XPosList.SelectedItem == null) return;

            _settings.LastPosId = XPosList.SelectedItem.ToString();
            await SaveSettingsAsync();
        }

        private async void AbcStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            var request = new AbcStartRequest
            {
                Mode = _settings.AbcMode,
                SaveMode = _settings.AbcSaveMode,
                BaseFolder = _settings.AbcBaseFolder,
                SammelordnerPath = _settings.AbcSammelordnerPath,
                SelectedPosValues = AbcPosList.SelectedItems.Cast<string>().ToList()
            };

            try
            {
                var logs = _automationService.StartAbcAutomation(request);
                foreach (var log in logs)
                {
                    // TODO: Display logs in UI
                    Console.WriteLine(log);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler bei der ABC-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void XStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            var request = new XStartRequest
            {
                Mode = _settings.XMode,
                SaveMode = _settings.XSaveMode,
                BaseFolder = _settings.XBaseFolder,
                SammelordnerPath = _settings.XSammelordnerPath,
                SelectedPosValues = XPosList.SelectedItems.Cast<string>().ToList()
            };

            try
            {
                var logs = _automationService.StartXAutomation(request);
                foreach (var log in logs)
                {
                    // TODO: Display logs in UI
                    Console.WriteLine(log);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler bei der X-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AbcStop_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement stop logic for ABC automation
            MessageBox.Show("ABC Automation Stop (Not Implemented)", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void XStop_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement stop logic for X automation
            MessageBox.Show("X Automation Stop (Not Implemented)", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async void AbcResetPos_Click(object sender, RoutedEventArgs e)
        {
            AbcPosList.SelectedItems.Clear();
            await SaveSettingsAsync();
        }

        private async void XResetPos_Click(object sender, RoutedEventArgs e)
        {
            XPosList.SelectedItems.Clear();
            await SaveSettingsAsync();
        }

        private async void AbcBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            if (_settings == null || AbcSammelordner == null) return;

            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                _settings.AbcSammelordnerPath = dialog.FolderName;
                AbcSammelordner.Text = dialog.FolderName;
                       await SaveSettingsAsync();
            }
        }

        private async void XBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            if (_settings == null || XSammelordner == null) return;

            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                _settings.XSammelordnerPath = dialog.FolderName;
                XSammelordner.Text = dialog.FolderName;
                await SaveSettingsAsync();
            }
        }

        private void AbcCapturePoint_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;
            // Logic for capturing a point for ABC automation
            MessageBox.Show("ABC Capture Point (Not Implemented)", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void XCapturePoint_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;
            // Logic for capturing a point for X automation
            MessageBox.Show("X Capture Point (Not Implemented)", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void XCalibration_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;
            _activeCalibrationPrompt = new CalibrationPromptWindow(this, _mainViewModel);
            _activeCalibrationPrompt.Owner = this;
            _mainViewModel.StartCalibration("laptop", "X-Liste"); // Example default values
            if (_activeCalibrationPrompt.ShowDialog() == true)
            {
                _mainViewModel.SaveCurrentCalibrationPoint(_mainViewModel.LastBoundWindow);
            }
            _activeCalibrationPrompt = null;
        }

        private void AbcCalibration_Click(object sender, RoutedEventArgs e)
        {
            if (_mainViewModel == null) return;
            _activeCalibrationPrompt = new CalibrationPromptWindow(this, _mainViewModel);
            _activeCalibrationPrompt.Owner = this;
            _mainViewModel.StartCalibration("laptop", "ABC Analyse"); // Example default values
            if (_activeCalibrationPrompt.ShowDialog() == true)
            {
                _mainViewModel.SaveCurrentCalibrationPoint(_mainViewModel.LastBoundWindow);
            }
            _activeCalibrationPrompt = null;
        }
    }
