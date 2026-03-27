
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
        }

        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            if (_hwndSource != null)
            {
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



        private void CalibrateButton_Click(object sender, RoutedEventArgs e)
        {
            RunCalibration(CalibrationProfiles.AbcAnalyse);
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

        private void AbcStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            var dateRange = AbcDateRuleSelector.GetOneTimeRange();
            var request = new AbcStartRequest
            {
                Mode = _settings.AbcMode,
                SaveMode = _settings.AbcSaveMode,
                BaseFolder = _settings.AbcBaseFolder,
                SammelordnerPath = _settings.AbcSammelordnerPath,
                SelectedPosValues = AbcPosList.SelectedItems.Cast<string>().ToList(),
                DateFrom = dateRange?.from ?? AbcDateRuleSelector.DateFrom,
                DateTo = dateRange?.to ?? AbcDateRuleSelector.DateTo
            };

            AbcLogBox.Clear();
            try
            {
                var logs = _automationService.StartAbcAutomation(request);
                foreach (var log in logs)
                    AbcLogBox.AppendText(log + Environment.NewLine);
            }
            catch (Exception ex)
            {
                AbcLogBox.AppendText($"Fehler: {ex.Message}{Environment.NewLine}");
                MessageBox.Show($"Fehler bei der ABC-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void XStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            int.TryParse(XYear.Text, out int year);
            int.TryParse(XCumPercent.Text, out int cumPercent);
            int.TryParse(XToWeek.Text, out int toWeek);

            var request = new XStartRequest
            {
                Mode = _settings.XMode,
                SaveMode = _settings.XSaveMode,
                BaseFolder = _settings.XBaseFolder,
                SammelordnerPath = _settings.XSammelordnerPath,
                SelectedPosValues = XPosList.SelectedItems.Cast<string>().ToList(),
                Year = year > 0 ? year : DateTime.Today.Year,
                CumPercent = cumPercent,
                ToWeek = toWeek
            };

            XLogBox.Clear();
            try
            {
                var logs = _automationService.StartXAutomation(request);
                foreach (var log in logs)
                    XLogBox.AppendText(log + Environment.NewLine);
            }
            catch (Exception ex)
            {
                XLogBox.AppendText($"Fehler: {ex.Message}{Environment.NewLine}");
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

        private void XCalibration_Click(object sender, RoutedEventArgs e)
        {
            RunCalibration(CalibrationProfiles.XListe);
        }

        private void AbcCalibration_Click(object sender, RoutedEventArgs e)
        {
            RunCalibration(CalibrationProfiles.AbcAnalyse);
        }

        private async void RunCalibration(string profileName)
        {
            if (_mainViewModel == null) return;

            MachineMode machineMode = profileName == CalibrationProfiles.XListe
                ? (_settings?.XMode ?? MachineMode.Laptop)
                : (_settings?.AbcMode ?? MachineMode.Laptop);
            string modeName = GetCalibrationModeName(machineMode);
            _mainViewModel.StartCalibration(modeName, profileName);

            var prompt = new CalibrationPromptWindow(this, _mainViewModel);
            if (prompt.ShowDialog() == true)
            {
                await SaveCalibrationDataAsync();
            }
        }
    }

}
