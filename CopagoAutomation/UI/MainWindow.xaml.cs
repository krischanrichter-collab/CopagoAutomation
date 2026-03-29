
using System;
using CopagoAutomation.Automation;
using System.Drawing; // For System.Drawing.Point
using System.IO;
using System.Linq;
using System.Threading;
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
        private CancellationTokenSource? _abcCts;
        private CancellationTokenSource? _xCts;
        private CancellationTokenSource? _stundenCts;
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

            StundenModeLaptop.Checked += StundenMode_Checked;
            StundenModeDocking.Checked += StundenMode_Checked;

            AbcSammelordner.LostFocus += AbcStorageFields_LostFocus;
            XSammelordner.LostFocus += XStorageFields_LostFocus;
            StundenSammelordner.LostFocus += StundenStorageFields_LostFocus;

            AbcPosList.SelectionChanged += AbcPosList_SelectionChanged;
            XPosList.SelectionChanged += XPosList_SelectionChanged;
            StundenPosList.SelectionChanged += StundenPosList_SelectionChanged;

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

                LoadXYearComboBox();
                // StundenDateSelector setzt sein Default-Datum selbst beim Laden
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
            ApplyStundenStorageSettingsToUi();
            RestorePosSelections();
            ApplyOutputFormatToUi();
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

        private void LoadXYearComboBox()
        {
            int currentYear = DateTime.Today.Year;
            XYear.Items.Clear();
            for (int y = currentYear - 3; y <= currentYear + 1; y++)
                XYear.Items.Add(y);
            XYear.SelectedItem = currentYear;
        }

        private void LoadPosEntries()
        {
            var allPos = PosRepository.GetAll();

            AbcPosList.Items.Clear();
            XPosList.Items.Clear();
            StundenPosList.Items.Clear();

            foreach (string pos in allPos)
            {
                AbcPosList.Items.Add(pos);
                XPosList.Items.Add(pos);
                StundenPosList.Items.Add(pos);
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

        private static string GetCalibrationModeName(MachineMode mode, OutputFormat format)
        {
            string modeBase     = mode   == MachineMode.Laptop  ? "laptop" : "dock";
            string formatSuffix = format == OutputFormat.Excel  ? "_excel" : "_pdf";
            return modeBase + formatSuffix;
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
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
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
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        }

        private async void AbcPosList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_settings == null) return;

            _settings.SelectedAbcPosValues = AbcPosList.SelectedItems.Cast<string>().ToList();
            await SaveSettingsAsync();
        }

        private async void XPosList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_settings == null) return;

            _settings.SelectedXPosValues = XPosList.SelectedItems.Cast<string>().ToList();
            await SaveSettingsAsync();
        }

        private void RestorePosSelections()
        {
            if (_settings == null) return;

            AbcPosList.SelectionChanged -= AbcPosList_SelectionChanged;
            foreach (var item in AbcPosList.Items)
                if (_settings.SelectedAbcPosValues.Contains(item.ToString()))
                    AbcPosList.SelectedItems.Add(item);
            AbcPosList.SelectionChanged += AbcPosList_SelectionChanged;

            XPosList.SelectionChanged -= XPosList_SelectionChanged;
            foreach (var item in XPosList.Items)
                if (_settings.SelectedXPosValues.Contains(item.ToString()))
                    XPosList.SelectedItems.Add(item);
            XPosList.SelectionChanged += XPosList_SelectionChanged;

            StundenPosList.SelectionChanged -= StundenPosList_SelectionChanged;
            foreach (var item in StundenPosList.Items)
                if (_settings.SelectedStundenleistungPosValues.Contains(item.ToString()))
                    StundenPosList.SelectedItems.Add(item);
            StundenPosList.SelectionChanged += StundenPosList_SelectionChanged;
        }

        private async void AbcStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            var dateRange = AbcDateRuleSelector.GetOneTimeRange();
            var occurrences = AbcDateRuleSelector.GenerateOccurrences();

            var request = new AbcStartRequest
            {
                Mode              = _settings.AbcMode,
                OutputFormat      = _settings.AbcOutputFormat,
                SaveMode          = _settings.AbcSaveMode,
                BaseFolder        = _settings.AbcBaseFolder ?? string.Empty,
                SammelordnerPath  = _settings.AbcSammelordnerPath,
                SelectedPosValues = AbcPosList.SelectedItems.Cast<string>().ToList(),
                DateFrom          = dateRange?.from ?? AbcDateRuleSelector.DateFrom,
                DateTo            = dateRange?.to   ?? AbcDateRuleSelector.DateTo,
                OccurrenceDates   = occurrences.Count > 0 ? occurrences : null
            };

            _abcCts = new CancellationTokenSource();
            AbcStartButton.IsEnabled = false;
            AbcStopButton.IsEnabled = true;
            AbcLogBox.Clear();
            try
            {
                var token = _abcCts.Token;
                var logs = await Task.Run(() => _automationService.StartAbcAutomation(request, token), token);
                foreach (var log in logs)
                    AbcLogBox.AppendText(log + Environment.NewLine);
            }
            catch (OperationCanceledException)
            {
                AbcLogBox.AppendText("Automation abgebrochen." + Environment.NewLine);
            }
            catch (Exception ex)
            {
                AbcLogBox.AppendText($"Fehler: {ex.Message}{Environment.NewLine}");
                MessageBox.Show($"Fehler bei der ABC-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                AbcStartButton.IsEnabled = true;
                AbcStopButton.IsEnabled = false;
                _abcCts.Dispose();
                _abcCts = null;
            }
        }

        private async void XStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            int year = XYear.SelectedItem is int y ? y : DateTime.Today.Year;
            int.TryParse(XCumPercent.Text, out int cumPercent);
            int.TryParse(XToWeek.Text, out int toWeek);

            var request = new XStartRequest
            {
                Mode         = _settings.XMode,
                OutputFormat = _settings.XOutputFormat,
                SaveMode     = _settings.XSaveMode,
                BaseFolder   = _settings.XBaseFolder ?? string.Empty,
                SammelordnerPath = _settings.XSammelordnerPath,
                SelectedPosValues = XPosList.SelectedItems.Cast<string>().ToList(),
                Year = year > 0 ? year : DateTime.Today.Year,
                CumPercent = cumPercent,
                FromWeek = toWeek,
                ToWeek = toWeek
            };

            _xCts = new CancellationTokenSource();
            XStartButton.IsEnabled = false;
            XStopButton.IsEnabled = true;
            XLogBox.Clear();
            try
            {
                var token = _xCts.Token;
                var logs = await Task.Run(() => _automationService.StartXAutomation(request, token), token);
                foreach (var log in logs)
                    XLogBox.AppendText(log + Environment.NewLine);
            }
            catch (OperationCanceledException)
            {
                XLogBox.AppendText("Automation abgebrochen." + Environment.NewLine);
            }
            catch (Exception ex)
            {
                XLogBox.AppendText($"Fehler: {ex.Message}{Environment.NewLine}");
                MessageBox.Show($"Fehler bei der X-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                XStartButton.IsEnabled = true;
                XStopButton.IsEnabled = false;
                _xCts.Dispose();
                _xCts = null;
            }
        }

        private void AbcStop_Click(object sender, RoutedEventArgs e)
        {
            _abcCts?.Cancel();
        }

        private void XStop_Click(object sender, RoutedEventArgs e)
        {
            _xCts?.Cancel();
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
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
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
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
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

            MachineMode machineMode = profileName switch
            {
                CalibrationProfiles.XListe          => _settings?.XMode ?? MachineMode.Laptop,
                CalibrationProfiles.Stundenleistung => _settings?.StundenleistungMode ?? MachineMode.Laptop,
                _                                   => _settings?.AbcMode ?? MachineMode.Laptop
            };
            OutputFormat outputFormat = profileName switch
            {
                CalibrationProfiles.XListe          => _settings?.XOutputFormat ?? OutputFormat.Pdf,
                CalibrationProfiles.Stundenleistung => _settings?.StundenleistungOutputFormat ?? OutputFormat.Pdf,
                _                                   => _settings?.AbcOutputFormat ?? OutputFormat.Pdf
            };
            string modeName = GetCalibrationModeName(machineMode, outputFormat);
            bool includeOptional = outputFormat == OutputFormat.Excel;
            _mainViewModel.StartCalibration(modeName, profileName, includeOptional);

            WindowState = WindowState.Minimized;

            var prompt = new CalibrationPromptWindow(this, _mainViewModel);
            if (prompt.ShowDialog() == true)
            {
                await SaveCalibrationDataAsync();
            }
        }

        // ===================== STUNDENLEISTUNG =====================

        private async void StundenMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.StundenleistungMode = StundenModeLaptop.IsChecked == true ? MachineMode.Laptop : MachineMode.Docking;
            await SaveSettingsAsync();
        }

        private async void StundenSaveModeChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.StundenleistungSaveMode = StundenSaveModeSemco.IsChecked == true ? SaveMode.SemcoUpload : SaveMode.Alternativ;
            UpdateStundenSaveModeUi();
            await SaveSettingsAsync();
        }

        private void ApplyStundenStorageSettingsToUi()
        {
            if (_settings == null || StundenSaveModeSemco == null || StundenSaveModeAlt == null || StundenSammelordner == null)
                return;

            if (_settings.StundenleistungSaveMode == SaveMode.SemcoUpload)
                StundenSaveModeSemco.IsChecked = true;
            else
                StundenSaveModeAlt.IsChecked = true;

            StundenSammelordner.Text = _settings.StundenleistungSammelordnerPath;
            UpdateStundenSaveModeUi();
        }

        private void UpdateStundenSaveModeUi()
        {
            if (_settings == null || StundenSammelordner == null) return;

            bool isSemco = _settings.StundenleistungSaveMode == SaveMode.SemcoUpload;
            StundenSammelordner.IsEnabled = !isSemco;
            StundenSammelordner.Text = isSemco ? "" : _settings.StundenleistungSammelordnerPath;
        }

        private async void StundenStorageFields_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_settings == null || StundenSammelordner == null) return;
            _settings.StundenleistungSammelordnerPath = StundenSammelordner.Text;
            await SaveSettingsAsync();
        }

        private async void StundenBrowseSammelordner_Click(object sender, RoutedEventArgs e)
        {
            if (_settings == null || StundenSammelordner == null) return;

            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                _settings.StundenleistungSammelordnerPath = dialog.FolderName;
                StundenSammelordner.Text = dialog.FolderName;
                await SaveSettingsAsync();
            }
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        }

        private async void StundenPosList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_settings == null) return;
            _settings.SelectedStundenleistungPosValues = StundenPosList.SelectedItems.Cast<string>().ToList();
            await SaveSettingsAsync();
        }

        private async void StundenResetPos_Click(object sender, RoutedEventArgs e)
        {
            StundenPosList.SelectedItems.Clear();
            await SaveSettingsAsync();
        }

        private void StundenCalibration_Click(object sender, RoutedEventArgs e)
        {
            RunCalibration(CalibrationProfiles.Stundenleistung);
        }

        private async void StundenStart_Click(object sender, RoutedEventArgs e)
        {
            if (_automationService == null || _settings == null) return;

            var request = new StundenleistungStartRequest
            {
                Mode             = _settings.StundenleistungMode,
                OutputFormat     = _settings.StundenleistungOutputFormat,
                SaveMode         = _settings.StundenleistungSaveMode,
                BaseFolder       = _settings.StundenleistungBaseFolder ?? string.Empty,
                SammelordnerPath = _settings.StundenleistungSammelordnerPath,
                SelectedPosValues = StundenPosList.SelectedItems.Cast<string>().ToList(),
                Dates            = StundenDateSelector.GetDates()
            };

            _stundenCts = new CancellationTokenSource();
            StundenStartButton.IsEnabled = false;
            StundenStopButton.IsEnabled  = true;
            StundenLogBox.Clear();
            try
            {
                var token = _stundenCts.Token;
                var logs = await Task.Run(() => _automationService.StartStundenleistungAutomation(request, token), token);
                foreach (var log in logs)
                    StundenLogBox.AppendText(log + Environment.NewLine);
            }
            catch (OperationCanceledException)
            {
                StundenLogBox.AppendText("Automation abgebrochen." + Environment.NewLine);
            }
            catch (Exception ex)
            {
                StundenLogBox.AppendText($"Fehler: {ex.Message}{Environment.NewLine}");
                MessageBox.Show($"Fehler bei der Stundenleistung-Automatisierung: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                StundenStartButton.IsEnabled = true;
                StundenStopButton.IsEnabled  = false;
                _stundenCts.Dispose();
                _stundenCts = null;
            }
        }

        private void StundenStop_Click(object sender, RoutedEventArgs e)
        {
            _stundenCts?.Cancel();
        }

        // ===================== OUTPUT FORMAT =====================

        private void ApplyOutputFormatToUi()
        {
            if (_settings == null) return;

            AbcFormatPdf.IsChecked      = _settings.AbcOutputFormat      == OutputFormat.Pdf;
            AbcFormatExcel.IsChecked    = _settings.AbcOutputFormat      == OutputFormat.Excel;
            XFormatPdf.IsChecked        = _settings.XOutputFormat        == OutputFormat.Pdf;
            XFormatExcel.IsChecked      = _settings.XOutputFormat        == OutputFormat.Excel;
            StundenFormatPdf.IsChecked  = _settings.StundenleistungOutputFormat == OutputFormat.Pdf;
            StundenFormatExcel.IsChecked= _settings.StundenleistungOutputFormat == OutputFormat.Excel;
        }

        private async void AbcFormatChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.AbcOutputFormat = AbcFormatExcel.IsChecked == true ? OutputFormat.Excel : OutputFormat.Pdf;
            await SaveSettingsAsync();
        }

        private async void XFormatChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.XOutputFormat = XFormatExcel.IsChecked == true ? OutputFormat.Excel : OutputFormat.Pdf;
            await SaveSettingsAsync();
        }

        private async void StundenFormatChanged(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.StundenleistungOutputFormat = StundenFormatExcel.IsChecked == true ? OutputFormat.Excel : OutputFormat.Pdf;
            await SaveSettingsAsync();
        }
    }

}
