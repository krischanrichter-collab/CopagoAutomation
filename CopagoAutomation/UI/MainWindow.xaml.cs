using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

		[StructLayout(LayoutKind.Sequential)]
		private struct POINT
		{
			public int X;
			public int Y;
		}

		[DllImport("user32.dll")]
		private static extern bool GetCursorPos(out POINT lpPoint);

		[DllImport("user32.dll")]
		private static extern IntPtr WindowFromPoint(POINT point);

		[DllImport("user32.dll")]
		private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

		[DllImport("user32.dll")]
		private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

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

			// ABC Speicherfelder automatisch speichern, wenn der Nutzer manuell tippt
			// AbcBaseFolder.LostFocus += AbcStorageFields_LostFocus; // Nicht mehr direkt editierbar
			AbcSammelordner.LostFocus += AbcStorageFields_LostFocus;

			// X-Liste Speicherfelder automatisch speichern, wenn der Nutzer manuell tippt
			// XBaseFolder.LostFocus += XStorageFields_LostFocus; // Nicht mehr direkt editierbar
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
				_calibrationService = new CalibrationService(_calibrationData);

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

		private bool TryGetCurrentClientCursorPosition(out int x, out int y)
		{
			x = 0;
			y = 0;

			if (!GetCursorPos(out POINT screenPoint))
				return false;

			IntPtr childWindow = WindowFromPoint(screenPoint);
			if (childWindow == IntPtr.Zero)
				return false;

			IntPtr rootWindow = GetAncestor(childWindow, GA_ROOT);
			if (rootWindow == IntPtr.Zero)
				return false;

			POINT clientPoint = screenPoint;

			if (!ScreenToClient(rootWindow, ref clientPoint))
				return false;

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

				RegisterHotKey(handle, normalId, MOD_CONTROL | MOD_ALT, normalVk);
				RegisterHotKey(handle, numpadId, MOD_CONTROL | MOD_ALT, numpadVk);
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
			if (_mainViewModel == null)
				return;

			if (!_mainViewModel.IsCalibrationRunning)
				return;

			if (_mainViewModel.CurrentCalibrationStep == null)
				return;

			if (_activeCalibrationPrompt == null)
				return;

			if (!TryGetDigitFromHotkeyId(hotkeyId, out int pressedDigit))
				return;

			int expectedDigit = (_mainViewModel.CalibrationRunner?.CurrentIndex ?? 0) + 1;

			if (pressedDigit != expectedDigit)
				return;

			if (!TryGetCurrentClientCursorPosition(out int x, out int y))
				return;

			bool captured = _mainViewModel.SetLastCapturedPosition(x, y);
			if (!captured)
				return;

			_activeCalibrationPrompt.SetCapturedPosition(x, y);
		}

		private void LogCurrentCalibrationStepToAbc()
		{
			if (_mainViewModel == null)
			{
				LogAbc("Kalibrierung nicht verfügbar.");
				return;
			}

			if (_mainViewModel.IsCalibrationFinished)
			{
				LogAbc("Kalibrierung abgeschlossen.");
				return;
			}

			if (_mainViewModel.CurrentCalibrationStep == null)
			{
				LogAbc("Kein aktueller Kalibrierschritt vorhanden.");
				return;
			}

			LogAbc($"Schritt: {_mainViewModel.CurrentCalibrationTitle}");
			LogAbc($"Hotkey: {_mainViewModel.CurrentCalibrationHotkeyText}");
			LogAbc(_mainViewModel.CurrentCalibrationInstructionText);
		}

		private void LogCurrentCalibrationStepToX()
		{
			if (_mainViewModel == null)
			{
				LogX("Kalibrierung nicht verfügbar.");
				return;
			}

			if (_mainViewModel.IsCalibrationFinished)
			{
				LogX("Kalibrierung abgeschlossen.");
				return;
			}

			if (_mainViewModel.CurrentCalibrationStep == null)
			{
				LogX("Kein aktueller Kalibrierschritt vorhanden.");
				return;
			}

			LogX($"Schritt: {_mainViewModel.CurrentCalibrationTitle}");
			LogX($"Hotkey: {_mainViewModel.CurrentCalibrationHotkeyText}");
			LogX(_mainViewModel.CurrentCalibrationInstructionText);
		}

		private string? BrowseForFolder(string currentPath)
		{
			var dialog = new OpenFolderDialog
			{
				Title = "Ordner auswählen"
			};

			if (!string.IsNullOrWhiteSpace(currentPath) && Directory.Exists(currentPath))
			{
				dialog.InitialDirectory = currentPath;
			}

			if (dialog.ShowDialog() == true)
			{
				return dialog.FolderName;
			}

			return null;
		}

		private async Task RunCalibrationWithPromptAsync(
			string modeName,
			string profileName,
			string profileDisplayName,
			Action<string> logger)
		{
			if (_mainViewModel == null)
				return;

			this.Hide();

			try
			{
				_mainViewModel.StartCalibration(modeName, profileName);
				logger($"Kalibrierung für '{profileDisplayName}' ({modeName}) gestartet.");

				while (!_mainViewModel.IsCalibrationFinished)
				{
					var step = _mainViewModel.CurrentCalibrationStep;
					if (step == null) break;

					using (var prompt = new CalibrationPromptWindow())
					{
						prompt.Owner = null;
						prompt.WindowStartupLocation = WindowStartupLocation.CenterScreen;
						prompt.Topmost = true;

						prompt.SetStepInfo(
							_mainViewModel.CurrentCalibrationTitle,
							_mainViewModel.CurrentCalibrationInstructionText,
							_mainViewModel.CurrentCalibrationHotkeyText);

						_activeCalibrationPrompt = prompt;

						bool? result = prompt.ShowDialog();

						_activeCalibrationPrompt = null;

						if (result == true)
						{
							_mainViewModel.NextStep();
						}
						else
						{
							_mainViewModel.CancelCalibration();
							logger("Kalibrierung abgebrochen.");
							return;
						}
					}
				}

				if (_mainViewModel.IsCalibrationFinished)
				{
					await SaveCalibrationDataAsync();
					logger("Kalibrierung erfolgreich abgeschlossen und gespeichert.");
				}
			}
			finally
			{
				this.Show();
				this.Activate();
			}
		}

		private async void AbcMode_Checked(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (AbcModeLaptop.IsChecked == true)
				_settings.AbcMode = MachineMode.Laptop;
			else
				_settings.AbcMode = MachineMode.Docking;

			UpdateAbcIniText();
			await SaveSettingsAsync();
		}

		private void UpdateAbcIniText()
		{
			string baseDir = AppDomain.CurrentDomain.BaseDirectory;

			if (_settings?.AbcMode == MachineMode.Laptop)
			{
				AbcActiveIniText.Text =
					$@"Laptop | INI: {Path.Combine(baseDir, "calibration_laptop.ini")}";
			}
			else
			{
				AbcActiveIniText.Text =
					$@"Docking | INI: {Path.Combine(baseDir, "calibration_docking.ini")}";
			}
		}

		private async void XMode_Checked(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (XModeLaptop.IsChecked == true)
				_settings.XMode = MachineMode.Laptop;
			else
				_settings.XMode = MachineMode.Docking;

			UpdateXIniText();
			await SaveSettingsAsync();
		}

		private void UpdateXIniText()
		{
			string baseDir = AppDomain.CurrentDomain.BaseDirectory;

			if (_settings?.XMode == MachineMode.Laptop)
			{
				XActiveIniText.Text =
					$@"Laptop | INI: {Path.Combine(baseDir, "calibration_laptop.ini")}";
			}
			else
			{
				XActiveIniText.Text =
					$@"Docking | INI: {Path.Combine(baseDir, "calibration_docking.ini")}";
			}
		}

		private void ApplyModesToUi()
		{
			if (_settings == null)
				_settings = new AppSettings();

			AbcModeLaptop.IsChecked = _settings.AbcMode == MachineMode.Laptop;
			AbcModeDocking.IsChecked = _settings.AbcMode == MachineMode.Docking;

			XModeLaptop.IsChecked = _settings.XMode == MachineMode.Laptop;
			XModeDocking.IsChecked = _settings.XMode == MachineMode.Docking;

			UpdateAbcIniText();
			UpdateXIniText();
		}

		private void ApplyStorageSettingsToUi()
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (AbcSammelordner == null || AbcSaveModeSemco == null || AbcSaveModeAlt == null)
				return;

			// AbcBaseFolder wird nicht mehr direkt in der UI angezeigt oder bearbeitet
			AbcSammelordner.Text = _settings.AbcSammelordnerPath ?? string.Empty;

			AbcSaveModeSemco.IsChecked = _settings.AbcSaveMode == SaveMode.SemcoUpload;
			AbcSaveModeAlt.IsChecked = _settings.AbcSaveMode == SaveMode.Alternativ;
		}

		private void ApplyXStorageSettingsToUi()
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (XSammelordner == null || XSaveModeSemco == null || XSaveModeAlt == null)
				return;

			// XBaseFolder wird nicht mehr direkt in der UI angezeigt oder bearbeitet
			XSammelordner.Text = _settings.XSammelordnerPath ?? string.Empty;

			XSaveModeSemco.IsChecked = _settings.XSaveMode == SaveMode.SemcoUpload;
			XSaveModeAlt.IsChecked = _settings.XSaveMode == SaveMode.Alternativ;
		}

		private async Task SaveAbcStorageSettingsFromUiAsync()
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (AbcSaveModeAlt == null || AbcSammelordner == null)
				return;

			_settings.AbcSaveMode = AbcSaveModeAlt.IsChecked == true
				? SaveMode.Alternativ
				: SaveMode.SemcoUpload;

			// AbcBaseFolder wird nicht mehr direkt aus der UI gelesen, da es nicht mehr editierbar ist
			_settings.AbcSammelordnerPath = AbcSammelordner.Text?.Trim();

			await SaveSettingsAsync();
		}

		private async Task SaveXStorageSettingsFromUiAsync()
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (XSaveModeAlt == null || XSammelordner == null)
				return;

			_settings.XSaveMode = XSaveModeAlt.IsChecked == true
				? SaveMode.Alternativ
				: SaveMode.SemcoUpload;

			// XBaseFolder wird nicht mehr direkt aus der UI gelesen, da es nicht mehr editierbar ist
			_settings.XSammelordnerPath = XSammelordner.Text?.Trim();

			await SaveSettingsAsync();
		}

		private void UpdateAbcSaveModeUi()
		{
			if (AbcSammelordner == null || AbcSaveModeSemco == null || AbcSaveModeAlt == null)
				return;

			bool isSemco = AbcSaveModeSemco.IsChecked == true;
			// AbcBaseFolder ist im Semco-Modus nicht mehr direkt in der UI editierbar
			// AbcSammelordner ist nur im Alternativ-Modus editierbar
			AbcSammelordner.IsEnabled = !isSemco;
		}

		private void UpdateXSaveModeUi()
		{
			if (XSammelordner == null || XSaveModeSemco == null || XSaveModeAlt == null)
				return;

			bool isSemco = XSaveModeSemco.IsChecked == true;
			// XBaseFolder ist im Semco-Modus nicht mehr direkt in der UI editierbar
			// XSammelordner ist nur im Alternativ-Modus editierbar
			XSammelordner.IsEnabled = !isSemco;
		}

		private async void AbcSaveModeChanged(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (AbcSaveModeSemco == null || AbcSaveModeAlt == null)
				return;

			UpdateAbcSaveModeUi();
			await SaveAbcStorageSettingsFromUiAsync();

			if (AbcSaveModeSemco.IsChecked == true)
				LogAbc("Speicher-Modus: Semco Upload");
			else
				LogAbc("Speicher-Modus: Alternativer Ordner");
		}

		private async void XSaveModeChanged(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				_settings = new AppSettings();

			if (XSaveModeSemco == null || XSaveModeAlt == null)
				return;

			UpdateXSaveModeUi();
			await SaveXStorageSettingsFromUiAsync();

			if (XSaveModeSemco.IsChecked == true)
				LogX("Speicher-Modus: Semco Upload");
			else
				LogX("Speicher-Modus: Alternativer Ordner");
		}

		private async void AbcStorageFields_LostFocus(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				return;

			await SaveAbcStorageSettingsFromUiAsync();
		}

		private async void XStorageFields_LostFocus(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				return;

			await SaveXStorageSettingsFromUiAsync();
		}

		private async void AbcCalibration_Click(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				_settings = new AppSettings();

			await RunCalibrationWithPromptAsync(
				GetCalibrationModeName(_settings.AbcMode),
				CalibrationProfiles.AbcAnalyse,
				"ABC Analyse",
				LogAbc);
		}

		private void AbcCapturePoint_Click(object sender, RoutedEventArgs e)
		{
			LogAbc("Der alte Button 'Mauspunkt übernehmen' wird für den neuen Dialog-Workflow nicht mehr verwendet.");
		}

		private async void AbcBrowseSammelordner_Click(object sender, RoutedEventArgs e)
		{
			string currentPath = AbcSammelordner.Text?.Trim() ?? string.Empty;
			string? selectedFolder = BrowseForFolder(currentPath);

			if (!string.IsNullOrWhiteSpace(selectedFolder))
			{
				AbcSammelordner.Text = selectedFolder;
				await SaveAbcStorageSettingsFromUiAsync();
				LogAbc($"Alternativer Ordner (ABC) gewählt: {selectedFolder}");
			}
			else
			{
				LogAbc("Browse (ABC Alternativer Ordner) abgebrochen.");
			}
		}

		private void AbcResetPos_Click(object sender, RoutedEventArgs e)
		{
			AbcPosList.UnselectAll();
			LogAbc("POS Auswahl zurückgesetzt.");
		}

		private async void AbcStart_Click(object sender, RoutedEventArgs e)
		{
			if (_automationService == null)
			{
				LogAbc("Automation ist noch nicht bereit. Bitte Fenster neu laden oder Kalibrierung prüfen.");
				return;
			}

			if (_settings == null)
				_settings = new AppSettings();

			await SaveAbcStorageSettingsFromUiAsync();

			var selectedPosValues = AbcPosList.SelectedItems
				.Cast<object>()
				.Select(item => item?.ToString() ?? string.Empty)
				.Where(value => !string.IsNullOrWhiteSpace(value))
				.ToList();

			string activePath = _settings.AbcSaveMode == SaveMode.SemcoUpload
				? (_settings.AbcBaseFolder ?? string.Empty)
				: (_settings.AbcSammelordnerPath ?? string.Empty);

			var request = new AbcStartRequest
			{
				Mode = _settings.AbcMode,
				BaseFolder = activePath,
				UseSammelordner = _settings.AbcSaveMode == SaveMode.Alternativ,
				SelectedPosCount = selectedPosValues.Count,
				SelectedPosValues = selectedPosValues,
				DateFrom = AbcDateRuleSelector.DateFrom,
				DateTo = AbcDateRuleSelector.DateTo
			};

			LogAbc($"Start ABC mit Speicher-Modus: {_settings.AbcSaveMode}");
			LogAbc($"Aktiver Zielpfad: {activePath}");

			var results = _automationService.StartAbcAutomation(request);

			foreach (var line in results)
			{
				LogAbc(line);
			}
		}

		private void AbcStop_Click(object sender, RoutedEventArgs e)
		{
			LogAbc("Stop (ABC) gedrückt.");
		}

		private async void XCalibration_Click(object sender, RoutedEventArgs e)
		{
			if (_settings == null)
				_settings = new AppSettings();

			await RunCalibrationWithPromptAsync(
				GetCalibrationModeName(_settings.XMode),
				CalibrationProfiles.XListe,
				"X-Liste",
				LogX);
		}

		private void XCapturePoint_Click(object sender, RoutedEventArgs e)
		{
			LogX("Der alte Button 'Mauspunkt übernehmen' wird für den neuen Dialog-Workflow nicht mehr verwendet.");
		}

		private async void XBrowseSammelordner_Click(object sender, RoutedEventArgs e)
		{
			string currentPath = XSammelordner.Text?.Trim() ?? string.Empty;
			string? selectedFolder = BrowseForFolder(currentPath);

			if (!string.IsNullOrWhiteSpace(selectedFolder))
			{
				XSammelordner.Text = selectedFolder;
				await SaveXStorageSettingsFromUiAsync();
				LogX($"Alternativer Ordner (X-Liste) gewählt: {selectedFolder}");
			}
			else
			{
				LogX("Browse (X-Liste Alternativer Ordner) abgebrochen.");
			}
		}

		private void XResetPos_Click(object sender, RoutedEventArgs e)
		{
			XPosList.UnselectAll();
			LogX("POS Auswahl zurückgesetzt.");
		}

		private async void XStart_Click(object sender, RoutedEventArgs e)
		{
			if (_automationService == null)
			{
				LogX("Automation ist noch nicht bereit. Bitte Fenster neu laden oder Kalibrierung prüfen.");
				return;
			}

			if (_settings == null)
				_settings = new AppSettings();

			await SaveXStorageSettingsFromUiAsync();

			var selectedPosValues = XPosList.SelectedItems
				.Cast<object>()
				.Select(item => item?.ToString() ?? string.Empty)
				.Where(value => !string.IsNullOrWhiteSpace(value))
				.ToList();

			string activePath = _settings.XSaveMode == SaveMode.SemcoUpload
				? (_settings.XBaseFolder ?? string.Empty)
				: (_settings.XSammelordnerPath ?? string.Empty);

			var request = new XStartRequest
			{
				Mode = _settings.XMode,
				BaseFolder = activePath,
				UseSammelordner = _settings.XSaveMode == SaveMode.Alternativ,
				SelectedPosCount = selectedPosValues.Count,
				SelectedPosValues = selectedPosValues,
				Year = int.Parse(XYear.Text),
				CumPercent = int.Parse(XCumPercent.Text),
				ToWeek = int.Parse(XToWeek.Text),
				SaveMode = _settings.XSaveMode
			};

			LogX($"Start X-Liste mit Speicher-Modus: {_settings.XSaveMode}");
			LogX($"Aktiver Zielpfad: {activePath}");

			var results = _automationService.StartXAutomation(request);

			foreach (var line in results)
			{
				LogX(line);
			}
		}

		private void XStop_Click(object sender, RoutedEventArgs e)
		{
			LogX("Stop (X-Liste) gedrückt.");
		}

		private void LogAbc(string msg)
		{
			AbcLogBox.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}\n");
			AbcLogBox.ScrollToEnd();
		}

		private void LogX(string msg)
		{
			XLogBox.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}\n");
			XLogBox.ScrollToEnd();
		}
	}
}
