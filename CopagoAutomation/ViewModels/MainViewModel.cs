using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using CopagoAutomation.Calibration;
using CopagoAutomation.Automation;

namespace CopagoAutomation.ViewModels
{
	public class MainViewModel : INotifyPropertyChanged
	{
		private readonly CalibrationService _calibrationService;
		private CalibrationRunner? _calibrationRunner;
		private string _currentCalibrationModeName = string.Empty;

		private bool _hasLastCapturedPosition;
		private int _lastCapturedX;
		private int _lastCapturedY;
		private BoundWindowInfo? _lastBoundWindow;

		public MainViewModel(CalibrationService calibrationService)
		{
			_calibrationService = calibrationService
				?? throw new ArgumentNullException(nameof(calibrationService));
		}

		public event PropertyChangedEventHandler? PropertyChanged;

		public CalibrationRunner? CalibrationRunner => _calibrationRunner;

		public BoundWindowInfo? LastBoundWindow => _lastBoundWindow;

		public string CurrentCalibrationModeName => _currentCalibrationModeName;

		public string CurrentCalibrationProfileName => _calibrationRunner?.ProfileName ?? string.Empty;

		public CalibrationStepDefinition? CurrentCalibrationStep => _calibrationRunner?.CurrentStep;

		public bool IsCalibrationRunning => _calibrationRunner != null && !_calibrationRunner.IsFinished;

		public bool HasLastCapturedPosition => _hasLastCapturedPosition;

		public int LastCapturedX => _lastCapturedX;

		public int LastCapturedY => _lastCapturedY;

		public void StartCalibration(string modeName, string profileName)
		{
			if (string.IsNullOrWhiteSpace(modeName))
				throw new ArgumentException("modeName darf nicht leer sein.", nameof(modeName));

			if (string.IsNullOrWhiteSpace(profileName))
				throw new ArgumentException("profileName darf nicht leer sein.", nameof(profileName));

			_currentCalibrationModeName = modeName;
			_calibrationRunner = new CalibrationRunner(profileName);

			ResetLastCapture();
			NotifyCalibrationStateChanged();
		}

		public bool SetLastCapturedPosition(int x, int y, BoundWindowInfo? boundWindow)
		{
			if (_calibrationRunner == null || _calibrationRunner.IsFinished)
				return false;

			if (_calibrationRunner.CurrentStep == null)
				return false;

			_lastCapturedX = x;
			_lastCapturedY = y;
			_lastBoundWindow = boundWindow;
			_hasLastCapturedPosition = true;

			OnPropertyChanged(nameof(HasLastCapturedPosition));
			OnPropertyChanged(nameof(LastCapturedX));
			OnPropertyChanged(nameof(LastCapturedY));
			OnPropertyChanged(nameof(LastBoundWindow));

			return true;
		}

		public bool SaveCurrentCalibrationPoint(BoundWindowInfo? boundWindow = null)
		{
			if (_calibrationRunner == null || _calibrationRunner.IsFinished)
				return false;

			var currentStep = _calibrationRunner.CurrentStep;
			if (currentStep == null)
				return false;

			if (!_hasLastCapturedPosition)
				return false;

			if (string.IsNullOrWhiteSpace(_currentCalibrationModeName))
				return false;

			_calibrationService.SetPoint(
				_currentCalibrationModeName,
				_calibrationRunner.ProfileName,
				currentStep.Key,
				_lastCapturedX,
				_lastCapturedY,
				boundWindow);

			_calibrationRunner.MoveNext();
			ResetLastCapture();
			NotifyCalibrationStateChanged();

			return true;
		}

		public bool TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundCopagoWindow)
		{
			x = 0;
			y = 0;
			boundCopagoWindow = null;

			if (_calibrationService.WindowAutomation == null)
				return false;

			System.Drawing.Point screenPoint = _calibrationService.WindowAutomation.GetCursorScreenPosition();

			IntPtr childWindow = WindowAutomation.WindowFromPoint(new WindowAutomation.POINT { X = screenPoint.X, Y = screenPoint.Y });
			if (childWindow != IntPtr.Zero)
			{
				IntPtr rootWindow = WindowAutomation.GetAncestor(childWindow, WindowAutomation.GA_ROOT);
				if (rootWindow != IntPtr.Zero && _calibrationService.WindowAutomation.TryBindWindowByHandle(rootWindow, out var copagoWindow))
					boundCopagoWindow = copagoWindow;
			}

			x = screenPoint.X;
			y = screenPoint.Y;
			return true;
		}

		public void CancelCalibration()
		{
			_calibrationRunner = null;
			_currentCalibrationModeName = string.Empty;

			ResetLastCapture();
			NotifyCalibrationStateChanged();
		}

		private void ResetLastCapture()
		{
			_hasLastCapturedPosition = false;
			_lastCapturedX = 0;
			_lastCapturedY = 0;
			_lastBoundWindow = null;

			OnPropertyChanged(nameof(HasLastCapturedPosition));
			OnPropertyChanged(nameof(LastCapturedX));
			OnPropertyChanged(nameof(LastCapturedY));
		}

		private void NotifyCalibrationStateChanged()
		{
			OnPropertyChanged(nameof(CurrentCalibrationModeName));
			OnPropertyChanged(nameof(CurrentCalibrationProfileName));
			OnPropertyChanged(nameof(CurrentCalibrationStep));
			OnPropertyChanged(nameof(IsCalibrationRunning));
			OnPropertyChanged(nameof(HasLastCapturedPosition));
		}

		protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
	}
}
