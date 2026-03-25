using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using CopagoAutomation.Calibration;
using CopagoAutomation.Automation; // Hinzugefügt, um WindowAutomation.BoundWindowInfo zu verwenden
using CopagoAutomation.Windows; // Hinzugefügt, um CalibrationPromptWindow zu verwenden

namespace CopagoAutomation.ViewModels
{
	public class MainViewModel : INotifyPropertyChanged
	{
		private readonly CalibrationService _calibrationService;
		private CalibrationRunner? _calibrationRunner;

		public MainViewModel(CalibrationService calibrationService)
		{
			_calibrationService = calibrationService
				?? throw new ArgumentNullException(nameof(calibrationService));
		}

		public event PropertyChangedEventHandler? PropertyChanged;

		public CalibrationRunner? CalibrationRunner => _calibrationRunner;
		public int CurrentCalibrationStepIndex
		{
			get
			{
				if (_calibrationRunner == null)
					return 0;

				return _calibrationRunner.CurrentIndex;
			}
		}
		private string _currentCalibrationModeName = string.Empty;

		private bool _hasLastCapturedPosition;
		private int _lastCapturedX;
            private int _lastCapturedY;
            private BoundWindowInfo? _lastBoundWindow;
            private CalibrationPromptWindow? _activeCalibrationPrompt;

            public BoundWindowInfo? LastBoundWindow => _lastBoundWindow;
            public CalibrationPromptWindow? ActiveCalibrationPrompt
            {
                get => _activeCalibrationPrompt;
                set
                {
                    _activeCalibrationPrompt = value;
                    OnPropertyChanged(nameof(ActiveCalibrationPrompt));
                }
            }

		public string CurrentCalibrationModeName => _currentCalibrationModeName;

		public string CurrentCalibrationProfileName => _calibrationRunner?.ProfileName ?? string.Empty;

		public CalibrationStepDefinition? CurrentCalibrationStep => _calibrationRunner?.CurrentStep;

		public string CurrentCalibrationTitle => CurrentCalibrationStep?.Title ?? string.Empty;

		public string CurrentCalibrationInstructionText => CurrentCalibrationStep?.InstructionText ?? string.Empty;

		public string CurrentCalibrationHotkeyText => CurrentCalibrationStep?.HotkeyText ?? string.Empty;

		public bool IsCalibrationRunning => _calibrationRunner != null && !_calibrationRunner.IsFinished;

		public bool IsCalibrationFinished => _calibrationRunner != null && _calibrationRunner.IsFinished;

		public bool HasCurrentCalibrationStep => CurrentCalibrationStep != null;

		public bool HasCalibrationRunner => _calibrationRunner != null;

		public bool HasLastCapturedPosition => _hasLastCapturedPosition;

		public int LastCapturedX => _lastCapturedX;

		public int LastCapturedY => _lastCapturedY;

		public bool IsCurrentCalibrationProfileComplete
		{
			get
			{
				if (string.IsNullOrWhiteSpace(CurrentCalibrationModeName))
					return false;

				if (string.IsNullOrWhiteSpace(CurrentCalibrationProfileName))
					return false;

				return _calibrationService.IsProfileComplete(
					CurrentCalibrationModeName,
					CurrentCalibrationProfileName);
			}
		}

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
                _activeCalibrationPrompt?.SetCapturedPosition(x, y);

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

        public void NextStep(BoundWindowInfo? boundWindow = null)
		{
			SaveCurrentCalibrationPoint(boundWindow);
		}

		public void CancelCalibration()
		{
			_calibrationRunner = null;
			_currentCalibrationModeName = string.Empty;

			ResetLastCapture();
			NotifyCalibrationStateChanged();
		}

		public void ResetCalibration()
		{
			if (_calibrationRunner == null)
				return;

			_calibrationRunner.Reset();
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
			OnPropertyChanged(nameof(CurrentCalibrationTitle));
			OnPropertyChanged(nameof(CurrentCalibrationInstructionText));
			OnPropertyChanged(nameof(CurrentCalibrationHotkeyText));
			OnPropertyChanged(nameof(IsCalibrationRunning));
			OnPropertyChanged(nameof(IsCalibrationFinished));
			OnPropertyChanged(nameof(HasCurrentCalibrationStep));
			OnPropertyChanged(nameof(HasCalibrationRunner));
			OnPropertyChanged(nameof(IsCurrentCalibrationProfileComplete));
			OnPropertyChanged(nameof(HasLastCapturedPosition));
			OnPropertyChanged(nameof(LastCapturedX));
			OnPropertyChanged(nameof(LastCapturedY));
		}

		protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
	}
}
