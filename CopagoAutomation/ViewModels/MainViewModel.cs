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

            public bool SetLastCapturedPositionForDigit(int digit)
            {
                // This method is a placeholder. In a real scenario, you might have predefined positions for digits.
                // For now, we'll just use the last captured position.
                // Or, if the digit is 0, it means capture current cursor position.
                if (digit == 0)
                {
                    if (TryGetCurrentClientCursorPosition(out int x, out int y, out BoundWindowInfo? boundCopagoWindow))
                    {
                        SetLastCapturedPosition(x, y, boundCopagoWindow);
    
                    return true;
                    }
                }
                // For other digits, we might need a different logic or pre-defined points.
                // For now, we'll just return false if it's not digit 0.
                return false;
            }

            public bool SaveCurrentCalibrationPoint(BoundWindowInfo? boundWindow = null)
					{
						System.Windows.MessageBox.Show("SaveCurrentCalibrationPoint: Start", "Debug");
						if (_calibrationRunner == null || _calibrationRunner.IsFinished)
						{
							System.Windows.MessageBox.Show("SaveCurrentCalibrationPoint: CalibrationRunner is null or finished, returning false.", "Debug");
							return false;
						}
	
				var currentStep = _calibrationRunner.CurrentStep;
                    if (currentStep == null)
						{
							System.Windows.MessageBox.Show("SaveCurrentCalibrationPoint: CurrentStep is null, returning false.", "Debug");
							return false;
						}
	
                    if (!_hasLastCapturedPosition)
						{
							System.Windows.MessageBox.Show("SaveCurrentCalibrationPoint: HasLastCapturedPosition is false, returning false.", "Debug");
							return false;
						}
	
                    if (string.IsNullOrWhiteSpace(_currentCalibrationModeName))
						{
							System.Windows.MessageBox.Show("SaveCurrentCalibrationPoint: CurrentCalibrationModeName is null or empty, returning false.", "Debug");
							return false;
						}
	

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

                // System.Windows.MessageBox.Show("TryGetCurrentClientCursorPosition: Start", "Debug");
                if (_calibrationService.WindowAutomation == null)
                {
                    // System.Windows.MessageBox.Show("TryGetCurrentClientCursorPosition: WindowAutomation is null", "Debug");
                    return false;
                }

                // Get current cursor position
                System.Drawing.Point screenPoint = _calibrationService.WindowAutomation.GetCursorScreenPosition();
                // System.Windows.MessageBox.Show($"TryGetCurrentClientCursorPosition: Screen Point = ({screenPoint.X}, {screenPoint.Y})", "Debug");

                // Try to get the window under the cursor and its root
                IntPtr childWindow = WindowAutomation.WindowFromPoint(new WindowAutomation.POINT { X = screenPoint.X, Y = screenPoint.Y });
                if (childWindow == IntPtr.Zero)
                {
                    // System.Windows.MessageBox.Show("TryGetCurrentClientCursorPosition: Child window is zero", "Debug");
                    return false;
                }

                IntPtr rootWindow = WindowAutomation.GetAncestor(childWindow, WindowAutomation.GA_ROOT);
                if (rootWindow == IntPtr.Zero)
                {
                    // System.Windows.MessageBox.Show("TryGetCurrentClientCursorPosition: Root window is zero", "Debug");
                    return false;
                }

                // Convert screen coordinates to client coordinates of the root window
                System.Drawing.Point clientPoint = screenPoint;
                WindowAutomation.POINT clientPointWin32 = new WindowAutomation.POINT { X = clientPoint.X, Y = clientPoint.Y };
                if (!WindowAutomation.ScreenToClient(rootWindow, ref clientPointWin32))
                {
                    // System.Windows.MessageBox.Show("TryGetCurrentClientCursorPosition: ScreenToClient failed", "Debug");
                    return false;
                }
                clientPoint = new System.Drawing.Point(clientPointWin32.X, clientPointWin32.Y);

                // If the root window is Copago, bind it
                if (_calibrationService.WindowAutomation.TryBindWindowByHandle(rootWindow, out var copagoWindow))
                {
                    boundCopagoWindow = copagoWindow;
                    // System.Windows.MessageBox.Show($"TryGetCurrentClientCursorPosition: Bound Copago Window = {copagoWindow.Title}", "Debug");
                }

                x = clientPoint.X;
                y = clientPoint.Y;
                // System.Windows.MessageBox.Show($"TryGetCurrentClientCursorPosition: Client Point = ({x}, {y}), returning true", "Debug");
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
				System.Windows.MessageBox.Show("ResetLastCapture: Resetting captured position.", "Debug");
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
