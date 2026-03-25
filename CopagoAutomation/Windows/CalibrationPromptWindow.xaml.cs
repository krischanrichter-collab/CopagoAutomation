using System;
using System.Windows;
using System.Windows.Input;
using CopagoAutomation.ViewModels;

namespace CopagoAutomation.Windows
{
	public partial class CalibrationPromptWindow : Window, IDisposable
	{
		private readonly MainViewModel _viewModel;

		public string StepTitle
		{
			get => StepTitleText.Text;
			set => StepTitleText.Text = value;
		}

		public string HotkeyText
		{
			get => HotkeyTextBlock.Text;
			set => HotkeyTextBlock.Text = value;
		}

		public string InstructionText
		{
			get => InstructionTextBlock.Text;
			set => InstructionTextBlock.Text = value;
		}







		public bool WasConfirmed { get; private set; }

		public bool WasCancelled { get; private set; }

		public CalibrationPromptWindow(Window owner, MainViewModel viewModel)
		{
            InitializeComponent();
            Owner = owner;
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            // Ensure the window stays on top and doesn't appear in the taskbar
            Topmost = true;
            ShowInTaskbar = false;

            Loaded += CalibrationPromptWindow_Loaded;

            OkButton.Click += OkButton_Click;
            CancelButton.Click += CancelButton_Click;

            ResetWindowState();
		}

		private void CalibrationPromptWindow_Loaded(object sender, RoutedEventArgs e)
		{
            Activate();
		}

		public void ResetWindowState()
		{
			WasConfirmed = false;
			WasCancelled = false;
			ResetCaptureState();
		}

            public void ResetCaptureState()
            {
                StatusTextBlock.Text = "Warte auf Tastenkombination...";
                OkButton.IsEnabled = false;
            }

		public void SetStepInfo(string stepTitle, string instructionText, string hotkeyText)
		{
			StepTitle = stepTitle;
			InstructionText = instructionText;
			HotkeyText = hotkeyText;

			ResetWindowState();
		}

            public void SetCapturedPosition(int x, int y)
            {
                StatusTextBlock.Text = $"Position erfasst: X={x}, Y={y}";
                OkButton.IsEnabled = true;
                OkButton.Focus();
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

			WasConfirmed = true;
			WasCancelled = false;
			DialogResult = true;
			Close();
		}

		private void CancelButton_Click(object sender, RoutedEventArgs e)
		{
			WasConfirmed = false;
			WasCancelled = true;
			DialogResult = false;
			Close();
		}

		public void Dispose()
		{
			// Keine speziellen Ressourcen freizugeben, aber IDisposable wird für 'using' benötigt
		}
	}
}
