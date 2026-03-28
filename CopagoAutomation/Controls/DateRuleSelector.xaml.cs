using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CopagoAutomation.Controls
{
		public partial class DateRuleSelector : UserControl
		{
			public DateTime DateFrom => OneTimeFrom?.SelectedDate ?? DateTime.Today.AddDays(-1);
			public DateTime DateTo => OneTimeTo?.SelectedDate ?? DateTime.Today.AddDays(-1);

			public DateRuleSelector()
			{
				InitializeComponent();

			// Wichtig: Defaults erst setzen, nachdem alles existiert.
			Loaded += DateRuleSelector_Loaded;
		}

		private void DateRuleSelector_Loaded(object sender, RoutedEventArgs e)
		{
			// Loaded nur 1x ausführen
			Loaded -= DateRuleSelector_Loaded;

			// --- Defaults ---
			ModeOneTime.IsChecked = true; // default Modus

			OneTimeFrom.SelectedDate = DateTime.Today.AddDays(-1);
			OneTimeTo.SelectedDate = DateTime.Today.AddDays(-1);

			ValidFrom.SelectedDate = DateTime.Today;
			ValidTo.SelectedDate = DateTime.Today.AddMonths(3);

			RuleType.SelectedIndex = 0; // wöchentlich

			WD_Di.IsChecked = true;     // default: Dienstag

			MonthOrdinal.SelectedIndex = 0;  // 1.
			MonthWeekday.SelectedIndex = 1;  // Dienstag

			// UI initial korrekt setzen
			RefreshModePanels();
			RefreshRulePanels();
		}

		private void Mode_Checked(object sender, RoutedEventArgs e)
		{
			RefreshModePanels();
		}

		private void RuleType_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			RefreshRulePanels();
		}

		private void RefreshModePanels()
		{
			// null-safe: falls jemand doch im Designer oder früh Events triggert
			bool recurring = ModeRecurring?.IsChecked == true;

			if (OneTimePanel != null)
				OneTimePanel.Visibility = recurring ? Visibility.Collapsed : Visibility.Visible;

			if (RecurringPanel != null)
				RecurringPanel.Visibility = recurring ? Visibility.Visible : Visibility.Collapsed;
		}

		private void RefreshRulePanels()
		{
			// null-safe
			bool monthly = RuleType?.SelectedIndex == 1;

			if (WeeklyPanel != null)
				WeeklyPanel.Visibility = monthly ? Visibility.Collapsed : Visibility.Visible;

			if (MonthlyPanel != null)
				MonthlyPanel.Visibility = monthly ? Visibility.Visible : Visibility.Collapsed;
		}

		// ---- Ergebnis-API: liefert entweder (From,To) oder eine Regel ----
		public (DateTime from, DateTime to)? GetOneTimeRange()
		{
			if (ModeOneTime?.IsChecked != true) return null;
			if (OneTimeFrom?.SelectedDate is not DateTime f || OneTimeTo?.SelectedDate is not DateTime t) return null;
			return (f.Date, t.Date);
		}

		public RecurrenceRule? GetRecurrenceRule()
		{
			if (ModeRecurring?.IsChecked != true) return null;
			if (ValidFrom?.SelectedDate is not DateTime vf || ValidTo?.SelectedDate is not DateTime vt) return null;

			if (RuleType?.SelectedIndex == 0)
			{
				var days = new List<DayOfWeek>();
				if (WD_Mo?.IsChecked == true) days.Add(DayOfWeek.Monday);
				if (WD_Di?.IsChecked == true) days.Add(DayOfWeek.Tuesday);
				if (WD_Mi?.IsChecked == true) days.Add(DayOfWeek.Wednesday);
				if (WD_Do?.IsChecked == true) days.Add(DayOfWeek.Thursday);
				if (WD_Fr?.IsChecked == true) days.Add(DayOfWeek.Friday);
				if (WD_Sa?.IsChecked == true) days.Add(DayOfWeek.Saturday);
				if (WD_So?.IsChecked == true) days.Add(DayOfWeek.Sunday);

				return new RecurrenceRule
				{
					Kind = RecurrenceKind.Weekly,
					ValidFrom = vf.Date,
					ValidTo = vt.Date,
					WeeklyDays = days
				};
			}
			else
			{
				int ordinal = MonthOrdinal?.SelectedIndex switch
				{
					0 => 1,
					1 => 2,
					2 => 3,
					3 => 4,
					_ => -1 // Letzter
				};

				DayOfWeek weekday = MonthWeekday?.SelectedIndex switch
				{
					0 => DayOfWeek.Monday,
					1 => DayOfWeek.Tuesday,
					2 => DayOfWeek.Wednesday,
					3 => DayOfWeek.Thursday,
					4 => DayOfWeek.Friday,
					5 => DayOfWeek.Saturday,
					_ => DayOfWeek.Sunday
				};

				return new RecurrenceRule
				{
					Kind = RecurrenceKind.MonthlyNthWeekday,
					ValidFrom = vf.Date,
					ValidTo = vt.Date,
					MonthlyOrdinal = ordinal,
					MonthlyWeekday = weekday
				};
			}
		}

		// Hilfsfunktion: erzeugt konkrete „Run-Daten“ (z.B. alle Treffer im Gültigkeitsbereich)
		public List<DateTime> GenerateOccurrences(int max = 200)
		{
			var rule = GetRecurrenceRule();
			if (rule is null) return new List<DateTime>();

			return rule.GenerateDates().Take(max).ToList();
		}
	}

	public enum RecurrenceKind
	{
		Weekly,
		MonthlyNthWeekday
	}

	public class RecurrenceRule
	{
		public RecurrenceKind Kind { get; set; }
		public DateTime ValidFrom { get; set; }
		public DateTime ValidTo { get; set; }

		// Weekly
		public List<DayOfWeek> WeeklyDays { get; set; } = new();

		// Monthly (nth weekday); MonthlyOrdinal: 1..4 or -1 for last
		public int MonthlyOrdinal { get; set; }
		public DayOfWeek MonthlyWeekday { get; set; }

		public IEnumerable<DateTime> GenerateDates()
		{
			if (ValidTo < ValidFrom) yield break;

			if (Kind == RecurrenceKind.Weekly)
			{
				if (WeeklyDays.Count == 0) yield break;

				for (var d = ValidFrom.Date; d <= ValidTo.Date; d = d.AddDays(1))
				{
					if (WeeklyDays.Contains(d.DayOfWeek))
						yield return d;
				}
			}
			else
			{
				var monthCursor = new DateTime(ValidFrom.Year, ValidFrom.Month, 1);
				var endMonth = new DateTime(ValidTo.Year, ValidTo.Month, 1);

				while (monthCursor <= endMonth)
				{
					var hit = GetNthWeekdayOfMonth(monthCursor.Year, monthCursor.Month, MonthlyWeekday, MonthlyOrdinal);
					if (hit >= ValidFrom.Date && hit <= ValidTo.Date)
						yield return hit;

					monthCursor = monthCursor.AddMonths(1);
				}
			}
		}

		private static DateTime GetNthWeekdayOfMonth(int year, int month, DayOfWeek weekday, int ordinal)
		{
			if (ordinal == -1)
			{
				var lastDay = new DateTime(year, month, DateTime.DaysInMonth(year, month));
				while (lastDay.DayOfWeek != weekday)
					lastDay = lastDay.AddDays(-1);
				return lastDay.Date;
			}

			var first = new DateTime(year, month, 1);
			while (first.DayOfWeek != weekday)
				first = first.AddDays(1);

			var result = first.AddDays((ordinal - 1) * 7);

			if (result.Month != month)
				return GetNthWeekdayOfMonth(year, month, weekday, -1);

			return result.Date;
		}
	}
}