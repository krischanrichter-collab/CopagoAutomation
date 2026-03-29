using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CopagoAutomation.Controls
{
    public partial class SingleDateRuleSelector : UserControl
    {
        public DateTime SelectedDate => SingleDate?.SelectedDate ?? DateTime.Today.AddDays(-1);

        public SingleDateRuleSelector()
        {
            InitializeComponent();
            Loaded += SingleDateRuleSelector_Loaded;
        }

        private void SingleDateRuleSelector_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= SingleDateRuleSelector_Loaded;

            ModeOneTime.IsChecked = true;

            SingleDate.SelectedDate = DateTime.Today.AddDays(-1);

            ValidFrom.SelectedDate  = DateTime.Today;
            ValidTo.SelectedDate    = DateTime.Today.AddMonths(3);

            RuleType.SelectedIndex      = 0;
            WD_Di.IsChecked             = true;
            MonthOrdinal.SelectedIndex  = 0;
            MonthWeekday.SelectedIndex  = 1;

            RefreshModePanels();
            RefreshRulePanels();
        }

        private void Mode_Checked(object sender, RoutedEventArgs e) => RefreshModePanels();

        private void RuleType_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshRulePanels();

        private void RefreshModePanels()
        {
            bool recurring = ModeRecurring?.IsChecked == true;
            if (OneTimePanel   != null) OneTimePanel.Visibility   = recurring ? Visibility.Collapsed : Visibility.Visible;
            if (RecurringPanel != null) RecurringPanel.Visibility = recurring ? Visibility.Visible   : Visibility.Collapsed;
        }

        private void RefreshRulePanels()
        {
            bool monthly = RuleType?.SelectedIndex == 1;
            if (WeeklyPanel  != null) WeeklyPanel.Visibility  = monthly ? Visibility.Collapsed : Visibility.Visible;
            if (MonthlyPanel != null) MonthlyPanel.Visibility = monthly ? Visibility.Visible   : Visibility.Collapsed;
        }

        /// <summary>
        /// Gibt alle relevanten Datumseinträge zurück:
        /// - Einmaliges Datum: Liste mit einem Eintrag
        /// - Wiederkehrend: alle generierten Vorkommen
        /// </summary>
        public List<DateTime> GetDates(int maxOccurrences = 200)
        {
            if (ModeOneTime?.IsChecked == true)
            {
                var date = SingleDate?.SelectedDate ?? DateTime.Today.AddDays(-1);
                return new List<DateTime> { date.Date };
            }

            return GetRecurrenceRule()?.GenerateDates().Take(maxOccurrences).ToList()
                   ?? new List<DateTime>();
        }

        private RecurrenceRule? GetRecurrenceRule()
        {
            if (ValidFrom?.SelectedDate is not DateTime vf || ValidTo?.SelectedDate is not DateTime vt)
                return null;

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
                    Kind       = RecurrenceKind.Weekly,
                    ValidFrom  = vf.Date,
                    ValidTo    = vt.Date,
                    WeeklyDays = days
                };
            }
            else
            {
                int ordinal = MonthOrdinal?.SelectedIndex switch
                {
                    0 => 1, 1 => 2, 2 => 3, 3 => 4, _ => -1
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
                    Kind             = RecurrenceKind.MonthlyNthWeekday,
                    ValidFrom        = vf.Date,
                    ValidTo          = vt.Date,
                    MonthlyOrdinal   = ordinal,
                    MonthlyWeekday   = weekday
                };
            }
        }
    }
}
