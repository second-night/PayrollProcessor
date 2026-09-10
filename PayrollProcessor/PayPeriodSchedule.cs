namespace PayrollProcessor
{
    /// <summary>
    /// Regular payroll is biweekly (12 weeks for 6 pay periods). Off-cycle / special runs
    /// sit on other dates and are included in reports without consuming a regular slot.
    /// </summary>
    internal static class PayPeriodSchedule
    {
        public const int RegularPayPeriodCount = 6;
        public const int DaysPerPayPeriod = 14;
        public const int WeeksPerPayPeriod = 2;

        public static List<DateTime> HistoryFilePayDates() =>
            EmployeePayrollHistory.EnumerateHistoryFiles()
                .Select(file => file.PayDate.Date)
                .Distinct()
                .OrderBy(date => date)
                .ToList();

        public static bool IsRegularPayDate(DateTime payDate, IReadOnlyCollection<DateTime>? regularAnchors = null)
        {
            DateTime date = payDate.Date;
            IReadOnlyCollection<DateTime> anchors = regularAnchors ?? HistoryFilePayDates();
            if (anchors.Count > 0)
            {
                return Math.Abs((date - anchors.Max()).Days) % DaysPerPayPeriod == 0;
            }

            return date.DayOfWeek == DayOfWeek.Friday;
        }

        public static (DateTime Start, DateTime End)? LastRegularPayPeriodWindow(DateTime? beforeDate = null)
        {
            List<DateTime> regular = HistoryFilePayDates()
                .Where(date => !beforeDate.HasValue || date < beforeDate.Value.Date)
                .OrderByDescending(date => date)
                .Take(RegularPayPeriodCount)
                .ToList();
            if (regular.Count == 0)
            {
                return null;
            }

            return (regular.Min(), regular.Max());
        }

        public static bool IsInLastRegularPayPeriodWindow(DateTime payDate, DateTime? beforeDate = null)
        {
            (DateTime Start, DateTime End)? window = LastRegularPayPeriodWindow(beforeDate);
            if (!window.HasValue)
            {
                return false;
            }

            DateTime date = payDate.Date;
            return date >= window.Value.Start && date <= window.Value.End;
        }

        public static int RegularPayDateCount(IEnumerable<DateTime> payDates,
            IReadOnlyCollection<DateTime>? regularAnchors = null)
        {
            IReadOnlyCollection<DateTime> anchors = regularAnchors ?? HistoryFilePayDates();
            return payDates.Select(date => date.Date).Distinct().Count(date => IsRegularPayDate(date, anchors));
        }
    }
}
