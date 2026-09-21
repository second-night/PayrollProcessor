namespace PayrollProcessor
{
    /// <summary>
    /// Rolling 3-calendar-month windows from the employee's latest hire/rehire month
    /// (1/1-3/31, 2/1-4/30, and so on). A window qualifies when total hours of service
    /// (including estimated coach hours) average at least 130 hours per month.
    /// </summary>
    internal static class AcaEligibility
    {
        public const int MonthsPerPeriod = 3;
        public const float MonthlyHourThreshold = 130f;

        public static float HoursOfService(PayrollHistoryPeriod period) =>
            period.TotalHours + period.EstimatedCoachHours;

        public static List<AcaEligibilityPeriod> BuildPeriods(PayrollHistoryEmployee employee, DateTime? asOfDate = null)
        {
            List<AcaEligibilityPeriod> windows = new();
            if (!employee.CurrentStartDate.HasValue)
            {
                return windows;
            }

            DateTime hireDate = employee.CurrentStartDate.Value.Date;
            DateTime endLimit = (asOfDate ?? DateTime.Today).Date;
            if (employee.TerminationDate.HasValue && employee.TerminationDate.Value.Date < endLimit)
            {
                endLimit = employee.TerminationDate.Value.Date;
            }

            DateTime windowStart = new(hireDate.Year, hireDate.Month, 1);
            while (windowStart <= endLimit)
            {
                DateTime windowEnd = windowStart.AddMonths(MonthsPerPeriod).AddDays(-1);
                DateTime hoursStart = windowStart < hireDate ? hireDate : windowStart;
                DateTime hoursEnd = windowEnd < endLimit ? windowEnd : endLimit;
                bool complete = windowEnd <= endLimit;
                if (hoursStart > hoursEnd)
                {
                    windowStart = windowStart.AddMonths(1);
                    continue;
                }

                List<PayrollHistoryPeriod> inWindow = employee.Periods.Values
                    .Where(period => period.PayDate.Date >= hoursStart && period.PayDate.Date <= hoursEnd)
                    .OrderBy(period => period.PayDate)
                    .ThenBy(period => period.Company)
                    .ToList();

                float llcHours = inWindow
                    .Where(period => period.Company == Company.VALLEY_BUS_LLC)
                    .Sum(HoursOfService);
                float coachesHours = inWindow
                    .Where(period => period.Company == Company.VALLEY_BUS_COACHES)
                    .Sum(HoursOfService);
                float totalHours = llcHours + coachesHours;
                float monthlyAverage = totalHours / MonthsPerPeriod;
                bool qualifies = monthlyAverage + 0.005f >= MonthlyHourThreshold;
                windows.Add(new AcaEligibilityPeriod
                {
                    Start = windowStart,
                    End = windowEnd,
                    IsComplete = complete,
                    TotalHours = totalHours,
                    LlcHours = llcHours,
                    CoachesHours = coachesHours,
                    EstimatedCoachHours = inWindow.Sum(period => period.EstimatedCoachHours),
                    MonthlyAverage = monthlyAverage,
                    Qualifies = qualifies,
                    PayDateCount = inWindow.Select(period => period.PayDate.Date).Distinct().Count()
                });

                windowStart = windowStart.AddMonths(1);
            }

            return windows;
        }

        public static List<AcaEligibilityPeriod> QualifyingWindowsCovering(PayrollHistoryEmployee employee, DateTime date)
        {
            DateTime asOf = date.Date;
            return BuildPeriods(employee, asOf)
                .Where(period => period.Qualifies && period.Start <= asOf && asOf <= period.End)
                .ToList();
        }
    }

    internal sealed class AcaEligibilityPeriod
    {
        public DateTime Start { get; init; }
        public DateTime End { get; init; }
        public bool IsComplete { get; init; }
        public float TotalHours { get; init; }
        public float LlcHours { get; init; }
        public float CoachesHours { get; init; }
        public float EstimatedCoachHours { get; init; }
        public float MonthlyAverage { get; init; }
        public bool Qualifies { get; init; }
        public int PayDateCount { get; init; }
    }
}
