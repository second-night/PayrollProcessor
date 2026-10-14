namespace PayrollProcessor
{
    /// <summary>
    /// Rolling 3-calendar-month windows from the employee's latest hire/rehire month
    /// (1/1-3/31, 2/1-4/30, and so on). A window qualifies when hours of service
    /// average at least 130 hours per month. A regular paycheck is spread across its
    /// work days, and only the days inside the window count.
    /// </summary>
    internal static class AcaEligibility
    {
        public const int MonthsPerPeriod = 3;
        public const float MonthlyHourThreshold = 130f;

        public static float HoursOfService(PayrollHistoryPeriod period) =>
            period.TotalHours + period.EstimatedCoachHours;

        /// <summary>
        /// First pay date that can contain work from the latest completed measurement window.
        /// </summary>
        public static DateTime HistoryLoadStart(DateTime asOfDate)
        {
            DateTime asOf = asOfDate.Date;
            DateTime monthEnd = new DateTime(asOf.Year, asOf.Month, 1).AddMonths(1).AddDays(-1);
            DateTime latestCompleteEnd = monthEnd <= asOf
                ? monthEnd
                : new DateTime(asOf.Year, asOf.Month, 1).AddDays(-1);
            DateTime latestCompleteStart = new DateTime(latestCompleteEnd.Year, latestCompleteEnd.Month, 1)
                .AddMonths(-(MonthsPerPeriod - 1));
            return latestCompleteStart.AddDays(-PayPeriodSchedule.DaysPerPayPeriod);
        }

        public static List<AcaEligibilityPeriod> BuildPeriods(PayrollHistoryEmployee employee, DateTime? asOfDate = null,
            IReadOnlyCollection<DateTime>? regularPayDates = null)
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

            IReadOnlyCollection<DateTime> anchors = regularPayDates ?? PayPeriodSchedule.HistoryFilePayDates();
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

                float llcHours = 0f;
                float coachesHours = 0f;
                float estimatedCoachHours = 0f;
                HashSet<DateTime> payDates = new();
                foreach (PayrollHistoryPeriod period in employee.Periods.Values)
                {
                    DateTime payDate = period.PayDate.Date;
                    if (payDate > endLimit)
                    {
                        continue;
                    }

                    float fraction = WorkHourFraction(payDate, hoursStart, hoursEnd, anchors);
                    if (fraction <= 0f)
                    {
                        continue;
                    }

                    float hours = HoursOfService(period) * fraction;
                    if (period.Company == Company.VALLEY_BUS_LLC)
                    {
                        llcHours += hours;
                    }
                    else if (period.Company == Company.VALLEY_BUS_COACHES)
                    {
                        coachesHours += hours;
                    }

                    estimatedCoachHours += period.EstimatedCoachHours * fraction;
                    payDates.Add(payDate);
                }

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
                    EstimatedCoachHours = estimatedCoachHours,
                    MonthlyAverage = monthlyAverage,
                    Qualifies = qualifies,
                    PayDateCount = payDates.Count
                });

                windowStart = windowStart.AddMonths(1);
            }

            return windows;
        }

        /// <summary>
        /// Windows that show the employee has met the full-time threshold as of this pay date:
        /// the latest completed window, when it qualifies, and any still-open window that has
        /// already reached 390 hours.
        /// </summary>
        public static List<AcaEligibilityPeriod> QualifyingWindowsForStatus(PayrollHistoryEmployee employee, DateTime date,
            IReadOnlyCollection<DateTime>? regularPayDates = null)
        {
            DateTime asOf = date.Date;
            List<AcaEligibilityPeriod> periods = BuildPeriods(employee, asOf, regularPayDates);
            List<AcaEligibilityPeriod> qualifying = new();
            AcaEligibilityPeriod? latestComplete = LatestCompleted(periods);
            if (latestComplete != null && latestComplete.Qualifies)
            {
                qualifying.Add(latestComplete);
            }

            qualifying.AddRange(periods.Where(period =>
                !period.IsComplete
                && period.Qualifies
                && period.Start.Date <= asOf
                && asOf <= period.End.Date));
            return qualifying;
        }

        /// <summary>
        /// The latest completed window when it finished under 130 hours per month.
        /// An open window that has already reached the threshold keeps the employee full-time.
        /// Returns null when a downgrade should not be recommended.
        /// </summary>
        public static AcaEligibilityPeriod? CompletedWindowBelowFullTime(PayrollHistoryEmployee employee, DateTime date,
            IReadOnlyCollection<DateTime>? regularPayDates = null)
        {
            if (!employee.CurrentStartDate.HasValue)
            {
                return null;
            }

            DateTime asOf = date.Date;
            IReadOnlyCollection<DateTime> anchors = regularPayDates ?? PayPeriodSchedule.HistoryFilePayDates();
            List<AcaEligibilityPeriod> periods = BuildPeriods(employee, asOf, anchors);
            AcaEligibilityPeriod? latestComplete = LatestCompleted(periods);
            if (latestComplete == null || latestComplete.Qualifies)
            {
                return null;
            }

            if (employee.CurrentStartDate.Value.Date > latestComplete.Start.Date)
            {
                return null;
            }

            bool openWindowQualifies = periods.Any(period =>
                !period.IsComplete
                && period.Qualifies
                && period.Start.Date <= asOf
                && asOf <= period.End.Date);
            if (openWindowQualifies)
            {
                return null;
            }

            if (!HasPriorRegularPayDates(employee, latestComplete, asOf, anchors))
            {
                return null;
            }

            return latestComplete;
        }

        internal static float WorkHourFraction(DateTime payDate, DateTime rangeStart, DateTime rangeEnd,
            IReadOnlyCollection<DateTime> anchors)
        {
            DateTime paid = payDate.Date;
            if (!PayPeriodSchedule.IsRegularPayDate(paid, anchors))
            {
                return paid >= rangeStart.Date && paid <= rangeEnd.Date ? 1f : 0f;
            }

            (DateTime workStart, DateTime workEnd) = PayPeriodSchedule.WorkPeriodForPayDate(paid);
            DateTime overlapStart = workStart > rangeStart.Date ? workStart : rangeStart.Date;
            DateTime overlapEnd = workEnd < rangeEnd.Date ? workEnd : rangeEnd.Date;
            if (overlapStart > overlapEnd)
            {
                return 0f;
            }

            int totalDays = (workEnd - workStart).Days + 1;
            int overlapDays = (overlapEnd - overlapStart).Days + 1;
            return overlapDays / (float)totalDays;
        }

        private static AcaEligibilityPeriod? LatestCompleted(List<AcaEligibilityPeriod> periods) =>
            periods.Where(period => period.IsComplete)
                .OrderByDescending(period => period.End)
                .ThenByDescending(period => period.Start)
                .FirstOrDefault();

        /// <summary>
        /// The paycheck dated <paramref name="asOf"/> is the run in progress, so a missing row
        /// means zero hours rather than a gap in ADP. Every earlier regular paycheck whose work
        /// days touch the window has to be present.
        /// </summary>
        private static bool HasPriorRegularPayDates(PayrollHistoryEmployee employee, AcaEligibilityPeriod window,
            DateTime asOf, IReadOnlyCollection<DateTime> anchors)
        {
            if (anchors.Count == 0)
            {
                return false;
            }

            int expected = 0;
            int actual = 0;
            DateTime earliest = window.Start.Date.AddDays(-PayPeriodSchedule.DaysPerPayPeriod);
            for (DateTime cursor = LatestRegularPayDateOnOrBefore(asOf, anchors);
                cursor >= earliest;
                cursor = cursor.AddDays(-PayPeriodSchedule.DaysPerPayPeriod))
            {
                if (cursor.Date >= asOf.Date
                    || WorkHourFraction(cursor, window.Start, window.End, anchors) <= 0f)
                {
                    continue;
                }

                expected++;
                if (employee.Periods.Keys.Any(key => key.PayDate.Date == cursor.Date))
                {
                    actual++;
                }
            }

            return expected > 0 && actual >= expected;
        }

        private static DateTime LatestRegularPayDateOnOrBefore(DateTime endLimit, IReadOnlyCollection<DateTime> anchors)
        {
            DateTime anchor = anchors.Max().Date;
            int dayGap = (endLimit.Date - anchor).Days;
            int steps = (int)Math.Floor(dayGap / (double)PayPeriodSchedule.DaysPerPayPeriod);
            DateTime cursor = anchor.AddDays(steps * PayPeriodSchedule.DaysPerPayPeriod);
            if (cursor > endLimit.Date)
            {
                cursor = cursor.AddDays(-PayPeriodSchedule.DaysPerPayPeriod);
            }

            return cursor;
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
