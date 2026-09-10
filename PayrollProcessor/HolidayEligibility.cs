using System.Globalization;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    /// <summary>
    /// Adds an 8-hour holiday shift for eligible hourly employees when a paid holiday
    /// falls in the current pay period. Hours come from payroll registers for the last 6
    /// regular (biweekly) pay periods — 12 weeks — including special runs in that window,
    /// plus Estimated Coach Hours from PayrollHistory_*.csv (and inferred coach hours on
    /// special payrolls with no history CSV).
    /// </summary>
    internal sealed class HolidayEligibility
    {
        private const float RequiredHours = 360f;
        private const float HolidayHoursToAward = 8f;
        private const int DaysFromPayDateToPeriodStart = 19;

        public void ApplyIfNeeded(DateTime firstDayWeek2)
        {
            DateTime periodStart = firstDayWeek2.Date.AddDays(-7);
            DateTime periodEnd = firstDayWeek2.Date.AddDays(6);
            DateTime currentPayDate = firstDayWeek2.Date.AddDays(12);
            List<(DateTime Date, string Name)> holidays = HolidaysInRange(periodStart, periodEnd);
            if (holidays.Count == 0)
            {
                return;
            }

            foreach ((DateTime holidayDate, string holidayName) in holidays)
            {
                Log("Paid holiday " + holidayName + " (" + holidayDate.ToString("M/d/yyyy")
                    + ") is in this pay period. Checking holiday pay eligibility.");
            }

            List<(DateTime PayDate, string Path)> lastSixFiles = EmployeePayrollHistory.EnumerateHistoryFiles()
                .Where(file => file.PayDate.Date < currentPayDate.Date)
                .OrderByDescending(file => file.PayDate)
                .Take(PayPeriodSchedule.RegularPayPeriodCount)
                .ToList();
            PayrollHistoryCatalog catalog = LoadPayrollHistory(lastSixFiles, firstDayWeek2, currentPayDate);
            DateTime hiredWithinStart = lastSixFiles.Count > 0
                ? lastSixFiles.Min(file => file.PayDate.Date).AddDays(-DaysFromPayDateToPeriodStart)
                : firstDayWeek2.Date.AddDays(-7 - (PayPeriodSchedule.RegularPayPeriodCount * PayPeriodSchedule.DaysPerPayPeriod));
            DateTime hoursRangeStart = lastSixFiles.Count > 0
                ? lastSixFiles.Min(file => file.PayDate.Date)
                : hiredWithinStart;
            DateTime hoursRangeEnd = lastSixFiles.Count > 0
                ? lastSixFiles.Max(file => file.PayDate.Date)
                : currentPayDate.Date.AddDays(-1);

            int awarded = 0;
            foreach (Employee employee in EmployeeDictionary.Values)
            {
                if (!employee.IsActive() || employee.IsSalaried || EmployeeIdsToIgnore.Contains(employee.IdNumber))
                {
                    continue;
                }

                float compensatedHours = GetCompensatedHours(employee.IdNumber, catalog, hoursRangeStart,
                    hoursRangeEnd, currentPayDate);
                bool hiredRecentlyAsFullTime = IsFullTime(employee)
                    && employee.HireDate.Date >= hiredWithinStart.Date;
                if (compensatedHours < RequiredHours && !hiredRecentlyAsFullTime)
                {
                    if (IsFullTime(employee))
                    {
                        Log("Holiday pay: " + employee.Name + " (" + employee.IdNumber
                            + ") is full-time but does not meet the "
                            + RequiredHours.ToString("0.##", CultureInfo.InvariantCulture)
                            + " hour requirement. Last 6 pay periods: "
                            + compensatedHours.ToString("0.##", CultureInfo.InvariantCulture) + " hours"
                            + (hiredRecentlyAsFullTime
                                ? " (still awarded as a full-time hire within the last 6 pay periods)"
                                : "") + ".");
                    }
                    continue;
                }

                foreach ((DateTime holidayDate, string _) in holidays)
                {
                    AddHolidayShift(employee, holidayDate, firstDayWeek2);
                    awarded++;
                    Log("Holiday pay: " + HolidayHoursToAward.ToString("0.##", CultureInfo.InvariantCulture)
                        + " hours on " + holidayDate.ToString("M/d/yyyy") + " for " + employee.Name
                        + " (" + employee.IdNumber + "). Last 6 pay periods: "
                        + compensatedHours.ToString("0.##", CultureInfo.InvariantCulture) + " hours"
                        + (hiredRecentlyAsFullTime ? " (full-time hire within last 6 pay periods)" : "") + ".");
                }
            }

            Log("Holiday pay awarded to " + awarded + " employee shift(s).");
        }

        private static float GetCompensatedHours(int employeeNumber, PayrollHistoryCatalog catalog,
            DateTime hoursRangeStart, DateTime hoursRangeEnd, DateTime currentPayDate)
        {
            if (!catalog.Employees.TryGetValue(employeeNumber, out PayrollHistoryEmployee? historyEmployee))
            {
                return 0f;
            }

            return historyEmployee.Periods.Values
                .Where(period =>
                {
                    DateTime payDate = period.PayDate.Date;
                    if (payDate >= currentPayDate.Date)
                    {
                        return false;
                    }
                    return payDate >= hoursRangeStart.Date && payDate <= hoursRangeEnd.Date;
                })
                .Sum(period => period.CompensatedHours + period.EstimatedCoachHours);
        }

        private static PayrollHistoryCatalog LoadPayrollHistory(
            List<(DateTime PayDate, string Path)> lastSixFiles, DateTime firstDayWeek2, DateTime currentPayDate)
        {
            PayrollHistoryCatalog catalog = new();
            DateTime rangeStart;
            DateTime rangeEnd;
            if (lastSixFiles.Count > 0)
            {
                rangeStart = lastSixFiles.Min(file => file.PayDate.Date);
                rangeEnd = lastSixFiles.Max(file => file.PayDate.Date);
            }
            else
            {
                rangeStart = firstDayWeek2.Date.AddDays(-7 - (PayPeriodSchedule.RegularPayPeriodCount * PayPeriodSchedule.DaysPerPayPeriod));
                rangeEnd = currentPayDate.Date.AddDays(-1);
            }

            try
            {
                catalog.LoadEmployees();
                catalog.LoadPayrollForDateRange(rangeStart, rangeEnd);
            }
            catch (Exception exception)
            {
                Log("Holiday eligibility could not load payroll history files: " + exception.Message, true);
            }

            return catalog;
        }

        private static void AddHolidayShift(Employee employee, DateTime holidayDate, DateTime firstDayWeek2)
        {
            Shift holidayShift = new(employee.PrimaryCompany, Jobs.HOLIDAY)
            {
                ShiftTime = HolidayHoursToAward,
                Date = holidayDate.Date,
                WeekNumber = holidayDate.Date.CompareTo(firstDayWeek2.Date) < 0 ? 1 : 2,
                ClockIn = new TimeSpan(8, 0, 0),
                ClockOut = new TimeSpan(16, 0, 0)
            };
            employee.Shifts.Add(holidayShift);
        }

        private static bool IsFullTime(Employee employee)
        {
            string category = employee.EmploymentCategory?.Trim() ?? "";
            return category.Equals("FT", StringComparison.OrdinalIgnoreCase)
                || category.Equals("ACAFT", StringComparison.OrdinalIgnoreCase);
        }

        private static List<(DateTime Date, string Name)> HolidaysInRange(DateTime periodStart, DateTime periodEnd)
        {
            List<(DateTime Date, string Name)> holidays = new();
            for (int year = periodStart.Year; year <= periodEnd.Year; year++)
            {
                foreach ((DateTime date, string name) in PaidHolidaysForYear(year))
                {
                    if (date.Date >= periodStart.Date && date.Date <= periodEnd.Date)
                    {
                        holidays.Add((date.Date, name));
                    }
                }
            }

            return holidays;
        }

        private static IEnumerable<(DateTime Date, string Name)> PaidHolidaysForYear(int year)
        {
            yield return (new DateTime(year, 1, 1), "New Years Day");
            yield return (LastMondayOfMonth(year, 5), "Memorial Day");
            yield return (new DateTime(year, 7, 4), "July 4th");
            yield return (NthWeekdayOfMonth(year, 9, DayOfWeek.Monday, 1), "Labor Day");
            yield return (NthWeekdayOfMonth(year, 11, DayOfWeek.Thursday, 4), "Thanksgiving");
            yield return (new DateTime(year, 12, 25), "Christmas Day");
        }

        private static DateTime LastMondayOfMonth(int year, int month)
        {
            DateTime date = new DateTime(year, month, DateTime.DaysInMonth(year, month));
            while (date.DayOfWeek != DayOfWeek.Monday)
            {
                date = date.AddDays(-1);
            }
            return date;
        }

        private static DateTime NthWeekdayOfMonth(int year, int month, DayOfWeek weekday, int nth)
        {
            DateTime date = new DateTime(year, month, 1);
            while (date.DayOfWeek != weekday)
            {
                date = date.AddDays(1);
            }
            return date.AddDays(7 * (nth - 1));
        }
    }
}
