using System.Globalization;
using System.Text;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    /// <summary>
    /// Persists the payroll totals used for employment-status checks.  Each file represents
    /// one pay date, so rerunning a primary payroll replaces only that payroll's history.
    /// </summary>
    public sealed class EmployeePayrollHistory
    {
        public sealed record Entry(int EmployeeNumber, Company Company, float TotalHours, float TotalCompensation,
            Dictionary<Jobs, float> PayRates, Dictionary<Jobs, float> HoursByJob, float GrossPay);

        private readonly Dictionary<int, DateTime> lastHoursDateByEmployee = new();
        private readonly Dictionary<(int EmployeeNumber, Jobs Job), float> latestPayRates = new();
        private readonly Dictionary<int, Dictionary<Jobs, float>> hoursByJobFromHistory = new();
        private readonly List<Dictionary<int, float>> previousPayPeriodHoursNewestFirst = new();
        private DateTime currentPayDate;
        private DateTime? earliestHistoryPayDate;
        private const int MaxTerminationLookbackPayPeriods = 8;
        private static readonly Jobs[] TrackedPayRateJobs =
        {
            Jobs.MECHANIC, Jobs.WASH_BAY, Jobs.ADMIN, Jobs.CLEANING, Jobs.SALARY
        };

        public HashSet<int> PartTimeEmployeesNeedingFullTimeStatus { get; } = new();
        public HashSet<int> FullTimeEmployeesNeedingPartTimeStatus { get; } = new();
        public HashSet<int> EmployeesNeedingTermination { get; } = new();
        public HashSet<int> EmployeesNeedingTerminationInNonPrimaryCompanyOnly { get; } = new();
        public Dictionary<int, List<Company>> EmployeesNeedingRehire { get; } = new();

        public void LoadPreviousPayPeriods(DateTime currentPayDate)
        {
            this.currentPayDate = currentPayDate;
            previousPayPeriodHoursNewestFirst.Clear();
            List<(DateTime PayDate, string Path)> allFiles = GetHistoryFiles()
                .Where(file => file.PayDate.Date < currentPayDate.Date)
                .OrderByDescending(file => file.PayDate)
                .ToList();
            Dictionary<string, List<Entry>> entriesByPath = new();
            foreach ((DateTime payDate, string path) in allFiles)
            {
                if (TryReadEntries(path, payDate, out List<Entry> entries, out _))
                {
                    entriesByPath[path] = entries;
                }
            }

            foreach ((DateTime payDate, string path) in allFiles)
            {
                if (!entriesByPath.TryGetValue(path, out List<Entry>? entries))
                {
                    Program.Log($"Payroll history file could not be fully loaded: {path}", true);
                    continue;
                }

                earliestHistoryPayDate = !earliestHistoryPayDate.HasValue || payDate < earliestHistoryPayDate
                    ? payDate : earliestHistoryPayDate;
                foreach (Entry entry in entries)
                {
                    foreach ((Jobs job, float rate) in entry.PayRates)
                    {
                        latestPayRates.TryAdd((entry.EmployeeNumber, job), rate);
                    }

                    if (!hoursByJobFromHistory.TryGetValue(entry.EmployeeNumber, out Dictionary<Jobs, float>? hoursByJob))
                    {
                        hoursByJob = new();
                        hoursByJobFromHistory[entry.EmployeeNumber] = hoursByJob;
                    }
                    foreach ((Jobs job, float hours) in entry.HoursByJob)
                    {
                        if (hours > 0.01f)
                        {
                            hoursByJob[job] = hoursByJob.GetValueOrDefault(job) + hours;
                        }
                    }

                    if (entry.TotalHours > 0.01f)
                    {
                        lastHoursDateByEmployee[entry.EmployeeNumber] =
                            !lastHoursDateByEmployee.TryGetValue(entry.EmployeeNumber, out DateTime latest) || payDate > latest
                                ? payDate : latest;
                    }
                }
            }

            foreach ((DateTime _, string path) in allFiles.Take(MaxTerminationLookbackPayPeriods))
            {
                if (!entriesByPath.TryGetValue(path, out List<Entry>? entries))
                {
                    break;
                }

                Dictionary<int, float> hoursByEmployee = new();
                foreach (Entry entry in entries)
                {
                    if (entry.TotalHours > 0.01f)
                    {
                        hoursByEmployee[entry.EmployeeNumber] =
                            hoursByEmployee.GetValueOrDefault(entry.EmployeeNumber) + entry.TotalHours;
                    }
                }
                previousPayPeriodHoursNewestFirst.Add(hoursByEmployee);
            }
        }

        public void EvaluateEmployees(IEnumerable<Employee> employees)
        {
            List<Employee> employeeList = employees.ToList();
            PayrollHistoryCatalog? catalog = TryLoadCatalogForStatusChecks();
            CheckPartTimeAcaEligibility(employeeList, catalog);
            foreach (Employee employee in employeeList)
            {
                if (employee.IdNumber == 503)
                {
                    //John McLaughlin exception
                    continue;
                }
                bool hasCurrentHours = GetCurrentHours(employee) > 0.01f;

                HashSet<int> terminationExceptions = new() {105, 187, 501, 503};
                int lookbackPayPeriods = GetTerminationLookbackPayPeriods(employee);
                if (previousPayPeriodHoursNewestFirst.Count >= lookbackPayPeriods
                    && !employee.IsTerminated 
                    && !employee.IsSalaried 
                    && !hasCurrentHours
                    && employee.WasAlreadyInPayroll
                    && !WasHiredInLast30Days(employee)
                    && GetHistoricalHours(employee.IdNumber, lookbackPayPeriods) < 0.01f 
                    && !terminationExceptions.Contains(employee.IdNumber))
                {
                    EmployeesNeedingTermination.Add(employee.IdNumber);
                }

                DateTime payPeriodStart = PayPeriodSchedule.WorkPeriodForPayDate(currentPayDate).Start;
                List<Company> companiesToRehire = SelectCompaniesToRehire(employee, payPeriodStart);
                if (companiesToRehire.Count > 0)
                {
                    if (ShouldConfirmSuspiciousRehire(employee)
                        && !ConfirmSuspiciousRehire(employee))
                    {
                        employee.DiscardPayPeriodHours();
                        Program.Log("Discarded the only shift for " + employee.Name
                            + " (" + employee.IdNumber + ") and did not rehire them.");
                        continue;
                    }
                    EmployeesNeedingRehire[employee.IdNumber] = companiesToRehire;
                }
                else if (employee.IsTerminated
                    && !HasCurrentPayInAnyCompany(employee)
                    && employee.ActiveCompanies.Any(company => company != employee.PrimaryCompany)
                    && !terminationExceptions.Contains(employee.IdNumber))
                {
                    EmployeesNeedingTerminationInNonPrimaryCompanyOnly.Add(employee.IdNumber);
                }
            }
        }

        private DateTime AcaHistoryRangeStart => AcaEligibility.HistoryLoadStart(currentPayDate);

        private PayrollHistoryCatalog? TryLoadCatalogForStatusChecks()
        {
            DateTime rangeStart = AcaHistoryRangeStart;
            try
            {
                PayrollHistoryCatalog catalog = new();
                catalog.LoadEmployees();
                catalog.LoadPayrollForDateRange(rangeStart, currentPayDate);
                string? missingAdpError = catalog.GetMissingAdpPayrollError(rangeStart, currentPayDate);
                if (missingAdpError != null)
                {
                    Program.Log(missingAdpError, true);
                }
                return catalog;
            }
            catch (Exception exception)
            {
                Program.Log("Could not load payroll history for employment-status checks: " + exception.Message, true);
                return null;
            }
        }

        private void CheckPartTimeAcaEligibility(IEnumerable<Employee> employees, PayrollHistoryCatalog? catalog)
        {
            if (catalog == null)
            {
                return;
            }

            List<Employee> employeeList = employees.ToList();
            List<Employee> statusEmployees = employeeList
                .Where(employee => employee.IdNumber != 503
                    && !employee.IsTerminated
                    && !employee.IsSalaried
                    && (WfnEmployeesReader.IsPartTimeCategory(employee.EmploymentCategory)
                        || WfnEmployeesReader.IsFullTimeCategory(employee.EmploymentCategory)))
                .ToList();
            List<Employee> vacationCheckEmployees = employeeList
                .Where(employee => employee.bShouldBeCheckedForVacation)
                .ToList();
            if (statusEmployees.Count == 0 && vacationCheckEmployees.Count == 0)
            {
                return;
            }

            OverlayCurrentPayPeriodHours(catalog, statusEmployees.Concat(vacationCheckEmployees)
                .GroupBy(employee => employee.IdNumber)
                .Select(group => group.First()));

            IReadOnlyCollection<DateTime> regularPayDates = PayPeriodSchedule.HistoryFilePayDates();
            List<string> fullTimeLines = new();
            List<string> partTimeLines = new();
            foreach (Employee employee in statusEmployees.OrderBy(employee => employee.LastName)
                .ThenBy(employee => employee.FirstName))
            {
                if (!catalog.Employees.TryGetValue(employee.IdNumber, out PayrollHistoryEmployee? history))
                {
                    continue;
                }
                if (!history.CurrentStartDate.HasValue && employee.HireDate != DateTime.MinValue)
                {
                    history.HireDate = employee.HireDate.Date;
                }

                if (WfnEmployeesReader.IsPartTimeCategory(employee.EmploymentCategory))
                {
                    List<AcaEligibilityPeriod> qualifying = AcaEligibility.QualifyingWindowsForStatus(
                        history, currentPayDate, regularPayDates);
                    if (qualifying.Count == 0)
                    {
                        continue;
                    }

                    PartTimeEmployeesNeedingFullTimeStatus.Add(employee.IdNumber);
                    fullTimeLines.Add(employee.Name + " (" + employee.IdNumber + "): "
                        + string.Join("; ", qualifying.Select(FormatMeasurementWindow)));
                }
                else if (WfnEmployeesReader.IsFullTimeCategory(employee.EmploymentCategory))
                {
                    AcaEligibilityPeriod? belowFullTime = AcaEligibility.CompletedWindowBelowFullTime(
                        history, currentPayDate, regularPayDates);
                    if (belowFullTime == null)
                    {
                        continue;
                    }

                    FullTimeEmployeesNeedingPartTimeStatus.Add(employee.IdNumber);
                    partTimeLines.Add(employee.Name + " (" + employee.IdNumber + "): "
                        + FormatMeasurementWindow(belowFullTime));
                }
            }

            if (fullTimeLines.Count > 0)
            {
                Program.DelayedLog("Part-time employees who have met the ACA full-time hour threshold "
                    + "(130 hours/month over a rolling 3-month period). Review in ADP; this was not imported "
                    + "as a status change." + Environment.NewLine + Environment.NewLine
                    + string.Join(Environment.NewLine, fullTimeLines));
            }

            if (partTimeLines.Count > 0)
            {
                Program.DelayedLog("Full-time employees whose latest completed 3-month period averaged under "
                    + "130 hours per month. Review in ADP; this was not imported as a status change."
                    + Environment.NewLine + Environment.NewLine
                    + string.Join(Environment.NewLine, partTimeLines));
            }

            LogVacationMeasurementHours(vacationCheckEmployees, catalog);
        }

        private void LogVacationMeasurementHours(IEnumerable<Employee> employees, PayrollHistoryCatalog catalog)
        {
            List<string> lines = new();
            foreach (Employee employee in employees.OrderBy(employee => employee.LastName)
                .ThenBy(employee => employee.FirstName))
            {
                string measurement = DescribeMeasurementPeriodHours(employee, catalog);
                lines.Add(employee.Name + " (" + employee.IdNumber + "): " + measurement);
            }

            if (lines.Count == 0)
            {
                return;
            }

            Program.DelayedLog("Part-time employees flagged for a vacation-policy check. Hours over the current "
                + "3-month measurement period:" + Environment.NewLine + Environment.NewLine
                + string.Join(Environment.NewLine, lines));
        }

        private string DescribeMeasurementPeriodHours(Employee employee, PayrollHistoryCatalog catalog)
        {
            if (!catalog.Employees.TryGetValue(employee.IdNumber, out PayrollHistoryEmployee? history))
            {
                return GetCurrentHours(employee).ToString("0.0", CultureInfo.InvariantCulture)
                    + " hours (current pay period only; no payroll history)";
            }

            if (!history.CurrentStartDate.HasValue && employee.HireDate != DateTime.MinValue)
            {
                history.HireDate = employee.HireDate.Date;
            }

            DateTime measurementStart = new DateTime(currentPayDate.Year, currentPayDate.Month, 1)
                .AddMonths(-(AcaEligibility.MonthsPerPeriod - 1));
            List<AcaEligibilityPeriod> periods = AcaEligibility.BuildPeriods(history, currentPayDate);
            AcaEligibilityPeriod? period = periods.FirstOrDefault(candidate => candidate.Start.Date == measurementStart)
                ?? periods.OrderBy(candidate => candidate.Start).FirstOrDefault();
            if (period == null)
            {
                float hours = history.Periods.Values
                    .Where(payPeriod => payPeriod.PayDate.Date >= measurementStart
                        && payPeriod.PayDate.Date <= currentPayDate.Date)
                    .Sum(AcaEligibility.HoursOfService);
                return hours.ToString("0.0", CultureInfo.InvariantCulture) + " hours ("
                    + measurementStart.ToString("M/d/yyyy", CultureInfo.InvariantCulture) + "-"
                    + currentPayDate.ToString("M/d/yyyy", CultureInfo.InvariantCulture) + ")";
            }

            return period.TotalHours.ToString("0.0", CultureInfo.InvariantCulture) + " hours ("
                + period.Start.ToString("M/d/yyyy", CultureInfo.InvariantCulture) + "-"
                + period.End.ToString("M/d/yyyy", CultureInfo.InvariantCulture) + ")";
        }

        private void OverlayCurrentPayPeriodHours(PayrollHistoryCatalog catalog, IEnumerable<Employee> employees)
        {
            foreach (Employee employee in employees)
            {
                if (!catalog.Employees.TryGetValue(employee.IdNumber, out PayrollHistoryEmployee? history))
                {
                    continue;
                }

                Dictionary<Company, List<Shift>> shiftsByCompany = employee.Shifts
                    .Where(shift => !shift.IsATotalsShift)
                    .GroupBy(shift => shift.CompanyName)
                    .ToDictionary(group => group.Key, group => group.ToList());
                Dictionary<Company, List<ManualEntry>> manualEntriesByCompany = employee.ManualEntries
                    .GroupBy(entry => entry.Company)
                    .ToDictionary(group => group.Key, group => group.ToList());
                HashSet<Company> companies = new(shiftsByCompany.Keys);
                companies.UnionWith(manualEntriesByCompany.Keys);

                foreach (Company company in companies)
                {
                    List<Shift> shifts = shiftsByCompany.GetValueOrDefault(company) ?? new List<Shift>();
                    List<ManualEntry> manualEntries = manualEntriesByCompany.GetValueOrDefault(company)
                        ?? new List<ManualEntry>();
                    float totalHours = shifts.Sum(shift => shift.AllHours(false))
                        + manualEntries.Sum(entry => entry.AllHours());
                    float estimatedCoachHours = shifts
                        .Where(shift => shift.JobType == Jobs.DRIVER_COACH)
                        .Sum(shift => shift.CoachTripDays * 8f);
                    if (shifts.Count == 0 && totalHours <= 0.01f && estimatedCoachHours <= 0.01f)
                    {
                        continue;
                    }

                    (DateTime PayDate, Company Company) key = (currentPayDate.Date, company);
                    history.Periods[key] = new PayrollHistoryPeriod
                    {
                        PayDate = currentPayDate.Date,
                        Company = company,
                        Source = "Current",
                        TotalHours = totalHours,
                        EstimatedCoachHours = estimatedCoachHours
                    };
                }
            }
        }

        private int GetTerminationLookbackPayPeriods(Employee employee)
        {
            if (WfnEmployeesReader.IsFullTimeCategory(employee.EmploymentCategory))
            {
                return 3;
            }

            return currentPayDate.Month switch
            {
                7 => 7,
                8 or 9 => 8,
                _ => 6
            };
        }

        private static string FormatMeasurementWindow(AcaEligibilityPeriod period) =>
            period.Start.ToString("M/d/yyyy", CultureInfo.InvariantCulture) + "-"
            + period.End.ToString("M/d/yyyy", CultureInfo.InvariantCulture) + " avg "
            + period.MonthlyAverage.ToString("0.0", CultureInfo.InvariantCulture) + " hrs/mo";

        private bool WasHiredInLast30Days(Employee employee)
        {
            if (employee.HireDate == DateTime.MinValue)
            {
                return false;
            }

            return employee.HireDate.Date >= currentPayDate.Date.AddDays(-30);
        }

        private float GetHistoricalHours(int employeeNumber, int lookbackPayPeriods)
        {
            float total = 0f;
            for (int index = 0; index < lookbackPayPeriods; index++)
            {
                total += previousPayPeriodHoursNewestFirst[index].GetValueOrDefault(employeeNumber);
            }
            return total;
        }

        public IReadOnlyDictionary<Jobs, float> GetHoursByJobFromHistory(int employeeNumber) =>
            hoursByJobFromHistory.TryGetValue(employeeNumber, out Dictionary<Jobs, float>? hoursByJob)
                ? hoursByJob
                : new Dictionary<Jobs, float>();

        public DateTime GetTerminationDate(int employeeNumber, DateTime fallbackPayDate)
        {
            if (lastHoursDateByEmployee.TryGetValue(employeeNumber, out DateTime lastHoursDate))
            {
                return lastHoursDate;
            }

            return earliestHistoryPayDate ?? fallbackPayDate;
        }

        public static bool ShouldIncludeSalary(Employee employee) =>
            employee.AnnualSalaryAmount > 0.001f && !employee.IsTerminated;

        public static float GetPerPayPeriodSalary(Employee employee) =>
            (float)Math.Round(employee.AnnualSalaryAmount / 26f, 2);

        public static float GetGrossPay(Employee employee)
        {
            float grossPay = employee.Shifts
                .Where(shift => !shift.IsATotalsShift)
                .Sum(shift => shift.TotalCompensation(employee));
            if (ShouldIncludeSalary(employee))
            {
                grossPay += GetPerPayPeriodSalary(employee);
            }
            return grossPay;
        }

        public void WriteCurrentPayPeriod(DateTime payDate, IEnumerable<Employee> employees)
        {
            string path = Path.Combine(HistoryDirectory, $"PayrollHistory_{payDate:yyyy-MM-dd}.csv");
            List<string> headers = new()
            {
                "Payroll Date", "Employee Number", "Company", "Estimated Coach Hours"
            };
            headers.Add("Record Count");
            headers.AddRange(Enum.GetNames<Jobs>().Select(job => $"{job} Hours"));
            headers.AddRange(Enum.GetNames<Jobs>().Select(job => $"{job} Compensation"));
            headers.Add("Total Hours");
            headers.Add("Total Hourly Compensation");
            headers.Add("Gross Pay");
            headers.AddRange(TrackedPayRateJobs.Select(job => $"{job} Rate"));

            Dictionary<int, float> grossPayByEmployee = employees.ToDictionary(employee => employee.IdNumber, GetGrossPay);
            List<List<string>> dataRows = new();
            foreach (Employee employee in employees.OrderBy(employee => employee.IdNumber))
            {
                bool hasAnyShifts = employee.Shifts.Any(shift => !shift.IsATotalsShift);
                bool includeSalaryOnlyRow = !hasAnyShifts && ShouldIncludeSalary(employee);
                foreach (Company company in Enum.GetValues<Company>())
                {
                    List<Shift> shifts = employee.Shifts
                        .Where(shift => shift.CompanyName == company && !shift.IsATotalsShift)
                        .ToList();
                    if (shifts.Count == 0 && !(includeSalaryOnlyRow && company == employee.PrimaryCompany))
                    {
                        continue;
                    }

                    Dictionary<Jobs, float> hoursByJob = new();
                    Dictionary<Jobs, float> compensationByJob = new();
                    float estimatedCoachHours = 0f;
                    foreach (Shift shift in shifts)
                    {
                        float hours = shift.AllHours(false);
                        if (shift.JobType == Jobs.DRIVER_COACH)
                        {
                            float coachEstimate = shift.CoachTripDays * 8f;
                            estimatedCoachHours += coachEstimate;
                            hours = Math.Max(hours, coachEstimate);
                        }
                        hoursByJob[shift.JobType] = hoursByJob.GetValueOrDefault(shift.JobType) + hours;
                        compensationByJob[shift.JobType] = compensationByJob.GetValueOrDefault(shift.JobType)
                            + shift.TotalCompensation(employee);
                    }

                    List<string> values = new()
                    {
                        payDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                        employee.IdNumber.ToString(CultureInfo.InvariantCulture),
                        company.ToString(),
                        FormatNumber(estimatedCoachHours)
                    };
                    values.AddRange(Enum.GetValues<Jobs>().Select(job => FormatNumber(hoursByJob.GetValueOrDefault(job))));
                    values.AddRange(Enum.GetValues<Jobs>().Select(job => FormatNumber(compensationByJob.GetValueOrDefault(job))));
                    values.Add(FormatNumber(hoursByJob.Values.Sum()));
                    values.Add(FormatNumber(compensationByJob.Values.Sum()));
                    values.Add(FormatNumber(grossPayByEmployee.GetValueOrDefault(employee.IdNumber)));
                    values.AddRange(TrackedPayRateJobs.Select(job => FormatNumber(GetCurrentPayRate(employee, job))));
                    dataRows.Add(values);
                }
            }

            Directory.CreateDirectory(HistoryDirectory);
            WritePayRateChanges(payDate, employees, dataRows.Count > 0);
            int recordCount = dataRows.Count;
            List<string> rows = new() { ToCsvRow(headers) };
            rows.AddRange(dataRows.Select(row =>
            {
                row.Insert(4, recordCount.ToString(CultureInfo.InvariantCulture));
                return ToCsvRow(row);
            }));
            File.WriteAllLines(path, rows);
        }

        private void WritePayRateChanges(DateTime payDate, IEnumerable<Employee> employees, bool hasPayrollRows)
        {
            if (!hasPayrollRows)
            {
                return;
            }

            string path = Path.Combine(HistoryDirectory, "PayRateChangeHistory.csv");
            List<string> headers = new() { "Date of Change", "Employee Number", "Employee Name", "Job", "Old Rate", "New Rate" };
            HashSet<string> existingChangeKeys = new();
            if (File.Exists(path))
            {
                foreach (string line in File.ReadLines(path).Skip(1))
                {
                    string[] values = ParseCsvRow(line);
                    if (values.Length >= 6)
                    {
                        existingChangeKeys.Add(string.Join("|", values[0], values[1], values[3], values[4], values[5]));
                    }
                }
            }

            List<string> newRows = new();
            foreach (Employee employee in employees.Where(employee => employee.Shifts.Any(shift => !shift.IsATotalsShift)))
            {
                foreach (Jobs job in TrackedPayRateJobs)
                {
                    float newRate = GetCurrentPayRate(employee, job);
                    if (!latestPayRates.TryGetValue((employee.IdNumber, job), out float oldRate)
                        || Math.Abs(oldRate - newRate) < 0.001f)
                    {
                        continue;
                    }

                    string[] values =
                    {
                        payDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                        employee.IdNumber.ToString(CultureInfo.InvariantCulture),
                        employee.Name,
                        job.ToString(),
                        FormatNumber(oldRate),
                        FormatNumber(newRate)
                    };
                    string key = string.Join("|", values[0], values[1], values[3], values[4], values[5]);
                    if (existingChangeKeys.Add(key))
                    {
                        newRows.Add(ToCsvRow(values));
                    }
                }
            }

            if (!File.Exists(path))
            {
                File.WriteAllLines(path, new[] { ToCsvRow(headers) }.Concat(newRows));
            }
            else if (newRows.Count > 0)
            {
                File.AppendAllLines(path, newRows);
            }
        }

        private static float GetCurrentPayRate(Employee employee, Jobs job) =>
            job == Jobs.SALARY ? employee.AnnualSalaryAmount : employee.PayRates.GetValueOrDefault(job);

        private static float GetCurrentHours(Employee employee) => employee.Shifts
            .Where(shift => !shift.IsATotalsShift)
            .Sum(shift => shift.JobType == Jobs.DRIVER_COACH
                ? Math.Max(shift.AllHours(false), shift.CoachTripDays * 8f)
                : shift.AllHours(false));

        private static bool HasValidCurrentShifts(Employee employee, Company company) =>
            employee.Shifts.Any(shift => shift.CompanyName == company && !shift.IsATotalsShift && shift.IsValid(employee));

        private static bool HasCurrentPay(Employee employee, Company company) =>
            HasValidCurrentShifts(employee, company)
            || employee.ManualEntries.Any(entry => entry.Company == company && entry.HasAnyValues());

        private static bool HasCurrentPayInAnyCompany(Employee employee) =>
            Enum.GetValues<Company>().Any(company => HasCurrentPay(employee, company));

        /// <summary>
        /// Companies that need a rehire row so this period's pay can be imported.
        /// A row is written only for a position that is still terminated, was terminated before
        /// this pay period, and either has pay this period or is the primary position of a
        /// company that is being rehired. Active positions are left alone. A termination during
        /// this pay period is a final check, so that position is not rehired.
        /// </summary>
        internal static List<Company> SelectCompaniesToRehire(Employee employee, DateTime payPeriodStart)
        {
            List<Company> companiesToRehire = Enum.GetValues<Company>()
                .Where(company => HasCurrentPay(employee, company)
                    && CompanyNeedsRehire(employee, company, payPeriodStart))
                .ToList();
            if (companiesToRehire.Any(company => company != employee.PrimaryCompany)
                && CompanyNeedsRehire(employee, employee.PrimaryCompany, payPeriodStart)
                && !companiesToRehire.Contains(employee.PrimaryCompany))
            {
                companiesToRehire.Add(employee.PrimaryCompany);
            }
            companiesToRehire.Sort((left, right) =>
                (left == employee.PrimaryCompany ? 0 : 1)
                .CompareTo(right == employee.PrimaryCompany ? 0 : 1));
            return companiesToRehire;
        }

        private static bool CompanyNeedsRehire(Employee employee, Company company, DateTime payPeriodStart)
        {
            if (employee.ActiveCompanies.Contains(company))
            {
                return false;
            }
            if (!employee.TerminationDateByCompany.TryGetValue(company, out DateTime terminationDate))
            {
                return false;
            }
            return terminationDate.Date < payPeriodStart.Date;
        }

        private bool ShouldConfirmSuspiciousRehire(Employee employee)
        {
            return HasNotBeenPaidInOverAYear(employee)
                && HasOnlyOnePayPeriodShift(employee)
                && !ExcelWorker.EmployeeExportByNumber.ContainsKey(employee.IdNumber);
        }

        private bool HasNotBeenPaidInOverAYear(Employee employee)
        {
            DateTime cutoff = currentPayDate.Date.AddYears(-1);
            if (lastHoursDateByEmployee.TryGetValue(employee.IdNumber, out DateTime lastHoursDate))
            {
                return lastHoursDate.Date <= cutoff;
            }
            return true;
        }

        private static bool HasOnlyOnePayPeriodShift(Employee employee) =>
            GetCurrentPayPeriodShifts(employee).Count == 1;

        private static List<Shift> GetCurrentPayPeriodShifts(Employee employee) =>
            employee.Shifts.Where(shift => !shift.IsATotalsShift && shift.IsValid(employee)).ToList();

        private bool ConfirmSuspiciousRehire(Employee employee)
        {
            Shift shift = GetCurrentPayPeriodShifts(employee)[0];
            string lastPaidText = lastHoursDateByEmployee.TryGetValue(employee.IdNumber, out DateTime lastHoursDate)
                ? lastHoursDate.ToShortDateString()
                : "no prior hours in payroll history";
            string shiftDateText = shift.Date == DateTime.MinValue ? "this pay period" : shift.Date.ToShortDateString();
            string message = "Should we really rehire this employee, or was this a mistake?"
                + Environment.NewLine + Environment.NewLine
                + employee.Name + " (" + employee.IdNumber + ") has not been paid since "
                + lastPaidText + " and only has one shift this pay period. They are not on WorkbrightEmployees."
                + Environment.NewLine + Environment.NewLine
                + "Shift: " + shift.JobType + " on " + shiftDateText
                + " (" + shift.AllHours(false).ToString("0.##") + " hours)"
                + Environment.NewLine + Environment.NewLine
                + "Yes = rehire and pay this shift."
                + Environment.NewLine
                + "No = discard the shift and do not rehire.";
            bool shouldRehire = PrintForm.InputBool(message, "Yes", "No");
            if (shouldRehire)
            {
                Program.Log("Rehire confirmed for " + employee.Name + " (" + employee.IdNumber
                    + ") after more than a year without pay.");
            }
            return shouldRehire;
        }

        internal static string HistoryFolder => HistoryDirectory;

        internal static IEnumerable<(DateTime PayDate, string Path)> EnumerateHistoryFiles() => GetHistoryFiles();

        private static IEnumerable<(DateTime PayDate, string Path)> GetHistoryFiles()
        {
            if (!Directory.Exists(HistoryDirectory))
            {
                return Enumerable.Empty<(DateTime, string)>();
            }

            return Directory.EnumerateFiles(HistoryDirectory, "PayrollHistory_*.csv")
                .Select(path => (Path: path, Name: Path.GetFileNameWithoutExtension(path)))
                .Select(file => (DateTime.TryParseExact(file.Name["PayrollHistory_".Length..], "yyyy-MM-dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime payDate), payDate, file.Path))
                .Where(file => file.Item1)
                .Select(file => (file.payDate, file.Path));
        }

        internal static bool TryReadEntries(string path, DateTime expectedPayDate, out List<Entry> entries,
            out bool hasGrossPayColumn)
        {
            entries = new();
            hasGrossPayColumn = false;
            if (!TryReadAllLines(path, out string[] lines, out _))
            {
                return false;
            }
            if (lines.Length < 2)
            {
                return false;
            }

            string[] headers = ParseCsvRow(lines[0]);
            int payDateColumn = Array.IndexOf(headers, "Payroll Date");
            int employeeColumn = Array.IndexOf(headers, "Employee Number");
            int companyColumn = Array.IndexOf(headers, "Company");
            int hoursColumn = Array.IndexOf(headers, "Total Hours");
            int recordCountColumn = Array.IndexOf(headers, "Record Count");
            int totalCompensationColumn = Array.IndexOf(headers, "Total Hourly Compensation");
            int grossPayColumn = Array.IndexOf(headers, "Gross Pay");
            hasGrossPayColumn = grossPayColumn >= 0;
            if (payDateColumn < 0 || employeeColumn < 0 || companyColumn < 0 || hoursColumn < 0 || recordCountColumn < 0)
            {
                return false;
            }
            Dictionary<Jobs, int> payRateColumns = TrackedPayRateJobs
                .Select(job => (Job: job, Column: Array.IndexOf(headers, $"{job} Rate")))
                .Where(rateColumn => rateColumn.Column >= 0)
                .ToDictionary(rateColumn => rateColumn.Job, rateColumn => rateColumn.Column);
            Dictionary<Jobs, int> jobHoursColumns = Enum.GetValues<Jobs>()
                .Select(job => (Job: job, Column: Array.IndexOf(headers, $"{job} Hours")))
                .Where(hoursColumn => hoursColumn.Column >= 0)
                .ToDictionary(hoursColumn => hoursColumn.Job, hoursColumn => hoursColumn.Column);

            int expectedRecordCount = -1;
            foreach (string line in lines.Skip(1))
            {
                string[] values = ParseCsvRow(line);
                if (string.IsNullOrWhiteSpace(line)
                    || values.Length <= new[] { payDateColumn, employeeColumn, companyColumn, hoursColumn, recordCountColumn }.Max()
                    || !TryParsePayrollDate(values[payDateColumn], out DateTime rowPayDate)
                    || rowPayDate.Date != expectedPayDate.Date
                    || !int.TryParse(values[employeeColumn], out int employeeNumber)
                    || !Enum.TryParse(values[companyColumn], out Company company)
                    || !float.TryParse(values[hoursColumn], NumberStyles.Float, CultureInfo.InvariantCulture, out float totalHours)
                    || !int.TryParse(values[recordCountColumn], out int recordCount))
                {
                    entries.Clear();
                    return false;
                }
                if (expectedRecordCount == -1)
                {
                    expectedRecordCount = recordCount;
                }
                if (expectedRecordCount != recordCount)
                {
                    entries.Clear();
                    return false;
                }
                float totalCompensation = 0f;
                if (totalCompensationColumn >= 0)
                {
                    if (values.Length <= totalCompensationColumn
                        || !float.TryParse(values[totalCompensationColumn], NumberStyles.Float, CultureInfo.InvariantCulture,
                            out totalCompensation))
                    {
                        entries.Clear();
                        return false;
                    }
                }
                float grossPay = 0f;
                if (grossPayColumn >= 0)
                {
                    if (values.Length <= grossPayColumn
                        || !float.TryParse(values[grossPayColumn], NumberStyles.Float, CultureInfo.InvariantCulture, out grossPay))
                    {
                        entries.Clear();
                        return false;
                    }
                }
                Dictionary<Jobs, float> payRates = new();
                foreach ((Jobs job, int column) in payRateColumns)
                {
                    if (values.Length <= column
                        || !float.TryParse(values[column], NumberStyles.Float, CultureInfo.InvariantCulture, out float rate))
                    {
                        entries.Clear();
                        return false;
                    }
                    payRates[job] = rate;
                }
                Dictionary<Jobs, float> hoursByJob = new();
                foreach ((Jobs job, int column) in jobHoursColumns)
                {
                    if (values.Length <= column
                        || !float.TryParse(values[column], NumberStyles.Float, CultureInfo.InvariantCulture, out float jobHours))
                    {
                        entries.Clear();
                        return false;
                    }
                    if (jobHours > 0.01f)
                    {
                        hoursByJob[job] = jobHours;
                    }
                }
                entries.Add(new Entry(employeeNumber, company, totalHours, totalCompensation, payRates, hoursByJob, grossPay));
            }
            return expectedRecordCount == entries.Count;
        }

        private static bool TryParsePayrollDate(string value, out DateTime payDate)
        {
            return DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out payDate)
                || DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out payDate);
        }

        private static string HistoryDirectory => Path.Combine(GetProjectDirectory(), "Payroll History");

        private static string GetProjectDirectory()
        {
            DirectoryInfo? directory = new(Directory.GetCurrentDirectory());
            while (directory != null)
            {
                if (directory.EnumerateFiles("*.sln").Any())
                {
                    return directory.FullName;
                }
                directory = directory.Parent;
            }
            return AppContext.BaseDirectory;
        }

        private static string FormatNumber(float value) => value.ToString("0.##", CultureInfo.InvariantCulture);
        private static string ToCsvRow(IEnumerable<string> values) => string.Join(",", values.Select(value =>
            $"\"{value.Replace("\"", "\"\"")}\""));

        internal static bool TryReadAllLines(string path, out string[] lines, out string? error)
        {
            lines = Array.Empty<string>();
            error = null;
            try
            {
                using FileStream stream = new(path, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
                using StreamReader reader = new(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                List<string> result = new();
                while (reader.ReadLine() is string line)
                {
                    result.Add(line);
                }
                lines = result.ToArray();
                return true;
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                error = exception.Message;
                return false;
            }
        }

        internal static string[] ParseCsvRow(string line)
        {
            List<string> values = new();
            StringBuilder value = new();
            bool quoted = false;
            for (int index = 0; index < line.Length; index++)
            {
                if (line[index] == '"' && quoted && index + 1 < line.Length && line[index + 1] == '"')
                {
                    value.Append('"');
                    index++;
                }
                else if (line[index] == '"')
                {
                    quoted = !quoted;
                }
                else if (line[index] == ',' && !quoted)
                {
                    values.Add(value.ToString());
                    value.Clear();
                }
                else
                {
                    value.Append(line[index]);
                }
            }
            values.Add(value.ToString());
            return values.ToArray();
        }
    }
}
