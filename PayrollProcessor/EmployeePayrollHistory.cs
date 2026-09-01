using System.Globalization;
using System.Text;

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

        private readonly Dictionary<int, float> fivePreviousPayPeriodHoursByEmployee = new();
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

        private bool HasFiveValidPreviousPayPeriods { get; set; }
        public HashSet<int> PartTimeEmployeesNeedingFullTimeStatus { get; } = new();
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

            HasFiveValidPreviousPayPeriods = allFiles.Take(5).Count() == 5
                && allFiles.Take(5).All(file => entriesByPath.ContainsKey(file.Path));
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

            foreach ((DateTime _, string path) in allFiles.Take(5))
            {
                if (!entriesByPath.TryGetValue(path, out List<Entry>? entries))
                {
                    continue;
                }
                foreach (Entry entry in entries.Where(entry => entry.TotalHours > 0.01f))
                {
                    fivePreviousPayPeriodHoursByEmployee[entry.EmployeeNumber] =
                        fivePreviousPayPeriodHoursByEmployee.GetValueOrDefault(entry.EmployeeNumber) + entry.TotalHours;
                }
            }
        }

        public void EvaluateEmployees(IEnumerable<Employee> employees)
        {
            foreach (Employee employee in employees)
            {
                if (employee.IdNumber == 503)
                {
                    //John McLaughlin exception
                    continue;
                }
                bool hasCurrentHours = GetCurrentHours(employee) > 0.01f;
                if (HasFiveValidPreviousPayPeriods && hasCurrentHours && employee.EmploymentCategory == "PT"
                    && fivePreviousPayPeriodHoursByEmployee.GetValueOrDefault(employee.IdNumber) + GetCurrentHours(employee) >= 360f)
                {
                    PartTimeEmployeesNeedingFullTimeStatus.Add(employee.IdNumber);
                }

                HashSet<int> terminationExceptions = new() {105, 187, 501, 503};
                int lookbackPayPeriods = GetTerminationLookbackPayPeriods(employee);
                if (previousPayPeriodHoursNewestFirst.Count >= lookbackPayPeriods
                    && !employee.IsTerminated 
                    && !employee.IsSalaried 
                    && !hasCurrentHours
                    && employee.WasAlreadyInPayroll
                    && GetHistoricalHours(employee.IdNumber, lookbackPayPeriods) < 0.01f 
                    && !terminationExceptions.Contains(employee.IdNumber))
                {
                    EmployeesNeedingTermination.Add(employee.IdNumber);
                }

                if (employee.IsTerminated && !ExcelWorker.EmployeeExportByNumber.ContainsKey(employee.IdNumber))
                {
                    List<Company> companiesToRehire = Enum.GetValues<Company>()
                        .Where(company => HasValidCurrentShifts(employee, company))
                        .ToList();
                    if (companiesToRehire.Any(company => company != employee.PrimaryCompany)
                        && !companiesToRehire.Contains(employee.PrimaryCompany))
                    {
                        companiesToRehire.Add(employee.PrimaryCompany);
                    }
                    if (companiesToRehire.Count > 0)
                    {
                        companiesToRehire.Sort((left, right) =>
                            (left == employee.PrimaryCompany ? 0 : 1)
                            .CompareTo(right == employee.PrimaryCompany ? 0 : 1));
                        EmployeesNeedingRehire[employee.IdNumber] = companiesToRehire;
                    }
                    else if (employee.ActiveCompanies.Any(company => company != employee.PrimaryCompany)
                        && !terminationExceptions.Contains(employee.IdNumber))
                    {
                        EmployeesNeedingTerminationInNonPrimaryCompanyOnly.Add(employee.IdNumber);
                    }
                }
            }
        }

        private int GetTerminationLookbackPayPeriods(Employee employee)
        {
            if (IsFullTime(employee))
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

        private static bool IsFullTime(Employee employee)
        {
            string category = employee.EmploymentCategory?.Trim() ?? "";
            return category.Equals("FT", StringComparison.OrdinalIgnoreCase)
                || category.Equals("ACAFT", StringComparison.OrdinalIgnoreCase);
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
            string[] lines;
            try
            {
                lines = File.ReadAllLines(path);
            }
            catch (IOException)
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

        private static string[] ParseCsvRow(string line)
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
