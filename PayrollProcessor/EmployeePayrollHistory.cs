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
        public sealed record Entry(int EmployeeNumber, Company Company, float TotalHours, float TotalCompensation);

        private readonly Dictionary<int, float> historicalHoursByEmployee = new();
        private readonly Dictionary<int, float> fivePreviousPayPeriodHoursByEmployee = new();
        private readonly Dictionary<int, DateTime> lastHoursDateByEmployee = new();
        private DateTime? earliestHistoryPayDate;

        public bool HasSixPreviousPayPeriods { get; private set; }
        private bool HasFiveValidPreviousPayPeriods { get; set; }
        public HashSet<int> PartTimeEmployeesNeedingFullTimeStatus { get; } = new();
        public HashSet<int> EmployeesNeedingTermination { get; } = new();
        public HashSet<int> EmployeesNeedingRehire { get; } = new();

        public void LoadPreviousPayPeriods(DateTime currentPayDate)
        {
            List<(DateTime PayDate, string Path)> allFiles = GetHistoryFiles()
                .Where(file => file.PayDate.Date < currentPayDate.Date)
                .OrderByDescending(file => file.PayDate)
                .ToList();
            List<(DateTime PayDate, string Path)> files = allFiles.Take(6).ToList();
            Dictionary<string, List<Entry>> entriesByPath = new();
            foreach ((DateTime payDate, string path) in allFiles)
            {
                if (TryReadEntries(path, payDate, out List<Entry> entries))
                {
                    entriesByPath[path] = entries;
                }
            }

            HasSixPreviousPayPeriods = files.Count == 6 && files.All(file => entriesByPath.ContainsKey(file.Path));
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
                    if (entry.TotalHours > 0.01f)
                    {
                        lastHoursDateByEmployee[entry.EmployeeNumber] =
                            !lastHoursDateByEmployee.TryGetValue(entry.EmployeeNumber, out DateTime latest) || payDate > latest
                                ? payDate : latest;
                    }
                }
            }

            foreach ((DateTime _, string path) in files)
            {
                if (!entriesByPath.TryGetValue(path, out List<Entry>? entries))
                {
                    continue;
                }
                foreach (Entry entry in entries)
                {
                    if (entry.TotalHours > 0.01f)
                    {
                        historicalHoursByEmployee[entry.EmployeeNumber] =
                            historicalHoursByEmployee.GetValueOrDefault(entry.EmployeeNumber) + entry.TotalHours;
                    }
                }
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
                bool hasCurrentHours = GetCurrentHours(employee) > 0.01f;
                if (HasFiveValidPreviousPayPeriods && hasCurrentHours && employee.EmploymentCategory == "PT"
                    && fivePreviousPayPeriodHoursByEmployee.GetValueOrDefault(employee.IdNumber) + GetCurrentHours(employee) >= 360f)
                {
                    PartTimeEmployeesNeedingFullTimeStatus.Add(employee.IdNumber);
                }

                if (HasSixPreviousPayPeriods && !employee.IsTerminated && !employee.IsSalaried && !hasCurrentHours
                    && historicalHoursByEmployee.GetValueOrDefault(employee.IdNumber) < 0.01f)
                {
                    EmployeesNeedingTermination.Add(employee.IdNumber);
                }

                if (employee.IsTerminated && hasCurrentHours
                    && !ExcelWorker.EmployeeExportByNumber.ContainsKey(employee.IdNumber))
                {
                    EmployeesNeedingRehire.Add(employee.IdNumber);
                }
            }
        }

        public DateTime GetTerminationDate(int employeeNumber, DateTime fallbackPayDate)
        {
            if (lastHoursDateByEmployee.TryGetValue(employeeNumber, out DateTime lastHoursDate))
            {
                return lastHoursDate;
            }

            return earliestHistoryPayDate ?? fallbackPayDate;
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
            headers.Add("Total Compensation");

            List<List<string>> dataRows = new();
            foreach (Employee employee in employees.OrderBy(employee => employee.IdNumber))
            {
                foreach (Company company in Enum.GetValues<Company>())
                {
                    List<Shift> shifts = employee.Shifts
                        .Where(shift => shift.CompanyName == company && !shift.IsATotalsShift)
                        .ToList();
                    if (shifts.Count == 0)
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
                    dataRows.Add(values);
                }
            }

            int recordCount = dataRows.Count;
            List<string> rows = new() { ToCsvRow(headers) };
            rows.AddRange(dataRows.Select(row =>
            {
                row.Insert(4, recordCount.ToString(CultureInfo.InvariantCulture));
                return ToCsvRow(row);
            }));
            Directory.CreateDirectory(HistoryDirectory);
            File.WriteAllLines(path, rows);
        }

        private static float GetCurrentHours(Employee employee) => employee.Shifts
            .Where(shift => !shift.IsATotalsShift)
            .Sum(shift => shift.JobType == Jobs.DRIVER_COACH
                ? Math.Max(shift.AllHours(false), shift.CoachTripDays * 8f)
                : shift.AllHours(false));

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

        private static bool TryReadEntries(string path, DateTime expectedPayDate, out List<Entry> entries)
        {
            entries = new();
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
            if (payDateColumn < 0 || employeeColumn < 0 || companyColumn < 0 || hoursColumn < 0 || recordCountColumn < 0)
            {
                return false;
            }

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
                entries.Add(new Entry(employeeNumber, company, totalHours, 0f));
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
