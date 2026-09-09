using System.Globalization;
using System.Text.RegularExpressions;

namespace PayrollProcessor
{
    internal sealed class PayrollHistoryEmployee
    {
        public int EmployeeNumber { get; init; }
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public string MiddleName { get; set; } = "";
        public DateTime? HireDate { get; set; }
        public DateTime? RehireDate { get; set; }
        public DateTime? TerminationDate { get; set; }
        public DateTime? LastPaidDate { get; set; }
        public string EmploymentStatus { get; set; } = "";
        public bool IsSalaried { get; set; }
        public float AnnualSalary { get; set; }
        public float HighestHourlyRate { get; set; }
        public Dictionary<(DateTime PayDate, Company Company), PayrollHistoryPeriod> Periods { get; } = new();

        public DateTime? CurrentStartDate
        {
            get
            {
                if (RehireDate.HasValue && (!HireDate.HasValue || RehireDate.Value.Date > HireDate.Value.Date))
                {
                    return RehireDate;
                }
                return HireDate;
            }
        }

        public string DisplayName
        {
            get
            {
                string name = (FirstName + " " + LastName).Trim();
                return name == "" ? EmployeeNumber.ToString() : name;
            }
        }
    }

    internal sealed class PayrollHistoryPeriod
    {
        public DateTime PayDate { get; set; }
        public Company Company { get; set; }
        public string Source { get; set; } = "";
        public float TotalHours { get; set; }
        public float RegularHours { get; set; }
        public float OvertimeHours { get; set; }
        public float HolidayHours { get; set; }
        public float VacationHours { get; set; }
        public float MinGuaranteeHours { get; set; }
        public float BackPayHours { get; set; }
        public float GrossPay { get; set; }
        public float NetPay { get; set; }
        public float RegularEarnings { get; set; }
        public float OvertimeEarnings { get; set; }
        public float BonusEarnings { get; set; }
        public float TipsEarnings { get; set; }
        public float MinGuaranteeEarnings { get; set; }
        public float HolidayEarnings { get; set; }
        public float VacationEarnings { get; set; }
        public float BackPayEarnings { get; set; }
        public float EmployeeTaxes { get; set; }
        public float BasePayEarnings => RegularEarnings + MinGuaranteeEarnings + BackPayEarnings;
        public float CompensatedHours =>
            RegularHours + OvertimeHours + HolidayHours + VacationHours + MinGuaranteeHours + BackPayHours;

        public void Add(PayrollHistoryPeriod other)
        {
            TotalHours += other.TotalHours;
            RegularHours += other.RegularHours;
            OvertimeHours += other.OvertimeHours;
            HolidayHours += other.HolidayHours;
            VacationHours += other.VacationHours;
            MinGuaranteeHours += other.MinGuaranteeHours;
            BackPayHours += other.BackPayHours;
            GrossPay += other.GrossPay;
            NetPay += other.NetPay;
            RegularEarnings += other.RegularEarnings;
            OvertimeEarnings += other.OvertimeEarnings;
            BonusEarnings += other.BonusEarnings;
            TipsEarnings += other.TipsEarnings;
            MinGuaranteeEarnings += other.MinGuaranteeEarnings;
            HolidayEarnings += other.HolidayEarnings;
            VacationEarnings += other.VacationEarnings;
            BackPayEarnings += other.BackPayEarnings;
            EmployeeTaxes += other.EmployeeTaxes;
            if (string.IsNullOrWhiteSpace(Source))
            {
                Source = other.Source;
            }
            else if (!string.IsNullOrWhiteSpace(other.Source) && Source != other.Source && !Source.Contains(other.Source))
            {
                Source += "/" + other.Source;
            }
        }
    }

    internal sealed class PayrollHistoryCatalog
    {
        public Dictionary<int, PayrollHistoryEmployee> Employees { get; } = new();
        public List<string> LoadWarnings { get; } = new();
        public int IsolvedFileCount { get; private set; }
        public int AdpFileCount { get; private set; }
        public DateTime? MostRecentPayDate { get; private set; }
        public int EmployeesWithMostRecentPayroll { get; private set; }
        private readonly HashSet<string> loadedFilePaths = new(StringComparer.OrdinalIgnoreCase);

        public string? MostRecentPayrollMissingError
        {
            get
            {
                if (EmployeesWithMostRecentPayroll > 0)
                {
                    return null;
                }
                if (MostRecentPayDate.HasValue)
                {
                    return "Most recent payroll data (" + MostRecentPayDate.Value.ToString("M/d/yyyy")
                        + ") was not found for at least one employee.";
                }
                return "Most recent payroll data was not found for at least one employee.";
            }
        }

        public void LoadEmployees()
        {
            Employees.Clear();
            loadedFilePaths.Clear();
            LoadWarnings.Clear();
            IsolvedFileCount = 0;
            AdpFileCount = 0;
            MostRecentPayDate = null;
            EmployeesWithMostRecentPayroll = 0;
            LoadWorkforceNowEmployees();
        }

        public void EnsurePayrollLoaded(PayrollHistoryEmployee employee, PayrollLoadNeed need)
        {
            string folder = EmployeePayrollHistory.HistoryFolder;
            if (!Directory.Exists(folder))
            {
                throw new DirectoryNotFoundException("Payroll History folder was not found: " + folder);
            }

            DateTime rangeStart = need.RangeStart.Date;
            DateTime rangeEnd = need.RangeEnd.Date;
            if (rangeEnd < rangeStart)
            {
                (rangeStart, rangeEnd) = (rangeEnd, rangeStart);
            }
            int housingStartYear = DateTime.Today.Year - 2;
            List<PayrollSourceFile> files = EnumeratePayrollFilesNewestFirst(folder);
            for (int index = 0; index < files.Count; index++)
            {
                PayrollSourceFile file = files[index];
                string fullPath = Path.GetFullPath(file.Path);
                if (loadedFilePaths.Contains(fullPath))
                {
                    continue;
                }
                if (ShouldSkipPayrollFile(file, employee, need, rangeStart, rangeEnd, housingStartYear))
                {
                    continue;
                }

                try
                {
                    if (file.IsAdp)
                    {
                        LoadAdpFile(file.Path);
                        AdpFileCount++;
                    }
                    else
                    {
                        LoadIsolvedFile(file.Path);
                        IsolvedFileCount++;
                    }
                }
                catch (Exception exception)
                {
                    LoadWarnings.Add("Could not read " + Path.GetFileName(file.Path) + ": " + exception.Message);
                }
                loadedFilePaths.Add(fullPath);

                if (!need.HousingYears && need.LastSixPayPeriods && DistinctPayDateCount(employee) >= 6)
                {
                    bool siblingPending = files.Skip(index + 1).Any(remaining =>
                        !remaining.IsAdp
                        && remaining.Year == file.Year
                        && remaining.Year > 0
                        && !loadedFilePaths.Contains(Path.GetFullPath(remaining.Path)));
                    if (!siblingPending)
                    {
                        break;
                    }
                }
            }

            RefreshLastPaidDates();
            DetermineMostRecentPayroll();
        }

        public void LoadPayrollForDateRange(DateTime startDate, DateTime endDate)
        {
            EnsurePayrollLoaded(new PayrollHistoryEmployee { EmployeeNumber = -1 }, new PayrollLoadNeed
            {
                LastSixPayPeriods = false,
                HousingYears = false,
                RangeStart = startDate,
                RangeEnd = endDate
            });
        }

        public IEnumerable<PayrollHistoryEmployee> Search(string query)
        {
            query = PayrollHistoryValueParser.Normalize(query);
            IEnumerable<PayrollHistoryEmployee> employees = Employees.Values
                .OrderBy(employee => employee.LastName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(employee => employee.FirstName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(employee => employee.EmployeeNumber);

            if (query == "")
            {
                return employees;
            }

            bool hasNumber = PayrollHistoryValueParser.TryGetInt(query, out int employeeNumber);
            return employees.Where(employee =>
                (hasNumber && (employee.EmployeeNumber == employeeNumber
                    || employee.EmployeeNumber.ToString(CultureInfo.InvariantCulture).Contains(query)))
                || employee.FirstName.Contains(query, StringComparison.OrdinalIgnoreCase)
                || employee.LastName.Contains(query, StringComparison.OrdinalIgnoreCase)
                || employee.DisplayName.Contains(query, StringComparison.OrdinalIgnoreCase)
                || employee.EmployeeNumber.ToString(CultureInfo.InvariantCulture).Contains(query));
        }

        public List<PayrollHistoryPeriod> GetPeriods(PayrollHistoryEmployee employee, bool lastSixPayPeriods,
            DateTime startDate, DateTime endDate)
        {
            List<DateTime> payDates = employee.Periods.Keys
                .Select(key => key.PayDate.Date)
                .Distinct()
                .OrderByDescending(date => date)
                .ToList();
            if (lastSixPayPeriods)
            {
                HashSet<DateTime> lastSix = payDates.Take(6).ToHashSet();
                return employee.Periods.Values
                    .Where(period => lastSix.Contains(period.PayDate.Date))
                    .OrderBy(period => period.PayDate)
                    .ThenBy(period => period.Company)
                    .ToList();
            }

            DateTime start = startDate.Date;
            DateTime end = endDate.Date;
            if (end < start)
            {
                (start, end) = (end, start);
            }
            return employee.Periods.Values
                .Where(period => period.PayDate.Date >= start && period.PayDate.Date <= end)
                .OrderBy(period => period.PayDate)
                .ThenBy(period => period.Company)
                .ToList();
        }

        private void RefreshLastPaidDates()
        {
            foreach (PayrollHistoryEmployee employee in Employees.Values)
            {
                if (employee.Periods.Count > 0)
                {
                    employee.LastPaidDate = employee.Periods.Keys.Max(key => key.PayDate);
                }
            }
        }

        private static int DistinctPayDateCount(PayrollHistoryEmployee employee) =>
            employee.Periods.Keys.Select(key => key.PayDate.Date).Distinct().Count();

        private static bool ShouldSkipPayrollFile(PayrollSourceFile file, PayrollHistoryEmployee employee,
            PayrollLoadNeed need, DateTime rangeStart, DateTime rangeEnd, int housingStartYear)
        {
            if (need.HousingYears)
            {
                return !file.IsAdp && file.Year > 0 && file.Year < housingStartYear;
            }

            if (need.LastSixPayPeriods)
            {
                if (file.IsAdp || DistinctPayDateCount(employee) < 6)
                {
                    return false;
                }

                int minYear = employee.Periods.Keys
                    .Select(key => key.PayDate.Date)
                    .Distinct()
                    .OrderByDescending(date => date)
                    .Take(6)
                    .Min()
                    .Year;
                return file.Year > 0 && file.Year < minYear;
            }

            if (file.IsAdp)
            {
                return rangeEnd < new DateTime(2026, 1, 1);
            }
            return file.Year > 0 && (file.Year < rangeStart.Year || file.Year > rangeEnd.Year);
        }

        private static List<PayrollSourceFile> EnumeratePayrollFilesNewestFirst(string folder)
        {
            List<PayrollSourceFile> files = new();
            foreach (string path in Directory.EnumerateFiles(folder, "AdpPayrollHistory*.xlsx")
                .Where(path => !Path.GetFileName(path).StartsWith("~$")))
            {
                files.Add(new PayrollSourceFile
                {
                    Path = path,
                    IsAdp = true,
                    Year = 0,
                    AdpNumber = ExtractTrailingNumber(Path.GetFileNameWithoutExtension(path))
                });
            }
            foreach (string path in Directory.EnumerateFiles(folder, "iSolvedPayrollRegister*.xlsx")
                .Where(path => !Path.GetFileName(path).StartsWith("~$")))
            {
                files.Add(new PayrollSourceFile
                {
                    Path = path,
                    IsAdp = false,
                    Year = GetIsolvedYear(Path.GetFileName(path)),
                    AdpNumber = 0
                });
            }

            return files
                .OrderBy(file => file.IsAdp ? 0 : 1)
                .ThenByDescending(file => file.AdpNumber)
                .ThenByDescending(file => file.Year)
                .ThenBy(file => file.Path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static int GetIsolvedYear(string fileName)
        {
            Match match = Regex.Match(fileName, @"_(\d{4})\.xlsx$", RegexOptions.IgnoreCase);
            return match.Success && int.TryParse(match.Groups[1].Value, out int year) ? year : 0;
        }

        private static string WorkforceNowPath()
        {
            string projectPath = Path.GetFullPath(Path.Combine(EmployeePayrollHistory.HistoryFolder, "..",
                "WfnEmployees.xlsx"));
            if (File.Exists(projectPath))
            {
                return projectPath;
            }

            const string specified =
                @"C:\Users\User\valleybusllc.com\Admin Team - Payroll - Payroll\Payroll\PayrollProcessor\WfnEmployees.xlsx";
            return File.Exists(specified) ? specified : projectPath;
        }

        private void DetermineMostRecentPayroll()
        {
            DateTime? latestFromFiles = null;
            List<DateTime> csvDates = EmployeePayrollHistory.EnumerateHistoryFiles()
                .Select(file => file.PayDate.Date)
                .ToList();
            if (csvDates.Count > 0)
            {
                latestFromFiles = csvDates.Max();
            }

            DateTime? latestFromRegisters = null;
            List<DateTime> registerDates = Employees.Values
                .SelectMany(employee => employee.Periods.Keys)
                .Select(key => key.PayDate.Date)
                .ToList();
            if (registerDates.Count > 0)
            {
                latestFromRegisters = registerDates.Max();
            }

            DateTime? knownLatest = null;
            if (latestFromFiles.HasValue && latestFromRegisters.HasValue)
            {
                knownLatest = latestFromFiles.Value >= latestFromRegisters.Value
                    ? latestFromFiles : latestFromRegisters;
            }
            else
            {
                knownLatest = latestFromFiles ?? latestFromRegisters;
            }

            MostRecentPayDate = knownLatest;
            if (MostRecentPayDate.HasValue)
            {
                DateTime expected = MostRecentPayDate.Value.Date;
                while (expected.AddDays(14).Date <= DateTime.Today)
                {
                    expected = expected.AddDays(14);
                }
                MostRecentPayDate = expected;
            }

            if (!MostRecentPayDate.HasValue)
            {
                EmployeesWithMostRecentPayroll = 0;
                return;
            }

            DateTime mostRecent = MostRecentPayDate.Value.Date;
            EmployeesWithMostRecentPayroll = Employees.Values.Count(employee =>
                employee.Periods.Keys.Any(key => key.PayDate.Date == mostRecent));
        }

        private void LoadIsolvedFile(string path)
        {
            List<string[]> rows = PayrollHistoryExcelReader.ReadSheet(path, "Check Details");
            if (rows.Count < 3)
            {
                LoadWarnings.Add(Path.GetFileName(path) + " did not contain Check Details rows.");
                return;
            }

            string[] headers = BuildIsolvedHeaders(rows[0], rows[1]);
            Dictionary<string, int> columns = BuildColumnMap(headers);
            string fileName = Path.GetFileName(path);
            for (int rowIndex = 2; rowIndex < rows.Count; rowIndex++)
            {
                string[] row = rows[rowIndex];
                if (!TryGetInt(columns, row, out int employeeNumber, "Number", "Employee Number", "Employee")
                    || employeeNumber <= 0)
                {
                    continue;
                }
                if (!TryGetDate(columns, row, out DateTime payDate, "Pay Date"))
                {
                    continue;
                }

                string companyCode = Get(columns, row, "Legal Company Code", "Company Code", "Company");
                string companyName = Get(columns, row, "Legal Company Name", "Company Name");
                if (!TryParseCompany(companyCode, companyName, fileName, out Company company))
                {
                    continue;
                }

                PayrollHistoryPeriod period = new()
                {
                    PayDate = payDate,
                    Company = company,
                    Source = "iSolved",
                    GrossPay = GetFloat(columns, row, "Gross Pay"),
                    NetPay = GetFloat(columns, row, "Net Pay"),
                    RegularHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "Hourly Regular", "Regular Hours")),
                    OvertimeHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "Overtime", "Overtim")),
                    HolidayHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "Holiday")),
                    VacationHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "Vacation")),
                    MinGuaranteeHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "Min Guaran", "Min Guarantee")),
                    BackPayHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "Back Pay")),
                    RegularEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Hourly Regular", "Regular Dollars", "Salary", "Driver Coach")),
                    OvertimeEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Overtime", "Overtim")),
                    BonusEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Bonus") && !ContainsAny(value, "Tips")),
                    TipsEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Tips")),
                    MinGuaranteeEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Min Guaran", "Min Guarantee")),
                    HolidayEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Holiday")),
                    VacationEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Vacation")),
                    BackPayEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "Back Pay")),
                    EmployeeTaxes = SumMatching(headers, row, value => ContainsAny(value, "SOC SEC EE", "MED EE", "FEDERAL WH", "NORTH DAKOTA WH", "MINNESOTA WH"))
                };
                period.TotalHours = SumMatching(headers, row, IsHours);
                if (period.TotalHours <= 0.001f)
                {
                    period.TotalHours = period.RegularHours + period.OvertimeHours + period.HolidayHours
                        + period.VacationHours + period.MinGuaranteeHours + period.BackPayHours;
                }

                RememberName(columns, row, employeeNumber, "Name");
                AddOrReplacePeriod(employeeNumber, period, overwriteExisting: false);
            }
        }

        private void LoadAdpFile(string path)
        {
            List<string[]> rows = PayrollHistoryExcelReader.ReadSheet(path, "Payroll History");
            if (rows.Count < 2)
            {
                LoadWarnings.Add(Path.GetFileName(path) + " did not contain Payroll History rows.");
                return;
            }

            string[] headers = rows[0].Select(PayrollHistoryValueParser.Normalize).ToArray();
            Dictionary<string, int> columns = BuildColumnMap(headers);
            Dictionary<(int EmployeeNumber, Company Company, DateTime PayDate), PayrollHistoryPeriod> aggregated = new();
            for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                string[] row = rows[rowIndex];
                string companyCode = Get(columns, row, "COMPANY CODE", "Company Code", "Payroll Company Code", "Company");
                if (companyCode.StartsWith("Totals", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (!TryGetInt(columns, row, out int employeeNumber, "FILE NUMBER", "File Number", "Employee Number")
                    || employeeNumber <= 0)
                {
                    continue;
                }
                if (!TryGetDate(columns, row, out DateTime payDate, "PAY DATE", "Pay Date"))
                {
                    continue;
                }
                if (!TryParseCompany(companyCode, "", Path.GetFileName(path), out Company company))
                {
                    continue;
                }

                float takeHome = GetFloat(columns, row, "TAKE HOME", "Take Home");
                float netPay = GetFloat(columns, row, "NET PAY", "Net Pay");
                PayrollHistoryPeriod period = new()
                {
                    PayDate = payDate,
                    Company = company,
                    Source = "ADP",
                    GrossPay = GetFloat(columns, row, "GROSS PAY", "Gross Pay"),
                    NetPay = Math.Abs(takeHome) > 0.001f ? takeHome : netPay,
                    RegularHours = GetFloat(columns, row, "REGULAR HOURS", "Regular Hours"),
                    OvertimeHours = GetFloat(columns, row, "OVERTIME HOURS", "Overtime Hours"),
                    HolidayHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "HOL-HOLIDAY", "HOLIDAY")),
                    VacationHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "VAC-VACATION", "VACATION")),
                    MinGuaranteeHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "MNG-MIN GUARANT", "MIN GUARANT")),
                    BackPayHours = SumMatching(headers, row, value => IsHours(value) && ContainsAny(value, "BCK-BACK PAY", "BACK PAY")),
                    RegularEarnings = GetFloat(columns, row, "REGULAR EARNINGS", "Regular Earnings"),
                    OvertimeEarnings = GetFloat(columns, row, "OVERTIME EARNINGS", "Overtime Earnings"),
                    BonusEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "BON-BONUS", "BONUS") && !ContainsAny(value, "TIPS")),
                    TipsEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "TIP-TIPS", "TIPS")),
                    MinGuaranteeEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "MNG-MIN GUARANT", "MIN GUARANT")),
                    HolidayEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "HOL-HOLIDAY", "HOLIDAY")),
                    VacationEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "VAC-VACATION", "VACATION")),
                    BackPayEarnings = SumMatching(headers, row, value => IsMoney(value) && ContainsAny(value, "BCK-BACK PAY", "BACK PAY")),
                    EmployeeTaxes = GetFloat(columns, row, "TOTAL EMPLOYEE TAX", "Total Employee Tax"),
                    TotalHours = GetFloat(columns, row, "TOTAL HOURS", "Total Hours")
                };
                if (period.TotalHours <= 0.001f)
                {
                    period.TotalHours = period.RegularHours + period.OvertimeHours + period.HolidayHours
                        + period.VacationHours + period.MinGuaranteeHours + period.BackPayHours;
                }

                RememberName(columns, row, employeeNumber, "NAME", "Name");
                (int EmployeeNumber, Company Company, DateTime PayDate) key = (employeeNumber, company, payDate);
                if (aggregated.TryGetValue(key, out PayrollHistoryPeriod? existing))
                {
                    existing.Add(period);
                }
                else
                {
                    aggregated[key] = period;
                }
            }

            foreach (((int EmployeeNumber, Company Company, DateTime PayDate) key, PayrollHistoryPeriod period) in aggregated)
            {
                AddOrReplacePeriod(key.EmployeeNumber, period, overwriteExisting: true);
            }
        }

        private void AddOrReplacePeriod(int employeeNumber, PayrollHistoryPeriod period, bool overwriteExisting)
        {
            PayrollHistoryEmployee employee = GetOrCreateEmployee(employeeNumber);
            (DateTime PayDate, Company Company) key = (period.PayDate.Date, period.Company);
            if (employee.Periods.ContainsKey(key))
            {
                if (overwriteExisting)
                {
                    employee.Periods[key] = period;
                }
                return;
            }
            employee.Periods[key] = period;
        }

        private void LoadWorkforceNowEmployees()
        {
            string path = WorkforceNowPath();
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("WfnEmployees.xlsx was not found: " + path);
            }

            try
            {
                List<string[]> rows = PayrollHistoryExcelReader.ReadSheet(path, null);
                if (rows.Count < 2)
                {
                    return;
                }

                Dictionary<string, int> columns = BuildColumnMap(rows[0].Select(PayrollHistoryValueParser.Normalize).ToArray());
                Dictionary<int, List<WfnRow>> rowsByEmployee = new();
                for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
                {
                    string[] row = rows[rowIndex];
                    if (!TryGetInt(columns, row, out int employeeNumber, "File Number", "FILE NUMBER")
                        || employeeNumber <= 0)
                    {
                        continue;
                    }

                    PayrollHistoryValueParser.ParseName(
                        Get(columns, row, "Legal Last Name") == ""
                            ? ""
                            : Get(columns, row, "Legal Last Name") + ", " + Get(columns, row, "Legal First Name") + " "
                                + Get(columns, row, "Legal Middle Name"),
                        out string firstName, out string middleName, out string lastName);
                    if (firstName == "")
                    {
                        firstName = Get(columns, row, "Legal First Name");
                    }
                    if (lastName == "")
                    {
                        lastName = Get(columns, row, "Legal Last Name");
                    }
                    if (middleName == "")
                    {
                        middleName = Get(columns, row, "Legal Middle Name");
                    }

                    PayrollHistoryEmployee employee = GetOrCreateEmployee(employeeNumber);
                    PreferName(employee, firstName, middleName, lastName);

                    if (!rowsByEmployee.TryGetValue(employeeNumber, out List<WfnRow>? wfnRows))
                    {
                        wfnRows = new();
                        rowsByEmployee[employeeNumber] = wfnRows;
                    }
                    TryGetDate(columns, row, out DateTime hireDate, "Hire Date");
                    TryGetDate(columns, row, out DateTime rehireDate, "Rehire Date");
                    TryGetDate(columns, row, out DateTime termDate, "Termination Date");
                    float annualSalary = GetFloat(columns, row, "Annual Salary");
                    float highestRate = 0f;
                    foreach (string header in columns.Keys)
                    {
                        if (header.StartsWith("Rate -", StringComparison.OrdinalIgnoreCase)
                            || header.StartsWith("Rate-", StringComparison.OrdinalIgnoreCase))
                        {
                            highestRate = Math.Max(highestRate, GetFloat(columns, row, header));
                        }
                    }
                    wfnRows.Add(new WfnRow(
                        Get(columns, row, "Position Status", "STATUS"),
                        Get(columns, row, "Primary Position"),
                        hireDate == DateTime.MinValue ? null : hireDate,
                        rehireDate == DateTime.MinValue ? null : rehireDate,
                        termDate == DateTime.MinValue ? null : termDate,
                        annualSalary,
                        highestRate));
                }

                foreach ((int employeeNumber, List<WfnRow> wfnRows) in rowsByEmployee)
                {
                    PayrollHistoryEmployee employee = GetOrCreateEmployee(employeeNumber);
                    employee.HireDate = MinDate(wfnRows.Select(row => row.HireDate));
                    employee.RehireDate = MaxDate(wfnRows.Select(row => row.RehireDate));
                    bool anyActive = wfnRows.Any(row =>
                        row.Status.Contains("Active", StringComparison.OrdinalIgnoreCase));
                    employee.TerminationDate = anyActive ? null : MaxDate(wfnRows.Select(row => row.TerminationDate));
                    employee.AnnualSalary = wfnRows.Max(row => row.AnnualSalary);
                    employee.HighestHourlyRate = wfnRows.Max(row => row.HighestRate);
                    employee.IsSalaried = employee.AnnualSalary > 50f;
                    if (anyActive)
                    {
                        employee.EmploymentStatus = "Actively Employed";
                    }
                    else if (employee.TerminationDate.HasValue)
                    {
                        employee.EmploymentStatus = "Terminated";
                    }
                    else
                    {
                        employee.EmploymentStatus = "Employed";
                    }
                }
            }
            catch (Exception exception)
            {
                LoadWarnings.Add("Could not read WfnEmployees.xlsx: " + exception.Message);
            }
        }

        private PayrollHistoryEmployee GetOrCreateEmployee(int employeeNumber)
        {
            if (!Employees.TryGetValue(employeeNumber, out PayrollHistoryEmployee? employee))
            {
                employee = new PayrollHistoryEmployee { EmployeeNumber = employeeNumber };
                Employees[employeeNumber] = employee;
            }
            return employee;
        }

        private void RememberName(Dictionary<string, int> columns, string[] row, int employeeNumber, params string[] nameHeaders)
        {
            string name = Get(columns, row, nameHeaders);
            PayrollHistoryValueParser.ParseName(name, out string firstName, out string middleName, out string lastName);
            PreferName(GetOrCreateEmployee(employeeNumber), firstName, middleName, lastName);
        }

        private static void PreferName(PayrollHistoryEmployee employee, string firstName, string middleName, string lastName)
        {
            if (!string.IsNullOrWhiteSpace(firstName) && string.IsNullOrWhiteSpace(employee.FirstName))
            {
                employee.FirstName = firstName;
            }
            if (!string.IsNullOrWhiteSpace(middleName) && string.IsNullOrWhiteSpace(employee.MiddleName))
            {
                employee.MiddleName = middleName;
            }
            if (!string.IsNullOrWhiteSpace(lastName) && string.IsNullOrWhiteSpace(employee.LastName))
            {
                employee.LastName = lastName;
            }
        }

        private static string[] BuildIsolvedHeaders(string[] topRow, string[] bottomRow)
        {
            int length = Math.Max(topRow.Length, bottomRow.Length);
            string[] headers = new string[length];
            for (int i = 0; i < length; i++)
            {
                string top = i < topRow.Length ? PayrollHistoryValueParser.Normalize(topRow[i]) : "";
                string bottom = i < bottomRow.Length ? PayrollHistoryValueParser.Normalize(bottomRow[i]) : "";
                headers[i] = CombineIsolvedHeader(top, bottom);
            }
            return headers;
        }

        private static string CombineIsolvedHeader(string top, string bottom)
        {
            if (top == "")
            {
                return bottom;
            }
            if (bottom == "")
            {
                return top;
            }
            if (IsIdentityHeader(bottom))
            {
                return bottom;
            }
            if (IsHeaderQualifier(bottom))
            {
                return top;
            }
            if (IsHours(top) || IsMoney(top))
            {
                return top;
            }
            return (top + " " + bottom).Trim();
        }

        private static bool IsIdentityHeader(string header)
        {
            string[] identity =
            {
                "Legal Company Code", "Legal Company Name", "Number", "Name", "Run#", "Pay Type", "Job Title",
                "Pay Date", "Check Type Description", "Check Number", "Gross Pay", "Paid Gross", "Net Pay", "Check Amount"
            };
            return identity.Any(value => header.Equals(value, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsHeaderQualifier(string header)
        {
            return header.Equals("Amount", StringComparison.OrdinalIgnoreCase)
                || header.Equals("Memo Dollars", StringComparison.OrdinalIgnoreCase)
                || header.Equals("Tax Amount", StringComparison.OrdinalIgnoreCase)
                || header.Equals("Hours", StringComparison.OrdinalIgnoreCase)
                || header.Equals("Dollars", StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, int> BuildColumnMap(string[] headers)
        {
            Dictionary<string, int> columns = new(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Length; i++)
            {
                string header = PayrollHistoryValueParser.Normalize(headers[i]);
                if (header != "" && !columns.ContainsKey(header))
                {
                    columns[header] = i;
                }
            }
            return columns;
        }

        private static string Get(Dictionary<string, int> columns, string[] row, params string[] headers)
        {
            foreach (string header in headers)
            {
                if (columns.TryGetValue(header, out int index) && index < row.Length)
                {
                    return PayrollHistoryValueParser.Normalize(row[index]);
                }
            }
            return "";
        }

        private static bool TryGetInt(Dictionary<string, int> columns, string[] row, out int number, params string[] headers)
        {
            return PayrollHistoryValueParser.TryGetInt(Get(columns, row, headers), out number);
        }

        private static bool TryGetDate(Dictionary<string, int> columns, string[] row, out DateTime date, params string[] headers)
        {
            return PayrollHistoryValueParser.TryGetDate(Get(columns, row, headers), out date);
        }

        private static float GetFloat(Dictionary<string, int> columns, string[] row, params string[] headers)
        {
            return PayrollHistoryValueParser.GetFloat(Get(columns, row, headers));
        }

        private static float SumMatching(string[] headers, string[] row, Func<string, bool> match)
        {
            float total = 0f;
            int length = Math.Min(headers.Length, row.Length);
            for (int i = 0; i < length; i++)
            {
                string header = PayrollHistoryValueParser.Normalize(headers[i]);
                if (match(header))
                {
                    total += PayrollHistoryValueParser.GetFloat(row[i]);
                }
            }
            return total;
        }

        private static bool IsHours(string header) =>
            header.Contains("Hours", StringComparison.OrdinalIgnoreCase);

        private static bool IsMoney(string header) =>
            header.Contains("Dollars", StringComparison.OrdinalIgnoreCase)
            || header.Contains("Earnings", StringComparison.OrdinalIgnoreCase)
            || header.Contains("Amount", StringComparison.OrdinalIgnoreCase);

        private static bool ContainsAny(string header, params string[] fragments) =>
            fragments.Any(fragment => header.Contains(fragment, StringComparison.OrdinalIgnoreCase));

        private static bool TryParseCompany(string companyCode, string companyName, string fileName, out Company company)
        {
            company = Company.VALLEY_BUS_LLC;
            string combined = companyCode + " " + companyName + " " + fileName;
            if (combined.Contains("MKZ", StringComparison.OrdinalIgnoreCase)
                || combined.Contains("3011-1", StringComparison.OrdinalIgnoreCase)
                || combined.Contains("Coaches", StringComparison.OrdinalIgnoreCase))
            {
                company = Company.VALLEY_BUS_COACHES;
                return true;
            }
            if (combined.Contains("MMF", StringComparison.OrdinalIgnoreCase)
                || combined.Contains("3011", StringComparison.OrdinalIgnoreCase)
                || combined.Contains("Valley Bus LLC", StringComparison.OrdinalIgnoreCase)
                || combined.Contains("_VB_", StringComparison.OrdinalIgnoreCase))
            {
                company = Company.VALLEY_BUS_LLC;
                return true;
            }
            return false;
        }

        private static int ExtractTrailingNumber(string fileName)
        {
            int index = fileName.Length - 1;
            while (index >= 0 && char.IsDigit(fileName[index]))
            {
                index--;
            }
            if (index == fileName.Length - 1)
            {
                return 1;
            }
            return int.TryParse(fileName[(index + 1)..], out int number) ? number : 1;
        }

        private static DateTime? MinDate(IEnumerable<DateTime?> dates)
        {
            DateTime? min = null;
            foreach (DateTime? date in dates)
            {
                if (date.HasValue && (!min.HasValue || date.Value < min.Value))
                {
                    min = date;
                }
            }
            return min;
        }

        private static DateTime? MaxDate(IEnumerable<DateTime?> dates)
        {
            DateTime? max = null;
            foreach (DateTime? date in dates)
            {
                if (date.HasValue && (!max.HasValue || date.Value > max.Value))
                {
                    max = date;
                }
            }
            return max;
        }

        private sealed record WfnRow(string Status, string PrimaryPosition, DateTime? HireDate, DateTime? RehireDate,
            DateTime? TerminationDate, float AnnualSalary, float HighestRate);

        private sealed class PayrollSourceFile
        {
            public string Path { get; init; } = "";
            public bool IsAdp { get; init; }
            public int Year { get; init; }
            public int AdpNumber { get; init; }
        }
    }

    internal sealed class PayrollLoadNeed
    {
        public bool LastSixPayPeriods { get; init; }
        public bool HousingYears { get; init; }
        public DateTime RangeStart { get; init; }
        public DateTime RangeEnd { get; init; }
    }
}
