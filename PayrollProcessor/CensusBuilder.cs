using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    /// <summary>
    /// Temporary census export: employees paid in 2026, with WFN demographics,
    /// Christmas 12/23/2025 bonuses, and 2025 bus-starting bonus totals.
    /// </summary>
    internal sealed class CensusBuilder
    {
        private const int CensusYear = 2026;
        private static readonly DateTime ChristmasBonusDate = new(2025, 12, 23);
        private const float BusStartingBonusCapPerPeriod = 200f;
        private static readonly HashSet<int> BusStartingBonusEmployeeNumbers = new()
        {
            2580, 1252, 1116, 602, 1808, 1127, 1224, 672
        };

        private static readonly string[] Headers =
        {
            "Employee Number",
            "Name",
            "Title",
            "Full-Time / Part-Time",
            "Employment Status",
            "Hire Date",
            "Compensation Type",
            "Current Annual Base Compensation Rate",
            "Christmas Bonus",
            "Bus Starting Bonus"
        };

        public void Run()
        {
            string logPath = DefaultDirectoryPath() + "census_build.log";
            void WriteLog(string message)
            {
                string line = DateTime.Now.ToString("HH:mm:ss") + " " + message;
                Console.WriteLine(line);
                File.AppendAllText(logPath, line + Environment.NewLine);
                Log(message);
            }

            try
            {
                File.WriteAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " Census builder started"
                    + Environment.NewLine);
                WriteLog("Loading employees from WfnEmployees.xlsx...");
                PayrollHistoryCatalog catalog = new();
                catalog.LoadEmployees();
                WriteLog("Loaded " + catalog.Employees.Count + " employees. Loading 2025-2026 payroll history...");
                catalog.LoadPayrollForDateRange(new DateTime(2025, 1, 1), new DateTime(CensusYear, 12, 31));
                WriteLog("Payroll loaded. Isolved files: " + catalog.IsolvedFileCount
                    + ", ADP files: " + catalog.AdpFileCount + ".");

                Dictionary<int, CensusWfnFields> wfnFields = LoadWorkforceNowFields();
                List<CensusRow> rows = BuildRows(catalog, wfnFields);
                List<CensusRow> active = rows
                    .Where(row => row.IsActive)
                    .OrderBy(row => row.LastName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(row => row.FirstName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(row => row.EmployeeNumber)
                    .ToList();
                List<CensusRow> inactive = rows
                    .Where(row => !row.IsActive)
                    .OrderBy(row => row.LastName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(row => row.FirstName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(row => row.EmployeeNumber)
                    .ToList();

                WriteLog("Writing spreadsheet for " + active.Count + " active and "
                    + inactive.Count + " inactive employees...");
                foreach (CensusRow row in rows.Where(row => row.BusStartingBonus.HasValue)
                    .OrderBy(row => row.EmployeeNumber))
                {
                    WriteLog("Bus starting bonus " + row.EmployeeNumber + " " + row.Name + ": "
                        + row.BusStartingBonus!.Value.ToString("0.00"));
                }
                int christmasCount = rows.Count(row => row.ChristmasBonus > 0.01f);
                WriteLog("Christmas bonus recipients in census: " + christmasCount);
                string path = WriteWorkbook(active, inactive);
                foreach (string warning in catalog.LoadWarnings)
                {
                    WriteLog("Warning: " + warning);
                }
                WriteLog("2026 census written: " + path
                    + " (" + active.Count + " active, " + inactive.Count + " inactive).");
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception exception)
            {
                WriteLog("Census builder failed: " + exception);
                throw;
            }
        }

        private static List<CensusRow> BuildRows(PayrollHistoryCatalog catalog,
            Dictionary<int, CensusWfnFields> wfnFields)
        {
            List<CensusRow> rows = new();
            foreach (PayrollHistoryEmployee employee in catalog.Employees.Values)
            {
                float compensation2026 = employee.Periods.Values
                    .Where(period => period.PayDate.Year == CensusYear)
                    .Sum(period => CompensationAmount(period));
                if (compensation2026 <= 0.01f)
                {
                    continue;
                }

                wfnFields.TryGetValue(employee.EmployeeNumber, out CensusWfnFields? fields);
                fields ??= new CensusWfnFields();

                float christmasBonus = employee.Periods.Values
                    .Where(period => period.PayDate.Date == ChristmasBonusDate)
                    .Sum(period => period.BonusEarnings);
                float? busStartingBonus = null;
                if (BusStartingBonusEmployeeNumbers.Contains(employee.EmployeeNumber))
                {
                    busStartingBonus = SumBusStartingBonus(employee);
                }

                bool isActive = employee.EmploymentStatus.Contains("Active", StringComparison.OrdinalIgnoreCase)
                    || fields.PositionStatus.Contains("Active", StringComparison.OrdinalIgnoreCase);
                bool isSalaried = employee.IsSalaried && employee.AnnualSalary > 0.01f;
                float hourlyRate = employee.HighestHourlyRate;
                if (!isSalaried && hourlyRate <= 0.01f)
                {
                    hourlyRate = DefaultNonCdlRate(fields.IsGrandForks,
                        YearsOfService(fields.YearsOfService, employee.CurrentStartDate));
                }
                rows.Add(new CensusRow
                {
                    EmployeeNumber = employee.EmployeeNumber,
                    FirstName = employee.FirstName,
                    LastName = employee.LastName,
                    Name = employee.DisplayName,
                    Title = fields.Title,
                    FullTimePartTime = fields.FullTimePartTime,
                    EmploymentStatus = string.IsNullOrWhiteSpace(employee.EmploymentStatus)
                        ? (isActive ? "Actively Employed" : fields.PositionStatus)
                        : employee.EmploymentStatus,
                    HireDate = employee.CurrentStartDate,
                    IsSalaried = isSalaried,
                    AnnualSalary = employee.AnnualSalary,
                    HighestHourlyRate = hourlyRate,
                    ChristmasBonus = christmasBonus,
                    BusStartingBonus = busStartingBonus,
                    IsActive = isActive
                });
            }

            return rows;
        }

        private static float CompensationAmount(PayrollHistoryPeriod period)
        {
            if (period.GrossPay > 0.01f)
            {
                return period.GrossPay;
            }

            return period.RegularEarnings + period.OvertimeEarnings + period.BonusEarnings
                + period.TipsEarnings + period.HolidayEarnings + period.VacationEarnings
                + period.MinGuaranteeEarnings + period.BackPayEarnings;
        }

        private static float SumBusStartingBonus(PayrollHistoryEmployee employee)
        {
            Dictionary<DateTime, float> bonusByPayDate = new();
            foreach (PayrollHistoryPeriod period in employee.Periods.Values)
            {
                if (period.PayDate.Year != 2025 || period.PayDate.Date == ChristmasBonusDate)
                {
                    continue;
                }

                bonusByPayDate.TryGetValue(period.PayDate.Date, out float existing);
                bonusByPayDate[period.PayDate.Date] = existing + period.BonusEarnings;
            }

            float total = 0f;
            foreach (float amount in bonusByPayDate.Values)
            {
                if (amount <= 0.01f)
                {
                    continue;
                }
                total += Math.Min(amount, BusStartingBonusCapPerPeriod);
            }
            return total;
        }

        private static Dictionary<int, CensusWfnFields> LoadWorkforceNowFields()
        {
            Dictionary<int, CensusWfnFields> fields = new();
            string path = Path.GetFullPath(Path.Combine(EmployeePayrollHistory.HistoryFolder, "..", "WfnEmployees.xlsx"));
            if (!File.Exists(path))
            {
                path = DefaultDirectoryPath() + "WfnEmployees.xlsx";
            }
            if (!File.Exists(path))
            {
                return fields;
            }

            List<string[]> rows = PayrollHistoryExcelReader.ReadSheet(path, null);
            if (rows.Count < 2)
            {
                return fields;
            }

            Dictionary<string, int> columns = BuildColumnMap(rows[0]);
            Dictionary<int, List<WfnCensusSourceRow>> byEmployee = new();
            for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                string[] row = rows[rowIndex];
                if (!TryGetInt(columns, row, out int employeeNumber, "File Number", "FILE NUMBER")
                    || employeeNumber <= 0)
                {
                    continue;
                }

                if (!byEmployee.TryGetValue(employeeNumber, out List<WfnCensusSourceRow>? sourceRows))
                {
                    sourceRows = new();
                    byEmployee[employeeNumber] = sourceRows;
                }

                string titleDescription = Get(columns, row, "Job Title Description");
                string titleCode = Get(columns, row, "Job Title Code");
                WfnEmployeesReader.TryParseYearsOfService(Get(columns, row, "Years of Service"), out int yearsOfService);
                sourceRows.Add(new WfnCensusSourceRow(
                    Get(columns, row, "Primary Position"),
                    Get(columns, row, "Position Status", "STATUS"),
                    Get(columns, row, "Worker Category Description"),
                    PreferTitle(titleDescription, titleCode),
                    IsGrandForksLocation(Get(columns, row, "Location Code"),
                        Get(columns, row, "Primary Address: City")),
                    yearsOfService));
            }

            foreach ((int employeeNumber, List<WfnCensusSourceRow> sourceRows) in byEmployee)
            {
                WfnCensusSourceRow chosen = ChooseSourceRow(sourceRows);
                fields[employeeNumber] = new CensusWfnFields
                {
                    Title = chosen.Title,
                    FullTimePartTime = MapFullTimePartTime(chosen.WorkerCategory),
                    PositionStatus = chosen.PositionStatus,
                    IsGrandForks = chosen.IsGrandForks,
                    YearsOfService = sourceRows.Max(row => row.YearsOfService)
                };
            }

            return fields;
        }

        private static WfnCensusSourceRow ChooseSourceRow(List<WfnCensusSourceRow> sourceRows)
        {
            WfnCensusSourceRow? primary = sourceRows.FirstOrDefault(row =>
                row.PrimaryPosition.Contains("yes", StringComparison.OrdinalIgnoreCase));
            if (primary != null)
            {
                return primary;
            }

            WfnCensusSourceRow? active = sourceRows.FirstOrDefault(row =>
                row.PositionStatus.Contains("Active", StringComparison.OrdinalIgnoreCase));
            return active ?? sourceRows[0];
        }

        private static string PreferTitle(string description, string code)
        {
            if (!string.IsNullOrWhiteSpace(description) && description.Any(char.IsLetter))
            {
                return description;
            }
            return string.IsNullOrWhiteSpace(code) ? description : code;
        }

        private static bool IsGrandForksLocation(string locationCode, string city)
        {
            return locationCode.Contains("GF", StringComparison.OrdinalIgnoreCase)
                || locationCode.Contains("Grand Forks", StringComparison.OrdinalIgnoreCase)
                || (string.IsNullOrWhiteSpace(locationCode)
                    && city.Contains("Grand Forks", StringComparison.OrdinalIgnoreCase));
        }

        private static int YearsOfService(int yearsFromWfn, DateTime? startDate)
        {
            if (yearsFromWfn > 0)
            {
                return yearsFromWfn;
            }
            if (!startDate.HasValue)
            {
                return 0;
            }
            return Math.Max(0, (int)((DateTime.Today - startDate.Value.Date).TotalDays / 365.25));
        }

        private static float DefaultNonCdlRate(bool isGrandForks, int yearsOfService)
        {
            float rate = isGrandForks
                ? GrandForksDefaultRates[Jobs.NON_CDL_DRIVER]
                : FargoDefaultRates[Jobs.NON_CDL_DRIVER];
            for (int years = 6; years > 0; --years)
            {
                if (yearsOfService >= years)
                {
                    rate += 0.25f * years;
                    break;
                }
            }
            if (yearsOfService > 9)
            {
                rate += TEN_YEAR_RATE_BUMP;
            }
            return rate;
        }

        private static string MapFullTimePartTime(string categoryDescription)
        {
            string mapped = WfnEmployeesReader.MapEmploymentCategory(categoryDescription);
            if (mapped == "ACAFT")
            {
                return "Full-Time";
            }
            if (mapped == "PT")
            {
                return "Part-Time";
            }
            return string.IsNullOrWhiteSpace(categoryDescription) ? "" : categoryDescription;
        }

        private static string WriteWorkbook(List<CensusRow> active, List<CensusRow> inactive)
        {
            string path = DefaultDirectoryPath() + "2026 Employee Census.xlsx";
            Excel.Application excelApp = new()
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false
            };
            Excel.Workbook? workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Add();
                Excel.Worksheet activeSheet = (Excel.Worksheet)workbook.Worksheets[1];
                activeSheet.Name = "Active";
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[2]).Delete();
                }
                Excel.Worksheet inactiveSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: activeSheet);
                inactiveSheet.Name = "Inactive";

                WriteSheet(activeSheet, active);
                WriteSheet(inactiveSheet, inactive);
                activeSheet.Activate();

                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                excelApp.ScreenUpdating = true;
                workbook.SaveAs(path);
                workbook.Close(true);
                workbook = null;
            }
            finally
            {
                workbook?.Close(false);
                excelApp.ScreenUpdating = true;
                excelApp.Quit();
            }

            return path;
        }

        private static void WriteSheet(Excel.Worksheet sheet, List<CensusRow> rows)
        {
            object[,] values = new object[rows.Count + 1, Headers.Length];
            for (int column = 0; column < Headers.Length; column++)
            {
                values[0, column] = Headers[column];
            }

            for (int i = 0; i < rows.Count; i++)
            {
                CensusRow row = rows[i];
                int excelRow = i + 1;
                values[excelRow, 0] = row.EmployeeNumber;
                values[excelRow, 1] = row.Name;
                values[excelRow, 2] = row.Title;
                values[excelRow, 3] = row.FullTimePartTime;
                values[excelRow, 4] = row.EmploymentStatus;
                values[excelRow, 5] = row.HireDate.HasValue ? row.HireDate.Value.ToOADate() : "";
                values[excelRow, 6] = row.IsSalaried ? "Annual Salary" : "Hourly";
                values[excelRow, 7] = row.IsSalaried
                    ? Math.Round(row.AnnualSalary, 2)
                    : (row.HighestHourlyRate > 0.01f ? Math.Round(row.HighestHourlyRate, 2) : "");
                values[excelRow, 8] = row.ChristmasBonus > 0.01f ? Math.Round(row.ChristmasBonus, 2) : "";
                values[excelRow, 9] = row.BusStartingBonus.HasValue
                    ? Math.Round(row.BusStartingBonus.Value, 2)
                    : "";
            }

            Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rows.Count + 1, Headers.Length]];
            range.Value2 = values;
            Excel.Range header = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, Headers.Length]];
            header.Font.Bold = true;
            sheet.Range[sheet.Cells[2, 6], sheet.Cells[Math.Max(2, rows.Count + 1), 6]].NumberFormat = "M/d/yyyy";
            sheet.Range[sheet.Cells[2, 8], sheet.Cells[Math.Max(2, rows.Count + 1), 10]].NumberFormat = "$#,##0.00";
            try
            {
                sheet.Activate();
                sheet.Application.ActiveWindow.SplitRow = 1;
                sheet.Application.ActiveWindow.FreezePanes = true;
            }
            catch
            {
            }
            sheet.Columns.AutoFit();
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

        private static bool TryGetInt(Dictionary<string, int> columns, string[] row, out int number,
            params string[] headers)
        {
            return PayrollHistoryValueParser.TryGetInt(Get(columns, row, headers), out number);
        }

        private sealed class CensusRow
        {
            public int EmployeeNumber { get; init; }
            public string FirstName { get; init; } = "";
            public string LastName { get; init; } = "";
            public string Name { get; init; } = "";
            public string Title { get; init; } = "";
            public string FullTimePartTime { get; init; } = "";
            public string EmploymentStatus { get; init; } = "";
            public DateTime? HireDate { get; init; }
            public bool IsSalaried { get; init; }
            public float AnnualSalary { get; init; }
            public float HighestHourlyRate { get; init; }
            public float ChristmasBonus { get; init; }
            public float? BusStartingBonus { get; init; }
            public bool IsActive { get; init; }
        }

        private sealed class CensusWfnFields
        {
            public string Title { get; init; } = "";
            public string FullTimePartTime { get; init; } = "";
            public string PositionStatus { get; init; } = "";
            public bool IsGrandForks { get; init; }
            public int YearsOfService { get; init; }
        }

        private sealed record WfnCensusSourceRow(string PrimaryPosition, string PositionStatus, string WorkerCategory,
            string Title, bool IsGrandForks, int YearsOfService);
    }
}
