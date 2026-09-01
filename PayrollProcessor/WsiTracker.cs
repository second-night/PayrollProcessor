using System.Diagnostics;
using System.Globalization;
using static PayrollProcessor.Program;
using Excel = Microsoft.Office.Interop.Excel;

namespace PayrollProcessor
{
    internal class WsiTracker
    {
        private const string AdpPayrollHistoryFileName = "AdpPayrollHistory.xlsx";

        private static readonly HashSet<int> AccountingClassEmployeeNumbers = new() { 1335, 1415, 250, 2183 };
        private static readonly HashSet<int> MechanicClassSalariedEmployeeNumbers = new() { 1355, 992 };
        private static readonly HashSet<int> WashBayClassSalariedEmployeeNumbers = new() { 1778 };

        public void RunIfApplicable(DateTime payDate)
        {
            if (!TryGetPrecedingQuarterEnd(payDate, out DateTime quarterEnd) || payDate.Date <= quarterEnd.Date)
            {
                return;
            }

            int quarter = QuarterNumber(quarterEnd);
            int year = quarterEnd.Year;
            string outputPath = Path.Combine(EmployeePayrollHistory.HistoryFolder, $"WSI_{quarter}_{year}.xlsx");
            if (File.Exists(outputPath))
            {
                Log("WSI quarterly report already exists and was not regenerated: " + outputPath);
                return;
            }

            DateTime previousQuarterEnd = PreviousQuarterEnd(quarterEnd);
            DateTime quarterStart = previousQuarterEnd.AddDays(1);
            if (!PrintForm.InputBool(
                    "WSI quarter " + quarter + " of " + year + " is ready to generate."
                    + "\n\nDownload a current copy of " + AdpPayrollHistoryFileName
                    + " (Valley Bus LLC, pay dates " + quarterStart.ToString("M/d/yyyy")
                    + " through " + quarterEnd.ToString("M/d/yyyy") + ") to your desktop, then click Ready.",
                    "Ready",
                    "Skip"))
            {
                Log("WSI quarterly report was skipped. Download " + AdpPayrollHistoryFileName
                    + " and it will run on the next primary payroll.", true);
                return;
            }

            if (!TryReadGrossWages(previousQuarterEnd, quarterEnd, out Dictionary<int, float> grossByEmployee,
                out Dictionary<int, WageFileName> namesByEmployee))
            {
                return;
            }

            Dictionary<int, WsiEmployeeTotals> hoursByEmployee = LoadHoursFromHistory(previousQuarterEnd, quarterEnd);
            List<WsiReportRow> rows = new();
            foreach ((int employeeNumber, float grossPayroll) in grossByEmployee.Where(entry => entry.Value > 0.01f))
            {
                if (!hoursByEmployee.TryGetValue(employeeNumber, out WsiEmployeeTotals? totals))
                {
                    totals = new WsiEmployeeTotals(employeeNumber);
                }

                EmployeeDictionary.TryGetValue(employeeNumber, out Employee? employee);
                namesByEmployee.TryGetValue(employeeNumber, out WageFileName? fileName);
                if (employee != null)
                {
                    employee.EnsureNameParts();
                }

                rows.Add(new WsiReportRow(
                    DetermineRateClass(employee, totals),
                    FirstNonEmpty(employee?.SocialSecurityNumber, fileName?.SocialSecurityNumber),
                    FirstNonEmpty(employee?.FirstName, fileName?.FirstName),
                    FirstNonEmpty(employee?.MiddleInitial, fileName?.MiddleInitial),
                    FirstNonEmpty(employee?.LastName, fileName?.LastName),
                    grossPayroll));
            }

            if (rows.Count == 0)
            {
                Log("WSI quarterly report was not created because no Valley Bus LLC wages were found in "
                    + AdpPayrollHistoryFileName + ".", true);
                return;
            }

            WriteReport(outputPath, rows.OrderBy(row => row.LastName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.FirstName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.SocialSecurityNumber, StringComparer.OrdinalIgnoreCase)
                .ToList());
            Log("WSI quarterly report written to " + outputPath + " for quarter " + quarter + " of " + year
                + " (" + rows.Count + " employees).", true);
        }

        private static bool TryReadGrossWages(DateTime previousQuarterEnd, DateTime quarterEnd,
            out Dictionary<int, float> grossByEmployee, out Dictionary<int, WageFileName> namesByEmployee)
        {
            grossByEmployee = new();
            namesByEmployee = new();
            if (!TryFindWageFile(AdpPayrollHistoryFileName, out string adpPath))
            {
                return false;
            }

            Excel.Application excelApp = new()
            {
                DisplayAlerts = false
            };
            try
            {
                ReadAdpGrossWages(excelApp, adpPath, previousQuarterEnd, quarterEnd, grossByEmployee, namesByEmployee);
            }
            catch (Exception)
            {
                Log("Error reading " + AdpPayrollHistoryFileName
                    + ". Please make sure the file is not open and run the process again.", true);
                return false;
            }
            finally
            {
                excelApp.Quit();
            }

            if (grossByEmployee.Count == 0)
            {
                Log("WSI quarterly report was not created because no Valley Bus LLC wages were found in "
                    + AdpPayrollHistoryFileName + ".", true);
                return false;
            }

            Log("WSI wage file loaded " + grossByEmployee.Count + " employees totaling "
                + Math.Round(grossByEmployee.Values.Sum(), 2).ToString("0.00") + ".", true);
            return true;
        }

        private static void ReadAdpGrossWages(Excel.Application excelApp, string path, DateTime previousQuarterEnd,
            DateTime quarterEnd, Dictionary<int, float> grossByEmployee, Dictionary<int, WageFileName> namesByEmployee)
        {
            Excel.Workbook? workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(path);
                Excel.Worksheet sheet = FindWorksheet(workbook, "Payroll History")
                    ?? (Excel.Worksheet)workbook.Worksheets[1];
                if (!TryGetSheetData(sheet, out object[,] cellData, out int rows, out int cols))
                {
                    Log("WSI could not read " + AdpPayrollHistoryFileName + ".", true);
                    return;
                }

                List<string> headers = ReadHeaders(cellData, cols);
                int employeeColumn = FindColumn(headers, "FILE NUMBER");
                if (employeeColumn < 0)
                {
                    employeeColumn = FindColumn(headers, "File Number");
                }
                int grossColumn = FindColumn(headers, "GROSS PAY");
                if (grossColumn < 0)
                {
                    grossColumn = FindColumn(headers, "Gross Pay");
                }
                int payDateColumn = FindColumn(headers, "PAY DATE");
                if (payDateColumn < 0)
                {
                    payDateColumn = FindColumn(headers, "Pay Date");
                }
                int nameColumn = FindColumn(headers, "NAME");
                int companyColumn = FindCompanyColumn(headers);
                if (employeeColumn < 0 || grossColumn < 0 || payDateColumn < 0)
                {
                    Log("WSI could not find FILE NUMBER / PAY DATE / GROSS PAY columns in " + AdpPayrollHistoryFileName
                        + ".", true);
                    return;
                }

                float total = 0f;
                HashSet<int> employeesWithWages = new();
                for (int row = 2; row <= rows; row++)
                {
                    if (!TryGetIntFromCell(cellData[row, employeeColumn + 1], out int employeeNumber)
                        || employeeNumber <= 0
                        || !TryGetDateFromCell(cellData[row, payDateColumn + 1], out DateTime rowPayDate)
                        || rowPayDate.Date <= previousQuarterEnd.Date
                        || rowPayDate.Date > quarterEnd.Date
                        || !IsValleyBusLlcCompany(companyColumn < 0 ? "" : CellString(cellData[row, companyColumn + 1]))
                        || !TryGetFloatFromCell(cellData[row, grossColumn + 1], out float gross)
                        || gross <= 0.01f)
                    {
                        continue;
                    }

                    AddGross(grossByEmployee, employeeNumber, gross);
                    total += gross;
                    employeesWithWages.Add(employeeNumber);
                    if (nameColumn >= 0)
                    {
                        ParseAdpName(CellString(cellData[row, nameColumn + 1]), out string firstName, out string middleName,
                            out string lastName);
                        RememberName(namesByEmployee, employeeNumber, "", firstName, middleName, lastName);
                    }
                }

                Log("WSI read " + AdpPayrollHistoryFileName + " for pay dates "
                    + previousQuarterEnd.AddDays(1).ToString("M/d/yyyy") + " through " + quarterEnd.ToString("M/d/yyyy")
                    + ": " + employeesWithWages.Count + " employees, " + Math.Round(total, 2).ToString("0.00") + ".");
            }
            finally
            {
                workbook?.Close(false);
            }
        }

        private static Dictionary<int, WsiEmployeeTotals> LoadHoursFromHistory(DateTime previousQuarterEnd,
            DateTime quarterEnd)
        {
            Dictionary<int, WsiEmployeeTotals> totalsByEmployee = new();
            List<(DateTime PayDate, string Path)> quarterFiles = EmployeePayrollHistory.EnumerateHistoryFiles()
                .Where(file => file.PayDate.Date > previousQuarterEnd.Date && file.PayDate.Date <= quarterEnd.Date)
                .OrderBy(file => file.PayDate)
                .ToList();
            foreach ((DateTime historyPayDate, string path) in quarterFiles)
            {
                if (!EmployeePayrollHistory.TryReadEntries(path, historyPayDate, out List<EmployeePayrollHistory.Entry> entries,
                    out _))
                {
                    Log("WSI quarterly report skipped payroll history file that could not be fully loaded: " + path, true);
                    continue;
                }

                foreach (EmployeePayrollHistory.Entry entry in entries.Where(entry => entry.Company == Company.VALLEY_BUS_LLC))
                {
                    if (!totalsByEmployee.TryGetValue(entry.EmployeeNumber, out WsiEmployeeTotals? totals))
                    {
                        totals = new WsiEmployeeTotals(entry.EmployeeNumber);
                        totalsByEmployee[entry.EmployeeNumber] = totals;
                    }

                    foreach ((Jobs job, float hours) in entry.HoursByJob)
                    {
                        totals.HoursByJob[job] = totals.HoursByJob.GetValueOrDefault(job) + hours;
                    }
                }
            }

            return totalsByEmployee;
        }

        private static bool TryGetPrecedingQuarterEnd(DateTime payDate, out DateTime quarterEnd)
        {
            int year = payDate.Year;
            DateTime[] quarterEnds =
            {
                new(year - 1, 12, 31),
                new(year, 3, 31),
                new(year, 6, 30),
                new(year, 9, 30),
                new(year, 12, 31)
            };
            DateTime? latest = null;
            foreach (DateTime candidate in quarterEnds)
            {
                if (candidate.Date < payDate.Date && (!latest.HasValue || candidate > latest.Value))
                {
                    latest = candidate;
                }
            }

            quarterEnd = latest ?? default;
            return latest.HasValue;
        }

        private static DateTime PreviousQuarterEnd(DateTime quarterEnd) => quarterEnd.Month switch
        {
            3 => new DateTime(quarterEnd.Year - 1, 12, 31),
            6 => new DateTime(quarterEnd.Year, 3, 31),
            9 => new DateTime(quarterEnd.Year, 6, 30),
            _ => new DateTime(quarterEnd.Year, 9, 30)
        };

        private static int QuarterNumber(DateTime quarterEnd) => quarterEnd.Month switch
        {
            3 => 1,
            6 => 2,
            9 => 3,
            _ => 4
        };

        private static int DetermineRateClass(Employee? employee, WsiEmployeeTotals totals)
        {
            if (employee != null && (employee.IsSalaried || employee.AnnualSalaryAmount > 0.001f))
            {
                if (AccountingClassEmployeeNumbers.Contains(employee.IdNumber))
                {
                    return 8747;
                }
                if (MechanicClassSalariedEmployeeNumbers.Contains(employee.IdNumber))
                {
                    return 3630;
                }
                if (WashBayClassSalariedEmployeeNumbers.Contains((int)employee.IdNumber))
                {
                    return 8380;
                }
                return 8805;
            }

            float mechanicHours = GetHours(totals.HoursByJob, Jobs.MECHANIC);
            float driverSchoolHours = GetHours(totals.HoursByJob, Jobs.DRIVER_SCHOOL);
            float aideSchoolHours = GetHours(totals.HoursByJob, Jobs.AIDE_SCHOOL);
            float schoolHours = driverSchoolHours + aideSchoolHours;
            float adminHours = GetHours(totals.HoursByJob, Jobs.ADMIN);
            float bodyShopHours = GetHours(totals.HoursByJob, Jobs.BODY_SHOP);
            float washBayHours = GetHours(totals.HoursByJob, Jobs.WASH_BAY, Jobs.WASH_BAY_OT);

            if (mechanicHours > 0.01f)
            {
                float driverRate = employee == null
                    ? 0f
                    : Math.Max(employee.PayRates.GetValueOrDefault(Jobs.DRIVER_SCHOOL),
                        employee.PayRates.GetValueOrDefault(Jobs.NON_CDL_DRIVER));
                float mechanicRate = employee?.PayRates.GetValueOrDefault(Jobs.MECHANIC) ?? 0f;
                if (driverRate > 0.001f && driverRate <= mechanicRate)
                {
                    return 8010;
                }
                if (IsRatioLessThan8To1(mechanicHours, schoolHours))
                {
                    return 8010;
                }
                return 3630;
            }

            if (adminHours > 0.01f && IsRatioLessThan8To1(schoolHours, adminHours))
            {
                return 8805;
            }
            if (bodyShopHours > 0.01f && IsRatioLessThan8To1(schoolHours, bodyShopHours))
            {
                return 8010;
            }
            if (washBayHours > 0.01f && IsRatioLessThan8To1(schoolHours, washBayHours))
            {
                return 8380;
            }
            return 7380;
        }

        private static float GetHours(Dictionary<Jobs, float> hoursByJob, params Jobs[] jobs)
        {
            float hours = 0f;
            foreach (Jobs job in jobs)
            {
                hours += hoursByJob.GetValueOrDefault(job);
            }
            return hours;
        }

        private static bool IsRatioLessThan8To1(float leadingQuantity, float trailingQuantity)
        {
            if (trailingQuantity <= 0.01f)
            {
                return false;
            }
            return leadingQuantity / trailingQuantity < 8f;
        }

        private static void WriteReport(string path, List<WsiReportRow> rows)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            Excel.Application excelApp = new()
            {
                DisplayAlerts = false
            };
            Excel.Workbook? workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Add();
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets[1];
                sheet.Name = "WSI";

                object[,] output = new object[rows.Count + 1, 6];
                output[0, 0] = "Rate Class";
                output[0, 1] = "Employee's SSN";
                output[0, 2] = "Employee's First Name";
                output[0, 3] = "Employee's Middle Initial";
                output[0, 4] = "Employee's Last Name";
                output[0, 5] = "Gross Payroll";

                for (int row = 0; row < rows.Count; row++)
                {
                    WsiReportRow reportRow = rows[row];
                    output[row + 1, 0] = reportRow.RateClass;
                    output[row + 1, 1] = reportRow.SocialSecurityNumber;
                    output[row + 1, 2] = reportRow.FirstName;
                    output[row + 1, 3] = reportRow.MiddleInitial;
                    output[row + 1, 4] = reportRow.LastName;
                    output[row + 1, 5] = Math.Round(reportRow.GrossPayroll, 2);
                }

                Excel.Range ssnColumn = sheet.Columns[2];
                ssnColumn.NumberFormat = "@";
                Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rows.Count + 1, 6]];
                range.Value2 = output;
                sheet.Columns[6].NumberFormat = "0.00";
                range.Columns.AutoFit();

                workbook.SaveAs(path);
                workbook.Close(true);
                workbook = null;

                if (rows.Count > 0)
                {
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
            }
            catch (Exception)
            {
                Log("Error saving WSI quarterly report " + path + ". Please make sure the file is not open and run the process again.",
                    true);
            }
            finally
            {
                workbook?.Close(false);
                excelApp.Quit();
            }
        }

        private static bool TryFindWageFile(string fileName, out string path)
        {
            string[] candidates =
            {
                DesktopPath() + fileName,
                Path.GetFullPath(Path.Combine(EmployeePayrollHistory.HistoryFolder, "..", fileName))
            };
            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    path = candidate;
                    return true;
                }
            }

            path = "";
            Log("WSI quarterly report needs " + fileName
                + " on the desktop. Download it and run the next primary payroll to generate the report.", true);
            return false;
        }

        private static bool TryGetSheetData(Excel.Worksheet sheet, out object[,] cellData, out int rows, out int cols)
        {
            cellData = new object[0, 0];
            rows = 0;
            cols = 0;
            Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["B2"]].CurrentRegion;
            if (range.Value2 is not object[,] values)
            {
                return false;
            }

            cellData = values;
            rows = cellData.GetLength(0);
            cols = cellData.GetLength(1);
            return rows >= 2 && cols >= 2;
        }

        private static Excel.Worksheet? FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private static List<string> ReadHeaders(object[,] cellData, int cols)
        {
            List<string> headers = new();
            for (int col = 1; col <= cols; col++)
            {
                headers.Add(CellString(cellData[1, col]));
            }
            return headers;
        }

        private static int FindColumn(List<string> headers, string headerName)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i].Equals(headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }

        private static int FindCompanyColumn(List<string> headers)
        {
            int column = FindColumn(headers, "Company Code");
            if (column >= 0)
            {
                return column;
            }
            column = FindColumn(headers, "Payroll Company Code");
            if (column >= 0)
            {
                return column;
            }
            return FindColumn(headers, "Company");
        }

        private static bool IsValleyBusLlcCompany(string company)
        {
            if (string.IsNullOrWhiteSpace(company))
            {
                return true;
            }

            return company.StartsWith("MMF", StringComparison.OrdinalIgnoreCase)
                || company.Contains("Valley Bus LLC", StringComparison.OrdinalIgnoreCase)
                || company.Contains("VALLEY_BUS_LLC", StringComparison.OrdinalIgnoreCase);
        }

        private static void AddGross(Dictionary<int, float> grossByEmployee, int employeeNumber, float gross)
        {
            grossByEmployee[employeeNumber] = grossByEmployee.GetValueOrDefault(employeeNumber) + gross;
        }

        private static void RememberName(Dictionary<int, WageFileName> namesByEmployee, int employeeNumber, string ssn,
            string firstName, string middleName, string lastName)
        {
            if (!namesByEmployee.TryGetValue(employeeNumber, out WageFileName? existing))
            {
                namesByEmployee[employeeNumber] = new WageFileName(ssn, firstName, middleName, lastName);
                return;
            }

            namesByEmployee[employeeNumber] = new WageFileName(
                FirstNonEmpty(existing.SocialSecurityNumber, ssn),
                FirstNonEmpty(existing.FirstName, firstName),
                FirstNonEmpty(existing.MiddleName, middleName),
                FirstNonEmpty(existing.LastName, lastName));
        }

        private static void ParseAdpName(string name, out string firstName, out string middleName, out string lastName)
        {
            firstName = "";
            middleName = "";
            lastName = "";
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            int comma = name.IndexOf(',');
            if (comma < 0)
            {
                string[] unsplit = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (unsplit.Length == 1)
                {
                    firstName = unsplit[0];
                }
                else if (unsplit.Length >= 2)
                {
                    firstName = unsplit[0];
                    lastName = unsplit[^1];
                    if (unsplit.Length > 2)
                    {
                        middleName = string.Join(" ", unsplit.Skip(1).Take(unsplit.Length - 2));
                    }
                }
                return;
            }

            lastName = name[..comma].Trim();
            string[] given = name[(comma + 1)..].Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (given.Length > 0)
            {
                firstName = given[0];
            }
            if (given.Length > 1)
            {
                middleName = string.Join(" ", given.Skip(1));
            }
        }

        private static string FirstNonEmpty(string? first, string? second) =>
            !string.IsNullOrWhiteSpace(first) ? first : second ?? "";

        private static string CellString(object? cell) => cell?.ToString()?.Trim() ?? "";

        private static bool TryGetDateFromCell(object? cellData, out DateTime date)
        {
            date = DateTime.MinValue;
            if (cellData == null)
            {
                return false;
            }

            if (cellData is DateTime dateTime)
            {
                date = dateTime;
                return true;
            }

            string text = CellString(cellData);
            if (text == "")
            {
                return false;
            }

            if (DateTime.TryParseExact(text, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return true;
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate)
                && oaDate > 20000)
            {
                date = DateTime.FromOADate(oaDate);
                return true;
            }

            return false;
        }

        private static bool TryGetFloatFromCell(object? cellData, out float outFloat)
        {
            outFloat = 0f;
            if (cellData == null)
            {
                return false;
            }

            switch (cellData)
            {
                case double value:
                    outFloat = (float)value;
                    return true;
                case float value:
                    outFloat = value;
                    return true;
                case int value:
                    outFloat = value;
                    return true;
                case decimal value:
                    outFloat = (float)value;
                    return true;
            }

            return float.TryParse(CellString(cellData), NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out outFloat);
        }

        private static bool TryGetIntFromCell(object? cellData, out int outInt)
        {
            outInt = 0;
            if (cellData == null)
            {
                return false;
            }

            if (cellData is double value)
            {
                outInt = (int)value;
                return true;
            }

            if (cellData is int intValue)
            {
                outInt = intValue;
                return true;
            }

            string text = CellString(cellData).TrimStart('0');
            if (text == "")
            {
                text = "0";
            }
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out outInt);
        }

        private sealed class WsiEmployeeTotals
        {
            public int EmployeeNumber { get; }
            public Dictionary<Jobs, float> HoursByJob { get; } = new();

            public WsiEmployeeTotals(int employeeNumber)
            {
                EmployeeNumber = employeeNumber;
            }
        }

        private sealed record WageFileName(string SocialSecurityNumber, string FirstName, string MiddleName, string LastName)
        {
            public string MiddleInitial
            {
                get
                {
                    foreach (char character in MiddleName ?? "")
                    {
                        if (char.IsLetter(character))
                        {
                            return char.ToUpperInvariant(character).ToString();
                        }
                    }
                    return "";
                }
            }
        }

        private sealed record WsiReportRow(int RateClass, string SocialSecurityNumber, string FirstName, string MiddleInitial,
            string LastName, float GrossPayroll);
    }
}
