using System.Diagnostics;
using static PayrollProcessor.Program;
using Excel = Microsoft.Office.Interop.Excel;

namespace PayrollProcessor
{
    internal class WsiTracker
    {
        private static readonly HashSet<int> AccountingClassEmployeeNumbers = new() { 1335, 1415, 250, 2183 };
        private static readonly HashSet<int> MechanicClassSalariedEmployeeNumbers = new() { 1355, 992 };
        private static readonly HashSet<int> WashBayClassSalariedEmployeeNumbers = new() { 1778 };

        public void RunIfApplicable(DateTime payDate, bool isPrimaryPayrollRun)
        {
            if (!isPrimaryPayrollRun)
            {
                return;
            }

            List<(DateTime PayDate, string Path)> historyFiles = EmployeePayrollHistory.EnumerateHistoryFiles()
                .OrderBy(file => file.PayDate)
                .ToList();
            if (!TryGetQuarterJustEnded(payDate, historyFiles, out int quarter, out int year, out DateTime quarterEnd))
            {
                return;
            }

            DateTime previousQuarterEnd = PreviousQuarterEnd(quarterEnd);
            DateTime windowStartExclusive = historyFiles
                .Select(file => file.PayDate.Date)
                .Where(historyPayDate => historyPayDate > previousQuarterEnd.Date && historyPayDate < payDate.Date)
                .DefaultIfEmpty(previousQuarterEnd.Date)
                .Min();
            List<(DateTime PayDate, string Path)> quarterFiles = historyFiles
                .Where(file => file.PayDate.Date > windowStartExclusive && file.PayDate.Date <= payDate.Date)
                .ToList();
            if (quarterFiles.Count == 0)
            {
                Log("WSI quarterly report was not created because no payroll history files were found for the quarter.", true);
                return;
            }

            Dictionary<int, WsiEmployeeTotals> totalsByEmployee = new();
            foreach ((DateTime historyPayDate, string path) in quarterFiles)
            {
                if (!EmployeePayrollHistory.TryReadEntries(path, historyPayDate, out List<EmployeePayrollHistory.Entry> entries,
                    out _))
                {
                    Log("WSI quarterly report skipped payroll history file that could not be fully loaded: " + path, true);
                    continue;
                }

                foreach (IGrouping<int, EmployeePayrollHistory.Entry> employeeEntries in entries
                    .Where(entry => entry.Company == Company.VALLEY_BUS_LLC)
                    .GroupBy(entry => entry.EmployeeNumber))
                {
                    if (!totalsByEmployee.TryGetValue(employeeEntries.Key, out WsiEmployeeTotals? totals))
                    {
                        totals = new WsiEmployeeTotals(employeeEntries.Key);
                        totalsByEmployee[employeeEntries.Key] = totals;
                    }

                    float payrollGrossPay = employeeEntries.Sum(entry => entry.TotalCompensation);
                    if (EmployeeDictionary.TryGetValue(employeeEntries.Key, out Employee? historicalEmployee)
                        && EmployeePayrollHistory.ShouldIncludeSalary(historicalEmployee)
                        && historicalEmployee.PrimaryCompany == Company.VALLEY_BUS_LLC)
                    {
                        payrollGrossPay += EmployeePayrollHistory.GetPerPayPeriodSalary(historicalEmployee);
                    }
                    totals.GrossPayroll += payrollGrossPay;
                    foreach (EmployeePayrollHistory.Entry entry in employeeEntries)
                    {
                        foreach ((Jobs job, float hours) in entry.HoursByJob)
                        {
                            totals.HoursByJob[job] = totals.HoursByJob.GetValueOrDefault(job) + hours;
                        }
                    }
                }
            }

            List<WsiReportRow> rows = new();
            foreach (WsiEmployeeTotals totals in totalsByEmployee.Values.Where(totals => totals.GrossPayroll > 0.01f))
            {
                EmployeeDictionary.TryGetValue(totals.EmployeeNumber, out Employee? employee);
                if (employee != null)
                {
                    employee.EnsureNameParts();
                }

                rows.Add(new WsiReportRow(
                    DetermineRateClass(employee, totals),
                    employee?.SocialSecurityNumber ?? "",
                    employee?.FirstName ?? "",
                    employee?.MiddleInitial ?? "",
                    employee?.LastName ?? "",
                    totals.GrossPayroll));
            }

            string outputPath = Path.Combine(EmployeePayrollHistory.HistoryFolder, $"WSI_{quarter}_{year}.xlsx");
            WriteReport(outputPath, rows.OrderBy(row => row.LastName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.FirstName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.SocialSecurityNumber, StringComparer.OrdinalIgnoreCase)
                .ToList());
            Log("WSI quarterly report written to " + outputPath + " for quarter " + quarter + " of " + year
                + " (" + rows.Count + " employees).", true);
        }

        private static bool TryGetQuarterJustEnded(DateTime payDate,
            IReadOnlyList<(DateTime PayDate, string Path)> historyFiles, out int quarter, out int year, out DateTime quarterEnd)
        {
            quarter = 0;
            year = 0;
            quarterEnd = default;
            if (!TryGetPrecedingQuarterEnd(payDate, out quarterEnd))
            {
                return false;
            }

            DateTime endedQuarter = quarterEnd;
            bool isFirstPayrollAfterQuarterEnd = !historyFiles.Any(file =>
                file.PayDate.Date > endedQuarter.Date && file.PayDate.Date < payDate.Date);
            if (!isFirstPayrollAfterQuarterEnd)
            {
                return false;
            }

            quarter = quarterEnd.Month switch
            {
                3 => 1,
                6 => 2,
                9 => 3,
                _ => 4
            };
            year = quarterEnd.Year;
            return true;
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
            Log("Latest in TryGetPrecedingQuarterEnd == " + latest.ToString());
            return latest.HasValue;
        }

        private static DateTime PreviousQuarterEnd(DateTime quarterEnd) => quarterEnd.Month switch
        {
            3 => new DateTime(quarterEnd.Year - 1, 12, 31),
            6 => new DateTime(quarterEnd.Year, 3, 31),
            9 => new DateTime(quarterEnd.Year, 6, 30),
            _ => new DateTime(quarterEnd.Year, 9, 30)
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
                    return 8292;
                }
                if (IsRatioLessThan8To1(mechanicHours, schoolHours))
                {
                    return 8292;
                }
                return 3630;
            }

            if (adminHours > 0.01f && IsRatioLessThan8To1(schoolHours, adminHours))
            {
                return 8805;
            }
            if (bodyShopHours > 0.01f && IsRatioLessThan8To1(schoolHours, bodyShopHours))
            {
                return 8292;
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

        private sealed class WsiEmployeeTotals
        {
            public int EmployeeNumber { get; }
            public float GrossPayroll { get; set; }
            public Dictionary<Jobs, float> HoursByJob { get; } = new();

            public WsiEmployeeTotals(int employeeNumber)
            {
                EmployeeNumber = employeeNumber;
            }
        }

        private sealed record WsiReportRow(int RateClass, string SocialSecurityNumber, string FirstName, string MiddleInitial,
            string LastName, float GrossPayroll);
    }
}
