using Excel = Microsoft.Office.Interop.Excel;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class ManualEntry
    {
        public int RowNumber;
        public int EmployeeNumber;
        public Employee? Employee;
        public string EmployeeFirstName = "";
        public float VacationHours;
        public float RoundUpVacationHours;
        public float BackpayHours;
        public float BackpayDollars;
        public string JobtypeText = "";
        public Jobs? Jobtype;
        public float MgHours;
        public float BonusDollars;
        public float Expense;
        public bool ShouldAddVacationToRoundUp;
        public bool IsForCoaches;
        public Company Company;
        public float VacationPayRate;
        public float JobPayRate;

        public bool HasAnyValues()
        {
            return VacationHours > 0.001f
                || BackpayHours > 0.001f
                || MgHours > 0.001f
                || BonusDollars > 0.001f
                || Expense > 0.001f
                || ShouldAddVacationToRoundUp;
        }

        /// <summary>
        /// Hours that count toward vacation accrual, mirroring Shift.AllHours()
        /// (vacation and MG hours count; backpay is a dollar earning and does not).
        /// </summary>
        public float AllHours()
        {
            return VacationHours + RoundUpVacationHours + MgHours;
        }
    }

    internal class ManualEntriesTracker
    {
        private const string FileName = "manual_entries.xlsx";
        private const string BackupFileName = "backup_manual_entries.xlsx";
        private const int StaleFileDays = 2;

        private static ManualEntriesTracker? Instance;
        private DateTime FirstDayWeek2;
        private readonly List<ManualEntry> Entries = new();

        public static ManualEntriesTracker GetInstance()
        {
            Instance ??= new ManualEntriesTracker();
            return Instance;
        }

        public List<ManualEntry> GetEntries()
        {
            return Entries;
        }

        public void Read(DateTime firstDayWeek2)
        {
            FirstDayWeek2 = firstDayWeek2;
            Entries.Clear();

            string filePath = DesktopPath() + FileName;
            if (!File.Exists(filePath))
            {
                Log("No manual_entries.xlsx found on desktop; skipping manual entries.");
                return;
            }

            BackupFile(filePath);

            DateTime lastModified = File.GetLastWriteTime(filePath);
            bool hasSpreadsheetValues = false;

            Excel.Application excelApp = new();
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            try
            {
                foreach (Excel.Worksheet sheet in workBook.Worksheets)
                {
                    Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["Z5000"]].CurrentRegion;
                    object[,] cellData = (object[,])range.Value2;
                    int rows = cellData.GetLength(0);
                    int cols = cellData.GetLength(1);

                    int headerRow = FindHeaderRow(cellData, rows, cols);
                    if (headerRow == 0)
                    {
                        Log("Could not find header row in " + FileName, true);
                        continue;
                    }

                    List<string> headers = ReadHeaders(cellData, headerRow, cols);
                    int employeeNumberCol = FindColumn(headers, "Employee Number");
                    int firstNameCol = FindColumn(headers, "Employee First Name");
                    int vacationCol = FindColumn(headers, "Vacation");
                    int backpayHoursCol = FindColumn(headers, "Backpay Hours");
                    int jobtypeCol = FindColumn(headers, "Jobtype");
                    int mgHoursCol = FindColumn(headers, "MG Hours");
                    int bonusDollarsCol = FindColumn(headers, "Bonus Dollars");
                    int expenseCol = FindColumn(headers, "Expense");
                    int isForCoachesCol = FindColumn(headers, "Is For Coaches");
                    int shouldAddVacationToRoundUpCol = FindColumn(headers, "Should Add Vacation To Round Up");

                    if (employeeNumberCol == -1)
                    {
                        Log("Missing required Employee Number column in " + FileName, true);
                        continue;
                    }

                    for (int row = headerRow + 1; row <= rows; row++)
                    {
                        if (!TryGetInt(cellData[row, employeeNumberCol + 1], out int employeeNumber))
                        {
                            continue;
                        }

                        string isForCoachesValue = isForCoachesCol == -1 ? "" : CellString(cellData[row, isForCoachesCol + 1]);
                        bool isForCoaches = IsYesValue(isForCoachesValue);
                        ManualEntry entry = new()
                        {
                            EmployeeNumber = employeeNumber,
                            RowNumber = row,
                            IsForCoaches = isForCoaches,
                            Company = isForCoaches ? Company.VALLEY_BUS_COACHES : Company.VALLEY_BUS_LLC
                        };

                        if (firstNameCol != -1)
                        {
                            entry.EmployeeFirstName = CellString(cellData[row, firstNameCol + 1]);
                        }
                        if (vacationCol != -1 && TryGetFloat(cellData[row, vacationCol + 1], out float vacationHours))
                        {
                            entry.VacationHours = vacationHours;
                        }
                        if (backpayHoursCol != -1 && TryGetFloat(cellData[row, backpayHoursCol + 1], out float backpayHours))
                        {
                            entry.BackpayHours = backpayHours;
                        }
                        if (jobtypeCol != -1)
                        {
                            entry.JobtypeText = CellString(cellData[row, jobtypeCol + 1]);
                        }
                        if (mgHoursCol != -1 && TryGetFloat(cellData[row, mgHoursCol + 1], out float mgHours))
                        {
                            entry.MgHours = mgHours;
                        }
                        if (bonusDollarsCol != -1 && TryGetFloat(cellData[row, bonusDollarsCol + 1], out float bonusDollars))
                        {
                            entry.BonusDollars = bonusDollars;
                        }
                        if (expenseCol != -1 && TryGetFloat(cellData[row, expenseCol + 1], out float expense))
                        {
                            entry.Expense = expense;
                        }
                        if (shouldAddVacationToRoundUpCol != -1)
                        {
                            entry.ShouldAddVacationToRoundUp = IsYesValue(CellString(cellData[row, shouldAddVacationToRoundUpCol + 1]));
                        }

                        if (!hasSpreadsheetValues && RowHasSpreadsheetValues(cellData, row, isForCoachesValue,
                            firstNameCol, vacationCol, backpayHoursCol, jobtypeCol, mgHoursCol, bonusDollarsCol,
                            expenseCol, shouldAddVacationToRoundUpCol))
                        {
                            hasSpreadsheetValues = true;
                        }

                        if (!entry.HasAnyValues())
                        {
                            continue;
                        }

                        ValidateEntry(entry);

                        if (entry.Employee != null)
                        {
                            CalculatePayRates(entry, entry.Employee);
                            entry.Employee.ManualEntries.Add(entry);
                        }

                        Entries.Add(entry);
                    }
                }
            }
            finally
            {
                workBook.Close(false);
                excelApp.Quit();
            }

            if (hasSpreadsheetValues)
            {
                DateTime today = new(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
                DateTime lastModifiedDate = new(lastModified.Year, lastModified.Month, lastModified.Day);
                if (lastModifiedDate.CompareTo(today.AddDays(-StaleFileDays)) < 0)
                {
                    Log("manual_entries.xlsx has not been edited in the last couple of days.", true);
                }
            }

            Log("Loaded " + Entries.Count + " manual entr" + (Entries.Count == 1 ? "y" : "ies") + " from " + FileName + ".");
        }

        /// <summary>
        /// Calculates pay rates for a manual entry the same way they are calculated for shifts,
        /// by passing a temporary shift (never added to the employee) through Employee.GetPayRateForShift().
        /// </summary>
        private void CalculatePayRates(ManualEntry entry, Employee employee)
        {
            if (entry.VacationHours > 0.001f || entry.ShouldAddVacationToRoundUp)
            {
                Shift temporaryShift = new(entry.Company, Jobs.VACATION)
                {
                    ShiftTime = entry.VacationHours,
                    Date = FirstDayWeek2.AddDays(-7)
                };
                entry.VacationPayRate = employee.GetPayRateForShift(temporaryShift);
                if (entry.VacationPayRate < 0.001f)
                {
                    Log("Manual entry row " + entry.RowNumber + ": no vacation pay rate found for "
                        + employee.Name + " (" + employee.IdNumber + ").", true);
                }
            }

            if (entry.Jobtype.HasValue)
            {
                Shift temporaryShift = new(entry.Company, entry.Jobtype.Value)
                {
                    ShiftTime = entry.MgHours + entry.BackpayHours,
                    Date = FirstDayWeek2.AddDays(-7)
                };
                entry.JobPayRate = employee.GetPayRateForShift(temporaryShift);
                if (entry.JobPayRate < 0.001f && (entry.MgHours > 0.001f || entry.BackpayHours > 0.001f))
                {
                    Log("Manual entry row " + entry.RowNumber + ": no pay rate found for "
                        + employee.Name + " (" + employee.IdNumber + ") for " + entry.Jobtype.Value + ".", true);
                }
            }

            if (entry.BackpayHours > 0.001f)
            {
                if (entry.JobPayRate > 0.001f)
                {
                    entry.BackpayDollars = (float)Math.Round(entry.BackpayHours * entry.JobPayRate, 2);
                }
                else
                {
                    Log("Manual backpay for " + entry.EmployeeNumber + " has no pay rate; using hours as the earnings amount.", true);
                    entry.BackpayDollars = entry.BackpayHours;
                }
            }
        }

        /// <summary>
        /// For entries flagged with "Should Add Vacation To Round Up", adds enough vacation hours to bring the
        /// employee up to the minimum compensated hours required to accrue vacation. Must run after
        /// TotalUpShiftsForEmployees() so all other inputs are reflected in the employee's totals.
        /// </summary>
        public void ApplyVacationRoundUp()
        {
            foreach (ManualEntry entry in Entries)
            {
                if (!entry.ShouldAddVacationToRoundUp || entry.Employee == null)
                {
                    continue;
                }

                Employee employee = entry.Employee;
                if (employee.IsSalaried || employee.IsPartialEntry || EmployeeIdsToIgnore.Contains(employee.IdNumber))
                {
                    continue;
                }

                float compensatedHours = VacationTracker.GetCompensatedHoursForPayPeriod(employee);
                if (compensatedHours >= VacationTracker.MinimumCompensatedHoursForAccrual)
                {
                    Log("Vacation round-up skipped for " + employee.Name + " (" + employee.IdNumber
                        + "): already at " + Math.Round(compensatedHours, 2) + " compensated hours.");
                    continue;
                }

                float hoursToAdd = (float)Math.Round(
                    VacationTracker.MinimumCompensatedHoursForAccrual - compensatedHours, 2);
                if (hoursToAdd < 0.001f)
                {
                    continue;
                }

                entry.RoundUpVacationHours = hoursToAdd;
                Log("Vacation round-up: " + employee.Name + " (" + employee.IdNumber + ") — added "
                    + hoursToAdd + " vacation hours to reach "
                    + VacationTracker.MinimumCompensatedHoursForAccrual + " compensated hours (was "
                    + Math.Round(compensatedHours, 2) + ").");
            }
        }

        private static bool RowHasSpreadsheetValues(object[,] cellData, int row, string isForCoachesValue,
            int firstNameCol, int vacationCol, int backpayHoursCol, int jobtypeCol, int mgHoursCol, int bonusDollarsCol,
            int expenseCol, int shouldAddVacationToRoundUpCol)
        {
            return HasCellValue(cellData, row, firstNameCol)
                || HasCellValue(cellData, row, vacationCol)
                || HasCellValue(cellData, row, backpayHoursCol)
                || HasCellValue(cellData, row, jobtypeCol)
                || HasCellValue(cellData, row, mgHoursCol)
                || HasCellValue(cellData, row, bonusDollarsCol)
                || HasCellValue(cellData, row, expenseCol)
                || HasCellValue(cellData, row, shouldAddVacationToRoundUpCol)
                || !string.IsNullOrWhiteSpace(isForCoachesValue);
        }

        private static bool HasCellValue(object[,] cellData, int row, int col)
        {
            if (col == -1)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(CellString(cellData[row, col + 1]));
        }

        public static bool FirstNameMatchesEmployee(Employee employee, string firstNameFromSheet)
        {
            if (string.IsNullOrWhiteSpace(firstNameFromSheet))
            {
                return true;
            }

            return StringSearch(employee.Name, firstNameFromSheet.Trim());
        }

        private void ValidateEntry(ManualEntry entry)
        {
            if (EmployeeDictionary.TryGetValue(entry.EmployeeNumber, out Employee? employee))
            {
                entry.Employee = employee;
                if (!FirstNameMatchesEmployee(employee, entry.EmployeeFirstName))
                {
                    Log("Manual entry row " + entry.RowNumber + ": Employee First Name \""
                        + entry.EmployeeFirstName + "\" does not match employee "
                        + employee.Name + " (" + employee.IdNumber + ").", true);
                }
            }
            else
            {
                Log("Manual entry row " + entry.RowNumber + ": Employee Number "
                    + entry.EmployeeNumber + " was not found.", true);
            }

            bool needsJobtype = entry.BackpayHours > 0.001f || entry.MgHours > 0.001f;
            if (needsJobtype && string.IsNullOrWhiteSpace(entry.JobtypeText))
            {
                Log("Manual entry row " + entry.RowNumber + ": Backpay Hours or MG Hours entered without Jobtype for employee "
                    + entry.EmployeeNumber + ".", true);
                return;
            }

            if (!string.IsNullOrWhiteSpace(entry.JobtypeText))
            {
                if (TryParseJobType(entry.JobtypeText, out Jobs jobType))
                {
                    entry.Jobtype = jobType;
                }
                else
                {
                    Log("Manual entry row " + entry.RowNumber + ": Jobtype \""
                        + entry.JobtypeText + "\" does not match a Jobs enum value for employee "
                        + entry.EmployeeNumber + ".", true);
                }
            }
        }

        private static bool TryParseJobType(string value, out Jobs jobType)
        {
            jobType = default;
            value = value.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (Enum.TryParse(value, true, out jobType) && Enum.IsDefined(typeof(Jobs), jobType))
            {
                return true;
            }

            string normalized = value.Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
            if (Enum.TryParse(normalized, true, out jobType) && Enum.IsDefined(typeof(Jobs), jobType))
            {
                return true;
            }

            if (int.TryParse(value, out int jobCode) && Enum.IsDefined(typeof(Jobs), jobCode))
            {
                jobType = (Jobs)jobCode;
                return true;
            }

            return false;
        }

        private static bool IsYesValue(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            value = value.Trim();
            return value.Equals("Y", StringComparison.OrdinalIgnoreCase)
                || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }

        private static void BackupFile(string filePath)
        {
            string backupPath = DesktopPath() + BackupFileName;
            File.Copy(filePath, backupPath, true);
            Log("Backed up " + FileName + " to " + BackupFileName + ".");
        }

        private static int FindHeaderRow(object[,] cellData, int rows, int cols)
        {
            for (int row = 1; row <= Math.Min(rows, 10); row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    if (CellString(cellData[row, col]).Equals("Employee Number", StringComparison.OrdinalIgnoreCase))
                    {
                        return row;
                    }
                }
            }

            return 0;
        }

        private static List<string> ReadHeaders(object[,] cellData, int headerRow, int cols)
        {
            List<string> headers = new();
            for (int col = 1; col <= cols; col++)
            {
                headers.Add(CellString(cellData[headerRow, col]));
            }
            return headers;
        }

        private static int FindColumn(List<string> headers, string headerName)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i].Trim().Equals(headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static string CellString(object? cell)
        {
            return cell?.ToString()?.Trim() ?? "";
        }

        private static bool TryGetInt(object? cell, out int value)
        {
            value = 0;
            if (cell == null)
            {
                return false;
            }

            if (cell is double d)
            {
                value = (int)d;
                return true;
            }

            string text = CellString(cell).Replace(",", "").Replace(".0", "");
            return int.TryParse(text, out value);
        }

        private static bool TryGetFloat(object? cell, out float value)
        {
            value = 0f;
            if (cell == null)
            {
                return false;
            }

            if (cell is double d)
            {
                value = (float)d;
                return true;
            }

            string text = CellString(cell).Replace(",", "");
            return float.TryParse(text, out value);
        }
    }
}
