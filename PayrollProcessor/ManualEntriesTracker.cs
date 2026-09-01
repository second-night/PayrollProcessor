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
        public float RegularHours;
        public int WeekNumber;
        public float BonusDollars;
        public float Expense;
        public float? SpecifiedPayRate;
        public bool ShouldAddVacationToRoundUp;
        public bool IsForCoaches;
        public Company Company;
        public float VacationPayRate;
        public float JobPayRate;

        public static bool HasAmount(float value) => Math.Abs(value) > 0.001f;

        public bool HasAnyValues()
        {
            return HasAmount(VacationHours)
                || HasAmount(BackpayHours)
                || HasAmount(BackpayDollars)
                || HasAmount(MgHours)
                || HasAmount(RegularHours)
                || HasAmount(BonusDollars)
                || HasAmount(Expense)
                || ShouldAddVacationToRoundUp;
        }

        /// <summary>
        /// Hours that count toward vacation accrual, mirroring Shift.AllHours()
        /// (vacation, MG, and regular hours count; backpay hours and backpay dollars do not).
        /// </summary>
        public float AllHours()
        {
            return VacationHours + RoundUpVacationHours + MgHours + RegularHours;
        }

        /// <summary>
        /// Mirrors timesheet handling: when jobtype is DRIVER_SCHOOL but the employee has no CDL driver
        /// pay rate, treat the entry as NON_CDL_DRIVER for pay rate and department code purposes.
        /// </summary>
        public Jobs GetResolvedJobType(Employee employee)
        {
            Jobs jobType = Jobtype ?? Jobs.ADMIN;
            if (jobType == Jobs.DRIVER_SCHOOL && !employee.PayRates.ContainsKey(Jobs.DRIVER_SCHOOL))
            {
                return Jobs.NON_CDL_DRIVER;
            }

            return jobType;
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

        /// <summary>
        /// Scans manual_entries.xlsx before Employee Export is read, mirroring PreCheckTimeSheets.
        /// Employees with manual entry values who are not in iSolved get partial entries so
        /// ReadEmployeeExport can import them from Employee Export.xlsx.
        /// </summary>
        public void PreCheckForNewEmployees()
        {
            string filePath = DesktopPath() + FileName;
            if (!File.Exists(filePath))
            {
                return;
            }

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
                        continue;
                    }

                    List<string> headers = ReadHeaders(cellData, headerRow, cols);
                    int employeeNumberCol = FindColumn(headers, "Employee Number");
                    int firstNameCol = FindColumn(headers, "Employee First Name");
                    int vacationCol = FindColumn(headers, "Vacation");
                    int backpayHoursCol = FindColumn(headers, "Backpay Hours");
                    int backpayDollarsCol = FindColumn(headers, "Backpay Dollars");
                    int jobtypeCol = FindColumn(headers, "Jobtype");
                    int payrateCol = FindColumn(headers, "Payrate");
                    int mgHoursCol = FindColumn(headers, "MG Hours");
                    int regularHoursCol = FindColumn(headers, "Regular Hours");
                    int weekNumberCol = FindColumn(headers, "Week Number");
                    int bonusDollarsCol = FindColumn(headers, "Bonus Dollars");
                    int expenseCol = FindColumn(headers, "Expense");
                    int isForCoachesCol = FindColumn(headers, "Is For Coaches");
                    int shouldAddVacationToRoundUpCol = FindColumn(headers, "Should Add Vacation To Round Up");

                    if (employeeNumberCol == -1)
                    {
                        continue;
                    }

                    for (int row = headerRow + 1; row <= rows; row++)
                    {
                        if (!TryGetInt(cellData[row, employeeNumberCol + 1], out int employeeNumber))
                        {
                            continue;
                        }

                        string isForCoachesValue = isForCoachesCol == -1 ? "" : CellString(cellData[row, isForCoachesCol + 1]);
                        if (!RowHasSpreadsheetValues(cellData, row, isForCoachesValue,
                            firstNameCol, vacationCol, backpayHoursCol, backpayDollarsCol, jobtypeCol, payrateCol,
                            mgHoursCol, regularHoursCol, weekNumberCol, bonusDollarsCol, expenseCol,
                            shouldAddVacationToRoundUpCol))
                        {
                            continue;
                        }

                        string employeeFirstName = firstNameCol == -1 ? "" : CellString(cellData[row, firstNameCol + 1]);

                        if (!EmployeeDictionary.ContainsKey(employeeNumber))
                        {
                            Employee emp = new(employeeNumber, employeeFirstName)
                            {
                                IsPartialEntry = true,
                                HadManualEntry = true
                            };
                            EmployeeDictionary.Add(employeeNumber, emp);
                        }
                        else
                        {
                            Employee employee = EmployeeDictionary[employeeNumber];
                            if (!employee.WasAlreadyInPayroll)
                            {
                                employee.HadManualEntry = true;
                            }
                        }
                    }
                }
            }
            finally
            {
                workBook.Close(false);
                excelApp.Quit();
            }
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
                    int backpayDollarsCol = FindColumn(headers, "Backpay Dollars");
                    int jobtypeCol = FindColumn(headers, "Jobtype");
                    int payrateCol = FindColumn(headers, "Payrate");
                    int mgHoursCol = FindColumn(headers, "MG Hours");
                    int regularHoursCol = FindColumn(headers, "Regular Hours");
                    int weekNumberCol = FindColumn(headers, "Week Number");
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
                        if (backpayDollarsCol != -1 && TryGetFloat(cellData[row, backpayDollarsCol + 1], out float backpayDollars))
                        {
                            entry.BackpayDollars = backpayDollars;
                        }
                        if (jobtypeCol != -1)
                        {
                            entry.JobtypeText = CellString(cellData[row, jobtypeCol + 1]);
                        }
                        if (payrateCol != -1 && TryGetFloat(cellData[row, payrateCol + 1], out float payrate) && ManualEntry.HasAmount(payrate))
                        {
                            entry.SpecifiedPayRate = payrate;
                        }
                        if (mgHoursCol != -1 && TryGetFloat(cellData[row, mgHoursCol + 1], out float mgHours))
                        {
                            entry.MgHours = mgHours;
                        }
                        if (regularHoursCol != -1 && TryGetFloat(cellData[row, regularHoursCol + 1], out float regularHours))
                        {
                            entry.RegularHours = regularHours;
                        }
                        if (weekNumberCol != -1 && TryGetInt(cellData[row, weekNumberCol + 1], out int weekNumber))
                        {
                            entry.WeekNumber = weekNumber;
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
                            firstNameCol, vacationCol, backpayHoursCol, backpayDollarsCol, jobtypeCol, payrateCol,
                            mgHoursCol, regularHoursCol, weekNumberCol, bonusDollarsCol, expenseCol,
                            shouldAddVacationToRoundUpCol))
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
        /// Calculates pay rates for a manual entry. A Payrate column value is used as-is.
        /// Otherwise rates are calculated the same way as shifts, by passing a temporary shift
        /// (never added to the employee) through Employee.GetPayRateForShift().
        /// Negative hour amounts use the absolute value only for that lookup.
        /// </summary>
        private void CalculatePayRates(ManualEntry entry, Employee employee)
        {
            if (ManualEntry.HasAmount(entry.VacationHours) || entry.ShouldAddVacationToRoundUp)
            {
                Shift temporaryShift = new(entry.Company, Jobs.VACATION)
                {
                    ShiftTime = Math.Abs(entry.VacationHours),
                    Date = FirstDayWeek2.AddDays(-7)
                };
                entry.VacationPayRate = employee.GetPayRateForShift(temporaryShift);
                if (entry.VacationPayRate < 0.001f)
                {
                    Log("Manual entry row " + entry.RowNumber + ": no vacation pay rate found for "
                        + employee.Name + " (" + employee.IdNumber + ").", true);
                }
            }

            if (entry.SpecifiedPayRate.HasValue && ManualEntry.HasAmount(entry.SpecifiedPayRate.Value))
            {
                entry.JobPayRate = entry.SpecifiedPayRate.Value;
                return;
            }

            if (entry.Jobtype.HasValue)
            {
                float jobHours = Math.Abs(entry.MgHours) + Math.Abs(entry.BackpayHours) + Math.Abs(entry.RegularHours);
                Jobs jobTypeForShift = entry.GetResolvedJobType(employee);
                Shift temporaryShift = new(entry.Company, jobTypeForShift)
                {
                    ShiftTime = jobHours,
                    Date = FirstDayWeek2.AddDays(-7)
                };
                entry.JobPayRate = employee.GetPayRateForShift(temporaryShift);
                if (entry.JobPayRate < 0.001f
                    && (ManualEntry.HasAmount(entry.MgHours) || ManualEntry.HasAmount(entry.BackpayHours)
                        || ManualEntry.HasAmount(entry.RegularHours)))
                {
                    Log("Manual entry row " + entry.RowNumber + ": no pay rate found for "
                        + employee.Name + " (" + employee.IdNumber + ") for " + entry.Jobtype.Value + ".", true);
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
            int firstNameCol, int vacationCol, int backpayHoursCol, int backpayDollarsCol, int jobtypeCol, int payrateCol,
            int mgHoursCol, int regularHoursCol, int weekNumberCol, int bonusDollarsCol, int expenseCol,
            int shouldAddVacationToRoundUpCol)
        {
            return HasCellValue(cellData, row, firstNameCol)
                || HasCellValue(cellData, row, vacationCol)
                || HasCellValue(cellData, row, backpayHoursCol)
                || HasCellValue(cellData, row, backpayDollarsCol)
                || HasCellValue(cellData, row, jobtypeCol)
                || HasCellValue(cellData, row, payrateCol)
                || HasCellValue(cellData, row, mgHoursCol)
                || HasCellValue(cellData, row, regularHoursCol)
                || HasCellValue(cellData, row, weekNumberCol)
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
                //Log("Manual entry row " + entry.RowNumber + ": Employee Number " + entry.EmployeeNumber + " was not found.", true);
                ValidateEntryAgainstEmployeeExport(entry);
            }

            bool needsJobtype = ManualEntry.HasAmount(entry.BackpayHours)
                || ManualEntry.HasAmount(entry.MgHours)
                || ManualEntry.HasAmount(entry.RegularHours);
            if (needsJobtype && string.IsNullOrWhiteSpace(entry.JobtypeText))
            {
                Log("Manual entry row " + entry.RowNumber + ": Backpay Hours, MG Hours, or Regular Hours entered without Jobtype for employee "
                    + entry.EmployeeNumber + ".", true);
            }

            if (ManualEntry.HasAmount(entry.RegularHours))
            {
                if (entry.WeekNumber < 1 || entry.WeekNumber > 2)
                {
                    Log("Manual entry row " + entry.RowNumber + ": Regular Hours entered without a valid Week Number (1 or 2) for employee "
                        + entry.EmployeeNumber + ".", true);
                }
            }
            else if (entry.WeekNumber > 0)
            {
                Log("Manual entry row " + entry.RowNumber + ": Week Number entered without Regular Hours for employee "
                    + entry.EmployeeNumber + ".", true);
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

        private static void ValidateEntryAgainstEmployeeExport(ManualEntry entry)
        {
            if (entry.Employee != null && entry.Employee.WasAlreadyInPayroll)
            {
                return;
            }

            if (ExcelWorker.EmployeeExportByNumber.TryGetValue(entry.EmployeeNumber, out (string FirstName, string LastName) exportNames))
            {
                if (!string.IsNullOrWhiteSpace(entry.EmployeeFirstName) && !string.IsNullOrWhiteSpace(exportNames.FirstName)
                    && !StringSearch(exportNames.FirstName, entry.EmployeeFirstName.Trim()))
                {
                    Log("Manual entry row " + entry.RowNumber + ": Employee First Name \""
                        + entry.EmployeeFirstName + "\" does not match Employee Export first name \""
                        + exportNames.FirstName + "\" for employee #" + entry.EmployeeNumber + ".", true);
                }
                return;
            }

            if (string.IsNullOrWhiteSpace(entry.EmployeeFirstName))
            {
                if (entry.Employee == null || entry.Employee.IsPartialEntry)
                {
                    Log("Manual entry row " + entry.RowNumber + ": Employee Number "
                        + entry.EmployeeNumber + " was not found on Employee Export.", true);
                }
                return;
            }

            foreach (KeyValuePair<int, (string FirstName, string LastName)> exportEntry in ExcelWorker.EmployeeExportByNumber)
            {
                if (StringSearch(exportEntry.Value.FirstName, entry.EmployeeFirstName.Trim()))
                {
                    string exportFullName = exportEntry.Value.FirstName + " " + exportEntry.Value.LastName;
                    Log("Manual entry row " + entry.RowNumber + ": Employee Number " + entry.EmployeeNumber
                        + " was not found on Employee Export, but #" + exportEntry.Key + " ("
                        + exportFullName + ") has a matching first name — check for a typo in the employee number.", true);
                    return;
                }
            }

            if (entry.Employee == null || entry.Employee.IsPartialEntry)
            {
                Log("Manual entry row " + entry.RowNumber + ": Employee Number "
                    + entry.EmployeeNumber + " was not found on Employee Export.", true);
            }
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

            string text = CellString(cell).Replace(",", "").Replace("$", "");
            return float.TryParse(text, out value);
        }
    }
}
