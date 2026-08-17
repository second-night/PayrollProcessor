using Excel = Microsoft.Office.Interop.Excel;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    /// <summary>
    /// Reads WfnEmployees.xlsx (ADP Workforce Now), the authoritative employee database.
    /// Employees may appear twice (Valley Bus LLC / MMF and Valley Bus Coaches / MKZ);
    /// numeric fields keep the greater value across duplicates.
    /// </summary>
    internal class WfnEmployeesReader
    {
        private const string FileName = "WfnEmployees.xlsx";

        /// <summary>
        /// Employee numbers that had a Years of Service value present in WfnEmployees.xlsx.
        /// Used so iSolved only fills this field when it was absent from WFN.
        /// </summary>
        internal static readonly HashSet<int> EmployeesWithYearsOfServiceFromWfn = new();

        /// <summary>
        /// Employee numbers that had a Vacation Balance value present in WfnEmployees.xlsx.
        /// </summary>
        internal static readonly HashSet<int> EmployeesWithVacationFromWfn = new();

        public void Read()
        {
            EmployeesWithYearsOfServiceFromWfn.Clear();
            EmployeesWithVacationFromWfn.Clear();

            string filePath = DesktopPath() + FileName;
            if (!File.Exists(filePath))
            {
                Log("ERROR: Please make sure there is an excel spreadsheet on your desktop named " + FileName, true);
                return;
            }

            Excel.Application excelApp = new();
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            try
            {
                foreach (Excel.Worksheet sheet in workBook.Worksheets)
                {
                    Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["B2"]].CurrentRegion;
                    object[,] cellData = (object[,])range.Value2;
                    int rows = cellData.GetLength(0);
                    int cols = cellData.GetLength(1);

                    List<string> headers = ReadHeaders(cellData, cols);
                    int ssnCol = FindColumn(headers, "Tax ID (SSN)");
                    int lastNameCol = FindColumn(headers, "Legal Last Name");
                    int firstNameCol = FindColumn(headers, "Legal First Name");
                    int birthDateCol = FindColumn(headers, "Birth Date");
                    int genderCol = FindColumn(headers, "Sex Code");
                    int mobilePhoneCol = FindColumn(headers, "Mobile Phone");
                    int homePhoneCol = FindColumn(headers, "Home Phone");
                    int hireDateCol = FindColumn(headers, "Hire Date");
                    int termDateCol = FindColumn(headers, "Termination Date");
                    int rehireDateCol = FindColumn(headers, "Rehire Date");
                    int salaryCol = FindColumn(headers, "Annual Salary");
                    int empNumberCol = FindColumn(headers, "File Number");
                    int locationCodeCol = FindColumn(headers, "Location Code");
                    int cityCol = FindColumn(headers, "Primary Address: City");
                    int yearsOfServiceCol = FindColumn(headers, "Years of Service");
                    int employmentCategoryCol = FindColumn(headers, "Worker Category Description");
                    int jobTitleCodeCol = FindColumn(headers, "Job Title Code");
                    int vacationCol = FindColumn(headers, "Vacation Balance");
                    int positionStatusCol = FindColumn(headers, "Position Status");
                    int fileNumberCol = FindColumn(headers, "File Number");
                    int primaryPositionCol = FindColumn(headers, "Primary Position");

                    Dictionary<Jobs, int> payColumns = new();
                    RegisterJobColumn(payColumns, Jobs.ADMIN, FindColumn(headers, "Rate - Admin"));
                    RegisterJobColumn(payColumns, Jobs.AIDE_SCHOOL, FindColumn(headers, "Rate - Para Charter"));
                    RegisterJobColumn(payColumns, Jobs.AIDE_CHARTER, FindColumn(headers, "Rate - Para School"));
                    RegisterJobColumn(payColumns, Jobs.BODY_SHOP, FindColumn(headers, "Rate - Body Shop"));
                    RegisterJobColumn(payColumns, Jobs.CLEANING, FindColumn(headers, "Rate - Cleaning"));
                    RegisterJobColumn(payColumns, Jobs.DRIVER_SCHOOL, FindColumn(headers, "Rate - Driver School"));
                    RegisterJobColumn(payColumns, Jobs.DRIVER_CHARTER, FindColumn(headers, "Rate - Driver Charter"));
                    RegisterJobColumn(payColumns, Jobs.MECHANIC, FindColumn(headers, "Rate - Mechanic"));
                    RegisterJobColumn(payColumns, Jobs.WASH_BAY, FindColumn(headers, "Rate - Wash Bay"));

                    if (empNumberCol == -1)
                    {
                        Log("WfnEmployees.xlsx: could not find 'File Number' column on sheet " + sheet.Name, true);
                        continue;
                    }

                    for (int rowNumber = 2; rowNumber <= rows; ++rowNumber)
                    {
                        if (!TryGetIntFromCell(cellData[rowNumber, empNumberCol + 1], out int employeeNumber))
                        {
                            continue;
                        }

                        string firstName = firstNameCol == -1 ? "" : CellString(cellData[rowNumber, firstNameCol + 1]);
                        string lastName = lastNameCol == -1 ? "" : CellString(cellData[rowNumber, lastNameCol + 1]);
                        if (!EmployeeDictionary.ContainsKey(employeeNumber))
                        {
                            string employeeName = (firstName + " " + lastName).Trim();
                            Employee newEmployee = new(employeeNumber, employeeName)
                            {
                                HireDate = DateTime.MinValue,
                                FirstName = firstName,
                                LastName = lastName
                            };
                            EmployeeDictionary.Add(employeeNumber, newEmployee);
                        }

                        Employee employee = EmployeeDictionary[employeeNumber];
                        employee.WasAlreadyInPayroll = true;
                        if (string.IsNullOrWhiteSpace(employee.FirstName) && !string.IsNullOrWhiteSpace(firstName))
                        {
                            employee.FirstName = firstName;
                        }
                        if (string.IsNullOrWhiteSpace(employee.LastName) && !string.IsNullOrWhiteSpace(lastName))
                        {
                            employee.LastName = lastName;
                        }

                        ApplyVacation(cellData, rowNumber, vacationCol, employee);
                        if (!IsPrimaryCompany(cellData, rowNumber, fileNumberCol, primaryPositionCol, positionStatusCol, employee))
                        {
                            continue;
                        }
                        ApplyStringIfPresent(cellData, rowNumber, ssnCol, ref employee.SocialSecurityNumber);
                        ApplyGender(cellData, rowNumber, genderCol, employee);
                        ApplyPhone(cellData, rowNumber, mobilePhoneCol, homePhoneCol, employee);
                        ApplyDateIfAbsent(cellData, rowNumber, birthDateCol, ref employee.BirthDate);
                        ApplyHireDate(cellData, rowNumber, hireDateCol, employee);
                        ApplyTermination(cellData, rowNumber, termDateCol, rehireDateCol, positionStatusCol, employee);
                        ApplyPayRates(cellData, rowNumber, payColumns, employee);
                        ApplySalary(cellData, rowNumber, salaryCol, employee);
                        ApplyLocation(cellData, rowNumber, locationCodeCol, cityCol, employee);
                        ApplyYearsOfService(cellData, rowNumber, yearsOfServiceCol, employee);
                        ApplyEmploymentCategory(cellData, rowNumber, employmentCategoryCol, employee);
                        ApplyJobTitleCode(cellData, rowNumber, jobTitleCodeCol, employee);
                    }
                }
            }
            finally
            {
                workBook.Close(false);
                excelApp.Quit();
            }
        }

        private static void ApplyStringIfPresent(object[,] cellData, int row, int col, ref string field)
        {
            if (col == -1 || !TryGetStringFromCell(cellData[row, col + 1], out string value))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(field))
            {
                field = value;
            }
        }

        private static void ApplyGender(object[,] cellData, int row, int genderCol, Employee employee)
        {
            if (genderCol == -1 || !TryGetStringFromCell(cellData[row, genderCol + 1], out string gender))
            {
                employee.IsMale = true;
                return;
            }

            employee.IsMale = gender != "F";
        }

        private static bool IsPrimaryCompany(object[,] cellData, int row, int positionCol, int primaryPositionCol, int positionStatusCol, Employee employee)
        {
            if (positionCol == -1 || !TryGetStringFromCell(cellData[row, positionCol + 1], out string positionId))
            {
                Log("Couldn't get position ID for employee " + employee.Name + " (" + employee.IdNumber + ") on row " + row, true);
                return false;
            }

            Company company = positionId.StartsWith("MMF") ? Company.VALLEY_BUS_LLC : Company.VALLEY_BUS_COACHES;
            TryGetStringFromCell(cellData[row, positionStatusCol + 1], out string positionStatus);
            if (positionStatus == "Active")
            {
                employee.ActiveCompanies.Add(company);
            }

            if (primaryPositionCol == -1 || !TryGetStringFromCell(cellData[row, primaryPositionCol + 1], out string bIsPrimaryPosition))
            {
                Log("Couldn't get primary position flag for employee " + employee.Name + " (" + employee.IdNumber + ") on row " + row, true);
                return false;
            }
            if (StringSearch(bIsPrimaryPosition, "yes"))
            {
                employee.PrimaryCompany = company;
                return true;
            }
            return false;
        }

        private static void ApplyPhone(object[,] cellData, int row, int mobileCol, int homeCol, Employee employee)
        {
            if (mobileCol != -1 && TryGetStringFromCell(cellData[row, mobileCol + 1], out string mobile))
            {
                employee.PhoneNumber = mobile;
                return;
            }

            if (homeCol != -1 && TryGetStringFromCell(cellData[row, homeCol + 1], out string home))
            {
                employee.PhoneNumber = home;
            }
        }

        private static void ApplyDateIfAbsent(object[,] cellData, int row, int col, ref DateTime field)
        {
            if (col == -1 || !TryGetDateFromCell(cellData[row, col + 1], out DateTime date))
            {
                return;
            }

            if (field == DateTime.MinValue)
            {
                field = date;
            }
        }

        private static void ApplyHireDate(object[,] cellData, int row, int hireDateCol, Employee employee)
        {
            if (hireDateCol == -1 || !TryGetDateFromCell(cellData[row, hireDateCol + 1], out DateTime hireDate))
            {
                return;
            }
            employee.HireDate = hireDate;
        }

        private static void ApplyTermination(object[,] cellData, int row, int termCol, int rehireCol, int positionStatusCol, Employee employee)
        {
            if (termCol == -1)
            {
                Log("Couldn't find Termination Date column for employee " + employee.Name + " (" + employee.IdNumber + ") on row " + row, true);
                return;
            }
            if (TryGetDateFromCell(cellData[row, termCol + 1], out DateTime termDate) && termDate != DateTime.MinValue)
            {
                employee.TerminationDate = termDate;
                DateTime rehireDate = DateTime.MinValue;
                if (rehireCol != -1)
                {
                    TryGetDateFromCell(cellData[row, rehireCol + 1], out rehireDate);
                }

                if (employee.TerminationDate.CompareTo(employee.HireDate) >= 0 && employee.TerminationDate.CompareTo(rehireDate) >= 0)
                {
                    employee.IsTerminated = true;
                }
                else
                {
                    Log("Rehire date for " + employee.Name + " is less than the hire or rehire date.", true);
                }
            }

            TryGetStringFromCell(cellData[row, positionStatusCol + 1], out string positionStatus);
            if (positionStatus == "Active" && employee.IsTerminated)
            {
                Log("Position Status for " + employee.Name + " is Active, but Termination Date is present.");
            }
            else if (positionStatus != "Active" && !employee.IsTerminated)
            {
                Log("Position Status for " + employee.Name + " is not Active, but Termination Date is absent.");
            }
        }

        private static void ApplyPayRates(object[,] cellData, int row, Dictionary<Jobs, int> payColumns, Employee employee)
        {
            foreach (KeyValuePair<Jobs, int> entry in payColumns)
            {
                if (entry.Value == -1)
                {
                    continue;
                }

                if (TryGetFloatFromCell(cellData[row, entry.Value + 1], out float payRate))
                {
                    employee.SetPayRate(entry.Key, payRate);
                    if (entry.Key == Jobs.DRIVER_CHARTER)
                    {
                        employee.SetPayRate(Jobs.DRIVER_CHARTER_PUBLIC, payRate);
                    }
                }
            }
        }

        private static void ApplySalary(object[,] cellData, int row, int salaryCol, Employee employee)
        {
            if (salaryCol == -1 || !TryGetFloatFromCell(cellData[row, salaryCol + 1], out float salary) || salary <= 50)
            {
                return;
            }

            employee.IsSalaried = true;
            employee.AnnualSalaryAmount = Math.Max(employee.AnnualSalaryAmount, salary);
        }

        private static void ApplyLocation(object[,] cellData, int row, int locationCol, int cityCol, Employee employee)
        {
            string locationCode = locationCol == -1 ? "" : CellString(cellData[row, locationCol + 1]);
            if (string.IsNullOrWhiteSpace(locationCode))
            {
                string city = cityCol == -1 ? "" : CellString(cellData[row, cityCol + 1]);
                if (StringSearch(city, "Grand Forks"))
                {
                    employee.IsAGrandForksEmployee = true;
                    Log("Assumed Location Code 'GF' for " + employee.Name + " (" + employee.IdNumber + ") because Location Code was blank and Primary Address: City is Grand Forks.");
                }
            }
            else if (StringSearch(locationCode, "GF") || StringSearch(locationCode, "Grand Forks"))
            {
                employee.IsAGrandForksEmployee = true;
            }
        }

        private static void ApplyYearsOfService(object[,] cellData, int row, int col, Employee employee)
        {
            if (col == -1 || !TryParseYearsOfService(cellData[row, col + 1], out int yearsOfService))
            {
                return;
            }

            EmployeesWithYearsOfServiceFromWfn.Add(employee.IdNumber);
            employee.YearsOfService = Math.Max(yearsOfService, employee.YearsOfService);
        }

        private static void ApplyEmploymentCategory(object[,] cellData, int row, int col, Employee employee)
        {
            string categoryDescription = "";
            if (col != -1)
            {
                if (TryGetStringFromCell(cellData[row, col + 1], out categoryDescription))
                { 
                    employee.EmploymentCategory = MapEmploymentCategory(categoryDescription); 
                }
            }
        }

        private static void ApplyJobTitleCode(object[,] cellData, int row, int col, Employee employee)
        {
            if (col == -1 || !TryGetStringFromCell(cellData[row, col + 1], out string jobTitleCode)
                || string.IsNullOrWhiteSpace(jobTitleCode))
            {
                return;
            }

            employee.JobTitleCode = jobTitleCode.Trim().ToUpperInvariant();
        }

        private static void ApplyVacation(object[,] cellData, int row, int col, Employee employee)
        {
            if (col == -1 || !TryGetFloatFromCell(cellData[row, col + 1], out float vacationHours))
            {
                return;
            }

            EmployeesWithVacationFromWfn.Add(employee.IdNumber);
            employee.VacationHours = Math.Max(vacationHours, employee.VacationHours);
        }

        internal static string MapEmploymentCategory(string categoryDescription)
        {
            if (string.IsNullOrWhiteSpace(categoryDescription))
            {
                return "PT";
            }

            if (StringSearch(categoryDescription, "Full Time") || categoryDescription.Equals("ACAFT", StringComparison.OrdinalIgnoreCase))
            {
                return "ACAFT";
            }

            if (StringSearch(categoryDescription, "Part Time") || categoryDescription.Equals("PT", StringComparison.OrdinalIgnoreCase))
            {
                return "PT";
            }

            return categoryDescription;
        }

        internal static bool TryParseYearsOfService(object? cellData, out int yearsOfService)
        {
            yearsOfService = 0;
            if (cellData == null)
            {
                return false;
            }

            if (cellData is double d)
            {
                yearsOfService = (int)d;
                return true;
            }

            string text = CellString(cellData);
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (int.TryParse(text, out yearsOfService))
            {
                return true;
            }

            // Examples: "27 year, 3 months", "0 year, 1 month"
            int yearIndex = text.IndexOf("year", StringComparison.OrdinalIgnoreCase);
            if (yearIndex <= 0)
            {
                return false;
            }

            string yearPart = text[..yearIndex].Trim().TrimEnd(',', ' ');
            return int.TryParse(yearPart, out yearsOfService);
        }

        private static void RegisterJobColumn(Dictionary<Jobs, int> columns, Jobs job, int columnIndex)
        {
            columns[job] = columnIndex;
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

        private static string CellString(object? cell)
        {
            return cell?.ToString()?.Trim() ?? "";
        }

        private static bool TryGetStringFromCell(object? cellData, out string outString)
        {
            outString = CellString(cellData);
            return outString != "";
        }

        private static bool TryGetDateFromCell(object? cellData, out DateTime date)
        {
            date = DateTime.MinValue;
            if (cellData == null)
            {
                return false;
            }

            if (cellData is DateTime dt)
            {
                date = dt;
                return true;
            }

            string? str = cellData.ToString();
            if (string.IsNullOrWhiteSpace(str))
            {
                return false;
            }

            if (double.TryParse(str, out double oaDate))
            {
                date = DateTime.FromOADate(oaDate);
                return true;
            }

            return DateTime.TryParse(str, out date);
        }

        private static bool TryGetFloatFromCell(object? cellData, out float outFloat)
        {
            outFloat = 0f;
            if (cellData == null)
            {
                return false;
            }

            if (cellData is double d)
            {
                outFloat = (float)d;
                return true;
            }

            return float.TryParse(CellString(cellData), out outFloat);
        }

        private static bool TryGetIntFromCell(object? cellData, out int outInt)
        {
            outInt = 0;
            if (cellData == null)
            {
                return false;
            }

            if (cellData is double d)
            {
                outInt = (int)d;
                return true;
            }

            string text = CellString(cellData).TrimStart('0');
            if (text == "")
            {
                text = "0";
            }
            return int.TryParse(text, out outInt);
        }
    }
}
