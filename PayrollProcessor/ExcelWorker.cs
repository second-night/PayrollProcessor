using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using static PayrollProcessor.ExcelWorker;
using static PayrollProcessor.Program;
using Excel = Microsoft.Office.Interop.Excel;
//using XmlExcel = OfficeOpenXml.Core.ExcelPackage;

namespace PayrollProcessor
{
    public class ExcelWorker
    {
        private const int GF_MAX_BUS = 399;
        private const int GF_MIN_BUS = 300;
        public static Dictionary<int, ImportedEmployee> ImportEmployees = new();
        public static Dictionary<int, (string FirstName, string LastName)> EmployeeExportByNumber = new();
        public DateTime FirstDayWeek2;
        public bool IsPrimaryPayrollRun { get; }
        public EmployeePayrollHistory PayrollHistory { get; } = new();
        HashSet<string> FieldsToInputEvenIfTheEmployeeWasAlreadyInPayroll = new() { "EmployeeNumber", "EmploymentCategory", "SSN" };

        public ExcelWorker()
        {
            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            bool manuallyEnterWeekTwoDate = PrintForm.InputDateTime(
                "Would you like to manually enter the first day of week 2 (as opposed to auto-detecting the date)?",
                "This is the primary payroll run (write payroll history records)",
                true,
                out DateTime dateTime,
                out bool isPrimaryPayrollRun);
            IsPrimaryPayrollRun = isPrimaryPayrollRun;
            if (manuallyEnterWeekTwoDate)
            {
                FirstDayWeek2 = dateTime;
            }
            else
            {
                if (DateTime.Now.Date.DayOfWeek == DayOfWeek.Tuesday)
                {
                    FirstDayWeek2 = today.AddDays(-9);
                }
                else if (DateTime.Now.Date.DayOfWeek == DayOfWeek.Wednesday)
                {
                    FirstDayWeek2 = today.AddDays(-10);
                }
                else if (DateTime.Now.Date.DayOfWeek == DayOfWeek.Thursday)
                {
                    FirstDayWeek2 = today.AddDays(-11);
                }
                else
                {
                    if (PrintForm.InputDateTime("Auto-detection failed. Please input the first day for week 2.", out dateTime))
                    {
                        FirstDayWeek2 = dateTime;
                    }
                    else
                    {
                        Log("Error: FirstDayWeek2 Failure.", true);
                    }
                }
            }

            if (FirstDayWeek2.DayOfWeek != DayOfWeek.Sunday)
            {
                Log("ERROR: FirstDayWeek2 must be on a Sunday. Exiting program.", true);
                Exit();
            }

            //manual override
            //FirstDayWeek2 = new DateTime(2023, 9, 24);
            //Log("FirstDayWeek2 override is active.", true);
        }

        public void ReadIsolvedEmployees()
        {
            return; //deprecated file
            if (!CheckForExcelFileOnDesktop("iSolvedEmployees.xlsx", out string filePath))
            {
                return;
            }
            Excel.Application excelApp = new Excel.Application();
            var fInfo = new FileInfo(filePath);
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            Dictionary<Jobs, int> payColumns = new();
            const int SSN_COLUMN = 2; //Tax ID (SSN)
            const int EMP_LAST_NAME_COLUMN = 3; //Legal Last Name
            const int EMP_FIRST_NAME_COLUMN = 4; //Legal First Name
            const int BIRTH_DATE_COLUMN = 7; //Birth Date
            const int GENDER_COLUMN = 8; //Sex Code
            const int PHONE_NUMBER_COLUMN = 15; //"Home Phone" or "Mobile Phone" - use mobile phone first and if nothing under mobile phone then use home phone.
            const int HIRE_DATE_COLUMN = 18; //Hire Date
            const int TERM_DATE_COLUMN = 19; //Termination Date
            const int REHIRE_DATE_COLUMN = 20; //Rehire Date
            const int SALARY_COLUMN = 28; //Annual Salary
            int ADMIN_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.ADMIN, 30); //Rate - Admin
            int AIDE_SCHOOL_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.AIDE_SCHOOL, 31); //Rate - Para Charter
            int AIDE_CHARTER_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.AIDE_CHARTER, 32); //Rate - Para School
            int BODY_SHOP_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.BODY_SHOP, 34); //Rate - Body Shop
            int CLEANING_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.CLEANING, 35); //Rate - Cleaning
            int DRIVER_SCHOOL_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.DRIVER_SCHOOL, 36); //Rate - Driver School
            int DRIVER_CHARTER_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.DRIVER_CHARTER, 37); //Rate - Driver Charter
            int MECHANIC_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.MECHANIC, 38); //Rate - Mechanic
            int WASH_BAY_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.WASH_BAY, 40); //Rate - Wash Bay
            const int EMP_NUMBER_COLUMN = 41; //"File Number" - (it is in string format, I'm not sure if that matters. It has extra '0's in front of the number)
            int TRAINING_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.TRAINING, 42);
            const int ORGANIZATION_TAG_COLUMN = 44; //"Location Code" - Entries will be 'FAR', 'GF', or blank, if they are blank, the 'FAR' is assumed unless the employees entry under the header "Primary Address: City" is Grand Forks, which would the mean an entry of 'GF' is assumed (throw a log if GF is assumed).
            const int YEARS_OF_SERVICE_COLUMN = 45; //"Years of Service" - this is in string format, example1: '27 year, 3 months', example2: '0 year, 1 month'. We only need to parse the year into integer form, months/decimals are not wanted.
            const int EMPLOYMENT_CATEGORY_COLUMN = 47; //"Worker Category Description" - entries will be 'Full Time' or 'Part Time' or blank, if they are blank the 'Part Time' is assumed.
            const int DD_ACCOUNT_1 = 48; //deprecated, don't worry about it.
            const int VACATION_HOURS_COLUMN = 55; //Vacation Balance

            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["B2"]].CurrentRegion;
                //Excel.Range excelRange = (Excel.Range)sheet.Range[sheet.Range["A1"], sheet.Range["P36"]];
                var cellData = (Object[,])range.Value2;
                int rows = cellData.GetLength(0);
                for (int rowNumber = 1; rowNumber <= rows; ++rowNumber)
                {
                    //Log("cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString() == " + cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString());
                    if (TryGetIntFromCell(cellData[rowNumber, EMP_NUMBER_COLUMN], out int employeeNumber))
                    {
                        bool existedFromWfn = Program.EmployeeDictionary.ContainsKey(employeeNumber);
                        if (!existedFromWfn)
                        {
                            string? employeeName = cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString() + " " + cellData[rowNumber, EMP_LAST_NAME_COLUMN].ToString();
                            Program.EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, employeeName)
                            {
                                HireDate = DateTime.MinValue
                            });
                            Log("Employee " + employeeName + " (" + employeeNumber + ") found only in iSolvedEmployees.xlsx (not in WfnEmployees.xlsx); importing from iSolved.");
                        }
                        Employee employee = Program.EmployeeDictionary[employeeNumber];
                        employee.WasAlreadyInPayroll = true;

                        if ((!existedFromWfn || string.IsNullOrWhiteSpace(employee.SocialSecurityNumber))
                            && TryGetStringFromCell(cellData[rowNumber, SSN_COLUMN], out string ssn))
                        {
                            if (existedFromWfn)
                            {
                                Log("Filled SocialSecurityNumber for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                            }
                            employee.SocialSecurityNumber = ssn;
                        }

                        if (!existedFromWfn)
                        {
                            employee.IsMale = !(TryGetStringFromCell(cellData[rowNumber, GENDER_COLUMN], out string gender) && gender == "F");
                        }

                        if ((!existedFromWfn || string.IsNullOrWhiteSpace(employee.PhoneNumber))
                            && TryGetStringFromCell(cellData[rowNumber, PHONE_NUMBER_COLUMN], out string phone))
                        {
                            if (existedFromWfn)
                            {
                                Log("Filled PhoneNumber for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                            }
                            employee.PhoneNumber = phone;
                        }

                        if ((!existedFromWfn || employee.HireDate == DateTime.MinValue)
                            && TryGetDateFromCell(cellData[rowNumber, HIRE_DATE_COLUMN], out DateTime hireDate))
                        {
                            if (existedFromWfn)
                            {
                                Log("Filled HireDate for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                            }
                            employee.HireDate = hireDate;
                        }

                        if ((!existedFromWfn || employee.BirthDate == DateTime.MinValue)
                            && TryGetDateFromCell(cellData[rowNumber, BIRTH_DATE_COLUMN], out DateTime birthDate))
                        {
                            if (existedFromWfn)
                            {
                                Log("Filled BirthDate for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                            }
                            employee.BirthDate = birthDate;
                        }

                        if ((!existedFromWfn || employee.TerminationDate == DateTime.MinValue)
                            && TryGetDateFromCell(cellData[rowNumber, TERM_DATE_COLUMN], out DateTime termDate))
                        {
                            if (existedFromWfn)
                            {
                                Log("Filled TerminationDate for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                            }
                            employee.TerminationDate = termDate;
                            TryGetDateFromCell(cellData[rowNumber, REHIRE_DATE_COLUMN], out DateTime rehireDate);
                            if (employee.TerminationDate.CompareTo(employee.HireDate) >= 0 && employee.TerminationDate.CompareTo(rehireDate) >= 0)
                            {
                                employee.IsTerminated = true;
                            }
                            else
                            {
                                Log("Rehire date for " + employee.Name + " is less than the hire or rehire date.", true);
                            }
                        }

                        foreach (KeyValuePair<Jobs, int> entry in payColumns)
                        {
                            if (TryGetFloatFromCell(cellData[rowNumber, entry.Value], out float payRate) && payRate > 0f)
                            {
                                float existingRate = employee.PayRates.GetValueOrDefault(entry.Key, 0f);
                                if (existedFromWfn && existingRate <= 0f)
                                {
                                    Log("Filled pay rate for " + entry.Key + " for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                                    employee.SetPayRate(entry.Key, payRate);
                                    if (entry.Key == Jobs.DRIVER_CHARTER)
                                    {
                                        employee.SetPayRate(Jobs.DRIVER_CHARTER_PUBLIC, payRate);
                                    }
                                }
                                else if (!existedFromWfn)
                                {
                                    employee.SetPayRate(entry.Key, Math.Max(payRate, existingRate));
                                }
                            }
                        }

                        if (TryGetFloatFromCell(cellData[rowNumber, SALARY_COLUMN], out float salary) && salary > 50)
                        {
                            if (!existedFromWfn)
                            {
                                employee.IsSalaried = true;
                                employee.AnnualSalaryAmount = Math.Max(employee.AnnualSalaryAmount, salary);
                                employee.PrimaryCompany = employee.IdNumber == 1734 || employee.IdNumber == 123 ? Company.VALLEY_BUS_COACHES : Company.VALLEY_BUS_LLC;
                            }
                            else if (!employee.IsSalaried || employee.AnnualSalaryAmount <= 0f)
                            {
                                Log("Filled AnnualSalaryAmount for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                                employee.IsSalaried = true;
                                employee.AnnualSalaryAmount = salary;
                                employee.PrimaryCompany = employee.IdNumber == 1734 || employee.IdNumber == 123 ? Company.VALLEY_BUS_COACHES : Company.VALLEY_BUS_LLC;
                            }
                        }

                        if (!employee.IsAGrandForksEmployee)
                        {
                            if (TryGetStringFromCell(cellData[rowNumber, ORGANIZATION_TAG_COLUMN], out string tag)
                                && (StringSearch(tag, "Grand Forks") || StringSearch(tag, "GF")))
                            {
                                if (existedFromWfn)
                                {
                                    Log("Filled IsAGrandForksEmployee for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                                }
                                employee.IsAGrandForksEmployee = true;
                            }
                        }

                        if (!employee.IsAMechanicApprentice)
                        {
                            if (TryGetStringFromCell(cellData[rowNumber, ORGANIZATION_TAG_COLUMN], out string orgTag) && StringSearch(orgTag, "Apprentice"))
                            {
                                if (existedFromWfn)
                                {
                                    Log("Filled IsAMechanicApprentice for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                                }
                                employee.IsAMechanicApprentice = true;
                            }
                        }

                        if (!WfnEmployeesReader.EmployeesWithYearsOfServiceFromWfn.Contains(employeeNumber))
                        {
                            int yearsOfService = 0;
                            bool gotYears = WfnEmployeesReader.TryParseYearsOfService(cellData[rowNumber, YEARS_OF_SERVICE_COLUMN], out yearsOfService)
                                || TryGetIntFromCell(cellData[rowNumber, YEARS_OF_SERVICE_COLUMN], out yearsOfService);
                            if (gotYears)
                            {
                                if (existedFromWfn)
                                {
                                    Log("Filled YearsOfService for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                                }
                                employee.YearsOfService = Math.Max(yearsOfService, employee.YearsOfService);
                            }
                        }

                        if (string.IsNullOrWhiteSpace(employee.EmploymentCategory)
                            && TryGetStringFromCell(cellData[rowNumber, EMPLOYMENT_CATEGORY_COLUMN], out string category))
                        {
                            if (existedFromWfn)
                            {
                                Log("Filled EmploymentCategory for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                            }
                            employee.EmploymentCategory = WfnEmployeesReader.MapEmploymentCategory(category);
                        }

                        //if (!employee.HasAnActiveDirectDepositAccount)
                        //{
                        //    for (int i = 0; i < 6; i++)
                        //    {
                        //        if (TryGetStringFromCell(cellData[rowNumber, DD_ACCOUNT_1 + i], out string accountStatus))
                        //        {
                        //            if (existedFromWfn && !employee.HasAnyDirectDepositAccount)
                        //            {
                        //                Log("Filled direct deposit status for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                        //            }
                        //            employee.HasAnyDirectDepositAccount = true;
                        //            if ((i == 5 && accountStatus != "") || accountStatus == "Active")
                        //            {
                        //                employee.HasAnActiveDirectDepositAccount = true;
                        //                break;
                        //            }
                        //        }
                        //    }
                        //}

                        //if (!WfnEmployeesReader.EmployeesWithVacationFromWfn.Contains(employeeNumber)
                        //    && TryGetFloatFromCell(cellData[rowNumber, VACATION_HOURS_COLUMN], out float vacationHours))
                        //{
                        //    if (existedFromWfn)
                        //    {
                        //        Log("Filled VacationHours for " + employee.Name + " (" + employee.IdNumber + ") from iSolvedEmployees.xlsx.");
                        //    }
                        //    employee.VacationHours = Math.Max(vacationHours, employee.VacationHours);
                        //}
                    }
                }
            }
            workBook.Close();
            excelApp.Quit();

            //Marshal.ReleaseComObject(workBook);
            //Marshal.ReleaseComObject(excelApp);
        }

        public void PreCheckTimeSheets()
        {
            const int EMP_NUMBER_COLUMN = 2;
            const int EMP_NAME_COLUMN = 3;
            const int DAY_COLUMN = 6;
            const int PUNCH_IN_COLUMN = 7;
            const int PUNCH_OUT_COLUMN = 10;
            const int ROUNDED_TIME_COLUMN = 14;
            const int JOB_TYPE_COLUMN = 15;
            const int NOTES_COLUMN = 18;
            const int BUS_NUMBER_COLUMN = 34;

            if (!CheckForExcelFileOnDesktop("Timesheets.xlsx", out string filePath))
            {
                return;
            }
            var lastModified = File.GetLastWriteTime(filePath);
            Excel.Application excelApp = new Excel.Application();
            var fInfo = new FileInfo(filePath);
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Excel.Range range = sheet.Range[sheet.Range["A6"], sheet.Range["B8"]];
                range = range.CurrentRegion;
                int rows = range.Value2.GetLength(0) + 6;
                range = sheet.Range[sheet.Range["A1"], sheet.Range["BZ" + rows]];
                var cellData = (Object[,])range.Value2;
                rows = cellData.GetLength(0);
                for (int rowNumber = 6; rowNumber <= rows; ++rowNumber)
                {
                    if (null != cellData[rowNumber, DAY_COLUMN])
                    {
                        if (!TryGetDateFromCell(cellData[rowNumber, DAY_COLUMN], out DateTime date))
                        {
                            Log("date == nothing for row: " + rowNumber);
                            continue;
                        }
                        if (!TryGetFloatFromCell(cellData[rowNumber, ROUNDED_TIME_COLUMN], out float time))
                        {
                            continue;
                        }
                        if (time < 0.1f)
                        {
                            continue;
                        }

                        if (!TryGetIntFromCell(cellData[rowNumber, EMP_NUMBER_COLUMN], out int employeeNumber))
                        {
                            Log("Couldn't get employee number", true);
                            continue;
                        }

                        if (!EmployeeDictionary.ContainsKey(employeeNumber))
                        {
                            string name = null == cellData[rowNumber, EMP_NAME_COLUMN] ? "" : (null == cellData[rowNumber, EMP_NAME_COLUMN].ToString() ? "" : new string((cellData[rowNumber, EMP_NAME_COLUMN].ToString())));
                            Employee emp = new(employeeNumber, name)
                            {
                                IsPartialEntry = true
                            };
                            EmployeeDictionary.Add(employeeNumber, emp);
                        }
                        EmployeeDictionary[employeeNumber].HadHoursInTimesheets = true;


                        Shift temporaryShift = new(Company.VALLEY_BUS_LLC); //temporary, doesn't get added to employee shifts.
                        temporaryShift.ShiftTime = time;
                        temporaryShift.Date = date;
                        Employee employee = EmployeeDictionary[employeeNumber];

                        if (TryGetIntFromCell(cellData[rowNumber, JOB_TYPE_COLUMN], out temporaryShift.JobInt))
                        {
                            temporaryShift.JobType = GetJobTypeFromCode(temporaryShift.JobInt);
                            if (temporaryShift.JobType == Jobs.DRIVER_SCHOOL && !employee.PayRates.ContainsKey(temporaryShift.JobType))
                            {
                                temporaryShift.JobType = Jobs.NON_CDL_DRIVER;
                            }
                        }

                        if (temporaryShift.IsASchoolRouteShift())
                        {
                            if (null != cellData[rowNumber, BUS_NUMBER_COLUMN] && StringSearch(cellData[rowNumber, BUS_NUMBER_COLUMN].ToString(), "old"))
                            {
                                string? busNumberString = cellData[rowNumber, BUS_NUMBER_COLUMN].ToString();
                                if (null != busNumberString)
                                {
                                    busNumberString = busNumberString.Replace("old", "");
                                    busNumberString = busNumberString.Replace("Old", "");
                                    busNumberString = busNumberString.Replace(" ", "");
                                    if (int.TryParse(busNumberString, out temporaryShift.BusNumber))
                                    {
                                    }
                                }
                            }
                            else
                            {
                                TryGetIntFromCell(cellData[rowNumber, BUS_NUMBER_COLUMN], out temporaryShift.BusNumber);
                            }
                        }

                        if (temporaryShift.BusNumber != 0)
                        {
                            //if (!employee.RoutesByBusNumber.ContainsKey(temporaryShift.BusNumber))
                            //{
                            //    employee.RoutesByBusNumber.Add(temporaryShift.BusNumber, new());
                            //}
                            //if (!employee.RoutesByBusNumber[temporaryShift.BusNumber].ContainsKey(temporaryShift.TimeContext()))
                            //{
                            //    employee.RoutesByBusNumber[temporaryShift.BusNumber].Add(temporaryShift.TimeContext(), new());
                            //}
                            //employee.RoutesByBusNumber[temporaryShift.BusNumber][temporaryShift.TimeContext()] += 1;

                            if (!employee.BusShiftTotals.ContainsKey(date.DayOfWeek))
                            {
                                employee.BusShiftTotals.Add(date.DayOfWeek, new());
                                employee.ShiftsByBusNumber.Add(date.DayOfWeek, new());
                            }
                            if (!employee.BusShiftTotals[date.DayOfWeek].ContainsKey(temporaryShift.TimeContext()))
                            {
                                employee.BusShiftTotals[date.DayOfWeek].Add(temporaryShift.TimeContext(), new());
                                employee.ShiftsByBusNumber[date.DayOfWeek].Add(temporaryShift.TimeContext(), new());
                            }
                            if (!employee.ShiftsByBusNumber[date.DayOfWeek][temporaryShift.TimeContext()].ContainsKey(temporaryShift.BusNumber))
                            {
                                employee.ShiftsByBusNumber[date.DayOfWeek][temporaryShift.TimeContext()].Add(temporaryShift.BusNumber, new());
                            }
                            employee.BusShiftTotals[date.DayOfWeek][temporaryShift.TimeContext()] += 1;
                            employee.ShiftsByBusNumber[date.DayOfWeek][temporaryShift.TimeContext()][temporaryShift.BusNumber] += 1;
                        }

                        if (TryGetStringFromCell(cellData[rowNumber, NOTES_COLUMN], out string notes))
                        {
                            if (notes == "bonus" || notes == "Bonus" || StringSearch(notes, "bonus") || StringSearch(notes, "starting"))
                            {
                                Program.BusStartingDays.Add(date.Day);
                            }
                        }
                    }
                }
            }
            workBook.Close();
            excelApp.Quit();

            //Marshal.ReleaseComObject(workBook);
            //Marshal.ReleaseComObject(excelApp);
        }

        public void ReadEmployeeExport()
        {
            EmployeeExportByNumber.Clear();
            if (!CheckForExcelFileOnDesktop("Employee Export.xlsx", out string filePath))
            {
                Log("Couldn't find employee export.", true);
                return;
            }
            var lastModified = System.IO.File.GetLastWriteTime(filePath);
            if (new DateTime(lastModified.Year, lastModified.Month, lastModified.Day).CompareTo(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)) < 0)
            {
                Log("Employee Export is old.", true);
            }
            Excel.Application excelApp = new Excel.Application();
            var fInfo = new FileInfo(filePath);
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            const int EMPLOYEE_NUMBER = 3;
            const int EMP_FIRST_NAME_COLUMN = 5;
            const int EMP_LAST_NAME_COLUMN = 7;
            const int EMPLOYEE_GROUPS = 44;

            List<string> headers = new() { "Date Received (Direct Deposit Authorization )", "Start Date", "Employee #", "Employee #", "SSN", "First Name", "Middle Name", "Last Name", "Email", "Street", "Apt/Suite/Unit", "Zip", "City", "State", "Birthdate", "Phone", "Date Received (Form I-9)", "Citizenship Designation (Form I-9)", "Gender", "Position", "Zip", "Filing Status (W4)", "Deductions (W4)", "Total Dependents Withholding (W4)", "Extra Withholding (W4)", "Exempt Status (W4)", "Account 1", "Account 1 - $ Specific Deposit Amount", "Account 1 - % Net Amount", "Account 1 - Account Number", "Account 1 - Deposit Instructions", "Account 1 - Routing Number", "Account 1 - Type", "Account 2", "Account 2 - $ Specific Deposit Amount", "Account 2 - % Net Amount", "Account 2 - Account Number", "Account 2 - Deposit Instructions", "Account 2 - Routing Number", "Account 2 - Type", "Account 3", "Account 3 - Account Number", "Account 3 - Routing Number", "Account 3 - Type", "Employee Groups", "VALLEY BUS LOCATION" };

            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["B2"]].CurrentRegion;
                //Excel.Range excelRange = (Excel.Range)sheet.Range[sheet.Range["A1"], sheet.Range["P36"]];
                var cellData = (Object[,])range.Value2;
                int rows = cellData.GetLength(0);
                bool bOverride = false;
                for (int rowNumber = 2; rowNumber <= rows; ++rowNumber)
                {
                    //Log("cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString() == " + cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString());
                    if (!TryGetIntFromCell(cellData[rowNumber, EMPLOYEE_NUMBER], out int employeeNumber))
                    {
                        if (null == cellData[rowNumber, EMP_FIRST_NAME_COLUMN] || null == cellData[rowNumber, EMP_LAST_NAME_COLUMN].ToString())
                        {
                            continue;
                        }
                        foreach (var employeeEntry in EmployeeDictionary)
                        {
                            if (StringSearch(employeeEntry.Value.Name, cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString()) && StringSearch(employeeEntry.Value.Name, cellData[rowNumber, EMP_LAST_NAME_COLUMN].ToString()))
                            {
                                employeeNumber = employeeEntry.Key;
                                break;
                            }
                        }
                    }
                    if (employeeNumber == 0)
                    {
                        continue;
                    }
                    string exportFirstName = cellData[rowNumber, EMP_FIRST_NAME_COLUMN]?.ToString()?.Trim() ?? "";
                    string exportLastName = cellData[rowNumber, EMP_LAST_NAME_COLUMN]?.ToString()?.Trim() ?? "";
                    EmployeeExportByNumber[employeeNumber] = (exportFirstName, exportLastName);
                    ImportedEmployee importedEmployee = new()
                    {
                        WasOnImployeeExportSheet = true
                    };
                    if (ImportEmployees.ContainsKey(employeeNumber))
                    {
                        Log("Fatal Error:\nDuplicate employee number " + employeeNumber.ToString() + " found.");
                        Exit();
                    }
                    ImportEmployees.Add(employeeNumber, importedEmployee);
                    importedEmployee.ImportFields.Add("TimeClockID", employeeNumber.ToString());
                    importedEmployee.ImportFields.Add("EmployeeNumber", employeeNumber.ToString());
                    importedEmployee.ImportFields.Add("WorkLocation", "Fargo");
                    importedEmployee.ImportFields.Add("PayType", "Hourly");
                    importedEmployee.ImportFields.Add("Frequency", "26");
                    if (!Program.EmployeeDictionary.ContainsKey(employeeNumber))
                    {
                        string? employeeName = cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString() + " " + cellData[rowNumber, EMP_LAST_NAME_COLUMN].ToString();
                        Program.EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, employeeName));
                    }
                    Employee employee = Program.EmployeeDictionary[employeeNumber];
                    employee.IsPartialEntry = false;

                    string[] socialSecurityNumberEntries = new string[2] { "", "" };
                    string[] birthDateEntries = new string[2] { "", "" };


                    bool isDriver = false;
                    bool isPara = false;

                    foreach (var header in headers)
                    {
                        Object? cell = null;
                        for (int columnNumber = 0; columnNumber < headers.Count; columnNumber++)
                        {
                            if (null != cellData[1, columnNumber + 1] && header == cellData[1, columnNumber + 1].ToString())
                            {
                                cell = cellData[rowNumber, columnNumber + 1];
                                break;
                            }
                        }
                        if (cell != null)
                        {
                            if (!TryGetStringFromCell(cell, out string cellString))
                            {
                                continue;
                            }
                            //switch (header)
                            //{
                            //    case "Date Received (Direct Deposit Authorization )":

                            //        if (TryGetDateFromCell(cell, out employee.DateOfDirectDepositUpdateInWorkBright))
                            //        {
                            //            if (employee.WasAlreadyInPayroll && !FieldsToInputEvenIfTheEmployeeWasAlreadyInPayroll.Contains(header))
                            //            {
                            //                if (employee.DateOfDirectDepositUpdateInWorkBright.AddDays(14).CompareTo(DateTime.Today) > 0)
                            //                {
                            //                    if (PrintForm.InputBool("Should direct deposit should be imported for " + employee.Name + " ? "))
                            //                    {
                            //                        bOverride = true;
                            //                        employee.needsDDImported = true;
                            //                    }
                            //                }
                            //            }
                            //        }
                            //        break;
                            //}
                            if (!bOverride && employee.WasAlreadyInPayroll && !FieldsToInputEvenIfTheEmployeeWasAlreadyInPayroll.Contains(header))
                            {
                                continue;
                                //employee.TerminationDate.AddMonths();
                                //if (!employee.IsTerminated)
                                //{
                                //    continue;
                                //}
                                //else
                                //{
                                //    Log("Re-importing " + employee.Name + " because they are terminated in payroll.");
                                //}
                            }
                            switch (header)
                            {
                                case "Start Date":
                                    if (TryGetDateFromCell(cell, out DateTime hireDate))
                                    {
                                        importedEmployee.ImportFields["HireDate"] = hireDate.ToShortDateString();
                                    }
                                    break;
                                case "Employee #":
                                    break;
                                case "SSN":
                                    socialSecurityNumberEntries[0] = cellString;
                                    break;
                                case "Full Social Security Number":
                                    socialSecurityNumberEntries[1] = cellString;
                                    break;
                                case "First Name":
                                    importedEmployee.ImportFields["FirstName"] = cellString;
                                    if (string.IsNullOrWhiteSpace(employee.FirstName) && !string.IsNullOrWhiteSpace(cellString))
                                    {
                                        employee.FirstName = cellString;
                                    }
                                    break;
                                case "Middle Name":
                                    importedEmployee.ImportFields["MiddleName"] = cellString;
                                    if (!string.IsNullOrWhiteSpace(cellString))
                                    {
                                        employee.MiddleName = cellString;
                                    }
                                    break;
                                case "Last Name":
                                    importedEmployee.ImportFields["LastName"] = cellString;
                                    if (string.IsNullOrWhiteSpace(employee.LastName) && !string.IsNullOrWhiteSpace(cellString))
                                    {
                                        employee.LastName = cellString;
                                    }
                                    break;
                                case "Email":
                                    importedEmployee.ImportFields["SelfServiceEnabled"] = "Y";
                                    importedEmployee.ImportFields["SelfServiceEmail"] = cellString;
                                    break;
                                case "Street":
                                    importedEmployee.ImportFields["Address1"] = cellString;
                                    break;
                                case "Apt/Suite/Unit":
                                    importedEmployee.ImportFields["Address2"] = cellString;
                                    break;
                                case "Zip":
                                    if (!importedEmployee.ImportFields.ContainsKey("ZipCode"))
                                    {
                                        importedEmployee.ImportFields["ZipCode"] = cellString;
                                        if (cellString == "58102")
                                        {
                                            cellString += "[ND0170050]";
                                        }
                                        importedEmployee.ImportFields["ResidentLocation"] = cellString;
                                    }
                                    break;
                                case "City":
                                    importedEmployee.ImportFields["City"] = cellString;
                                    break;
                                case "State":
                                    importedEmployee.ImportFields["State"] = cellString;
                                    break;
                                case "Birthdate":
                                    birthDateEntries[0] = cellString;
                                    break;
                                case "Birth Date":
                                    birthDateEntries[1] = cellString;
                                    break;
                                case "Phone":
                                    importedEmployee.ImportFields["HomePhone"] = cellString;
                                    employee.PhoneNumber = cellString;
                                    break;
                                case "Date Received (Form I-9)":
                                    importedEmployee.ImportFields["I9Completed"] = cellString == "" ? "N" : "Y";
                                    if (TryGetDateFromCell(cell, out DateTime iNineDate))
                                    {
                                        importedEmployee.ImportFields["I9CompletedDate"] = iNineDate.ToShortDateString();
                                    }
                                    break;
                                case "Citizenship Designation (Form I-9)":
                                    if (StringSearch(cellString, "citizen"))
                                    {
                                        importedEmployee.ImportFields["Citizenship"] = "1";
                                    }
                                    else if (StringSearch(cellString, "national"))
                                    {
                                        importedEmployee.ImportFields["Citizenship"] = "5";
                                    }
                                    else if (StringSearch(cellString, "permanent"))
                                    {
                                        importedEmployee.ImportFields["Citizenship"] = "3";
                                    }
                                    else if (StringSearch(cellString, "alien"))
                                    {
                                        importedEmployee.ImportFields["Citizenship"] = "2";
                                    }
                                    else
                                    {
                                        Log("ERROR: Couldn't find citizenship for " + cellString + " (" + employee.Name);
                                    }
                                    break;
                                case "Gender":
                                    employee.IsMale = !StringSearch(cellString, "Female");
                                    importedEmployee.ImportFields["Gender"] = employee.IsMale ? "M" : "F";
                                    break;
                                case "Position":
                                    bool fT = false;
                                    if (StringSearch(cellString, "mechanic"))
                                    {
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.MECHANIC).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.MECHANIC, false);
                                        fT = true;
                                    }
                                    else if (StringSearch(cellString, "wash bay"))
                                    {
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.WASH_BAY).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.WASH_BAY, false);
                                        fT = true;
                                    }
                                    else if (StringSearch(cellString, "para"))
                                    {
                                        isPara = true;
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.AIDE_SCHOOL).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.AIDE_SCHOOL, false);
                                    }
                                    else if (StringSearch(cellString, "driver"))
                                    {
                                        isDriver = true;
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.DRIVER_SCHOOL).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.DRIVER_SCHOOL, false);
                                    }
                                    else if (StringSearch(cellData[rowNumber, EMPLOYEE_GROUPS].ToString(), "driver"))
                                    {
                                        isDriver = true;
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.DRIVER_SCHOOL).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.DRIVER_SCHOOL, false);
                                    }
                                    else if (StringSearch(cellData[rowNumber, EMPLOYEE_GROUPS].ToString(), "para") || StringSearch(cellData[rowNumber, EMPLOYEE_GROUPS].ToString(), "aide"))
                                    {
                                        isPara = true;
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.AIDE_SCHOOL).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.AIDE_SCHOOL, false);
                                    }
                                    else
                                    {
                                        Log("Giving para as job to emp: " + employee.Name + " for position: " + cellString);
                                        importedEmployee.ImportFields["Job"] = ((int)Jobs.AIDE_SCHOOL).ToString();
                                        importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.AIDE_SCHOOL, false);
                                    }
                                    importedEmployee.ImportFields["EmploymentCategory"] = fT ? "ACAFT" : "PT";
                                    employee.EmploymentCategory = fT ? "ACAFT" : "PT";
                                    break;
                                case "Filing Status (W4)":
                                    if (StringSearch(cellString, "single"))
                                    {
                                        importedEmployee.ImportFields["FedFilingStatus"] = "FDS2";
                                        importedEmployee.ImportFields["StateFilingStatus"] = "NDS";
                                    }
                                    else if (StringSearch(cellString, "Household"))
                                    {
                                        importedEmployee.ImportFields["FedFilingStatus"] = "FDH";
                                        importedEmployee.ImportFields["StateFilingStatus"] = "NDH";
                                    }
                                    else
                                    {
                                        importedEmployee.ImportFields["FedFilingStatus"] = "FDM2";
                                        importedEmployee.ImportFields["StateFilingStatus"] = "NDM";
                                    }
                                    break;
                                case "Deductions (W4)":
                                    importedEmployee.ImportFields["FedDeductions"] = cellString;
                                    importedEmployee.ImportFields["StateExemptions"] = cellString;
                                    break;
                                case "Total Dependents Withholding (W4)":
                                    importedEmployee.ImportFields["FedDependentsAmt"] = cellString;
                                    break;
                                case "Extra Withholding (W4)":
                                    importedEmployee.ImportFields["FedAddlAmount"] = cellString;
                                    break;
                                case "Exempt Status (W4)":
                                    if (StringSearch(cellString, "EXEMPT"))
                                    {
                                        importedEmployee.ImportFields["FedBlockTax"] = "true";
                                        importedEmployee.ImportFields["StateBlockTax"] = "true";
                                    }
                                    break;
                                case "Account 1":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2";
                                case "Account 1 - $ Specific Deposit Amount":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2 - $ Specific Deposit Amount";
                                case "Account 1 - % Net Amount":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2 - % Net Amount";
                                case "Account 1 - Account Number":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2 - Account Number";
                                case "Account 1 - Deposit Instructions":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2 - Deposit Instructions";
                                case "Account 1 - Routing Number":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2 - Routing Number";
                                case "Account 1 - Type":
                                    importedEmployee.LatestAccountIndex = 0;
                                    goto case "Account 2 - Type";
                                case "Account 2":
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 2 - $ Specific Deposit Amount":
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Amount"] = cellString;
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 2 - % Net Amount":
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Percent"] = cellString;
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 2 - Account Number":
                                    employee.HasAnActiveDirectDepositAccount = true;
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Key"] = employeeNumber.ToString();
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["AccountNumber"] = cellString;
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Status"] = "A";
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 2 - Deposit Instructions":
                                    if (StringSearch(cellString, "Entire Net"))
                                    {
                                        importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Sequence"] = "0";
                                        importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Amount"] = "";
                                        importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Percent"] = "";
                                    }
                                    else
                                    {
                                        importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["Sequence"] = "1";
                                    }
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 2 - Routing Number":
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["RoutingNumber"] = cellString;
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 2 - Type":
                                    importedEmployee.DDAccounts[importedEmployee.LatestAccountIndex]["AccountType"] = StringSearch(cellString, "saving") ? "S" : "C";
                                    importedEmployee.LatestAccountIndex = 1;
                                    break;
                                case "Account 3":
                                    break;
                                case "Account 3 - Account Number":
                                    if (null != cellString && "" != cellString)
                                    {
                                        Log("3 Accounts found", true);
                                    }
                                    break;
                                case "Account 3 - Routing Number":
                                    break;
                                case "Account 3 - Type":
                                    break;
                                case "Employee Groups":
                                    if (StringSearch(cellString, "GRAND FORKS"))
                                    {
                                        employee.IsAGrandForksEmployee = true;
                                        importedEmployee.ImportFields["OrganizationValue2"] = "GF";
                                    }
                                    break;
                                case "VALLEY BUS LOCATION":
                                    importedEmployee.ImportFields["Rate_Training"] = TRAINING_RATE.ToString();
                                    employee.PayRates[Jobs.TRAINING] = TRAINING_RATE;
                                    if (StringSearch(cellString, "GF"))
                                    {
                                        employee.IsAGrandForksEmployee = true;
                                        importedEmployee.ImportFields["OrganizationValue2"] = "GF";
                                    }
                                    isPara = StringSearch(cellString, "para") || StringSearch(cellString, "aide") || StringSearch(cellString, "van driver");
                                    isDriver = !isPara && (StringSearch(cellString, "driver"));
                                    break;
                                case "Date Received (Direct Deposit Authorization )":
                                    if (TryGetDateFromCell(cell, out employee.DateOfDirectDepositUpdateInWorkBright))
                                    {
                                    }
                                    break;
                            }
                        }
                    }

                    if (isPara)
                    {
                        float payRate = employee.IsAGrandForksEmployee ? GrandForksDefaultRates[Jobs.AIDE_SCHOOL] : FargoDefaultRates[Jobs.AIDE_SCHOOL];
                        importedEmployee.ImportFields["Rate_AidDlySchool"] = payRate.ToString();
                        employee.PayRates[Jobs.AIDE_SCHOOL] = payRate;

                        payRate = employee.IsAGrandForksEmployee ? GrandForksDefaultRates[Jobs.AIDE_CHARTER] : FargoDefaultRates[Jobs.AIDE_CHARTER];
                        importedEmployee.ImportFields["Rate_AidDlyChrter"] = payRate.ToString();
                        employee.PayRates[Jobs.AIDE_CHARTER] = payRate;
                    }
                    else if (isDriver)
                    {
                        float payRate = employee.IsAGrandForksEmployee ? GrandForksDefaultRates[Jobs.DRIVER_SCHOOL] : FargoDefaultRates[Jobs.DRIVER_SCHOOL];
                        importedEmployee.ImportFields["Rate_DrvrDlySchool"] = payRate.ToString();
                        employee.PayRates[Jobs.DRIVER_SCHOOL] = payRate;

                        payRate = employee.IsAGrandForksEmployee ? GrandForksDefaultRates[Jobs.DRIVER_CHARTER] : FargoDefaultRates[Jobs.DRIVER_CHARTER];
                        importedEmployee.ImportFields["Rate_DrvrSchoolChrtr"] = payRate.ToString();
                        employee.PayRates[Jobs.DRIVER_CHARTER] = payRate;
                    }

                    string ssn = "";
                    if ("" != socialSecurityNumberEntries[0])
                    {
                        if ("" != socialSecurityNumberEntries[1] && socialSecurityNumberEntries[0] != socialSecurityNumberEntries[1])
                        {
                            if (socialSecurityNumberEntries[0].Replace("-", "") != socialSecurityNumberEntries[1].Replace("-", ""))
                            {
                                Log("social security number mismatch for employee: " + employee.Name);
                            }
                        }
                        ssn = socialSecurityNumberEntries[0];
                    }
                    else if ("" != socialSecurityNumberEntries[1])
                    {
                        if (!socialSecurityNumberEntries[1].Contains('-'))
                        {
                            socialSecurityNumberEntries[1] = socialSecurityNumberEntries[1].Insert(3, "-").Insert(6, "-");
                        }
                        ssn = socialSecurityNumberEntries[1];
                    }
                    importedEmployee.ImportFields["SSN"] = ssn;
                    employee.SocialSecurityNumber = ssn;
                    CheckEmployeeNumberWithSocialSecurityNumber(employee);
                    if ("" != birthDateEntries[0])
                    {
                        double d = double.Parse(birthDateEntries[0]);
                        employee.BirthDate = DateTime.FromOADate(d);
                        importedEmployee.ImportFields["BirthDate"] = employee.BirthDate.ToShortDateString();
                    }
                    else if ("" != birthDateEntries[1])
                    {
                        double d = double.Parse(birthDateEntries[1]);
                        employee.BirthDate = DateTime.FromOADate(d);
                        importedEmployee.ImportFields["BirthDate"] = employee.BirthDate.ToShortDateString();
                    }
                }
            }

            workBook.Close();
            excelApp.Quit();

            //Marshal.ReleaseComObject(workBook);
            //Marshal.ReleaseComObject(excelApp);
        }

        private HashSet<string> BusProblems = new();
        public void ReadTimeSheets()
        {
            const int EMP_NUMBER_COLUMN = 2;
            const int EMP_NAME_COLUMN = 3;
            const int DAY_COLUMN = 6;
            const int PUNCH_IN_COLUMN = 7;
            const int PUNCH_OUT_COLUMN = 10;
            const int ROUNDED_TIME_COLUMN = 14;
            const int JOB_TYPE_COLUMN = 15;
            const int NOTES_COLUMN = 17;
            const int BUS_NUMBER_COLUMN = 34;

            if (!CheckForExcelFileOnDesktop("Timesheets.xlsx", out string filePath))
            {
                return;
            }
            var lastModified = System.IO.File.GetLastWriteTime(filePath);
            if (new DateTime(lastModified.Year, lastModified.Month, lastModified.Day).CompareTo(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)) < 0)
            {
                Log("Timesheets is old.", true);
            }
            Excel.Application excelApp = new Excel.Application();
            var fInfo = new FileInfo(filePath);
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            //var employeeScheduleData = LoadEmployeeScheduleData();
            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Excel.Range range = sheet.Range[sheet.Range["A6"], sheet.Range["B8"]];
                range = range.CurrentRegion;
                int rows = range.Value2.GetLength(0) + 6;
                //Excel.Range range = sheet.Range[sheet.Range["A6"]].CurrentRegion;
                range = sheet.Range[sheet.Range["A1"], sheet.Range["BZ" + rows]];
                var cellData = (Object[,])range.Value2;
                rows = cellData.GetLength(0);
                for (int rowNumber = 6; rowNumber <= rows; ++rowNumber)
                {
                    if (null != cellData[rowNumber, DAY_COLUMN])
                    {
                        if (!TryGetDateFromCell(cellData[rowNumber, DAY_COLUMN], out DateTime date))
                        {
                            Log("date == nothing for row: " + rowNumber);
                            continue;
                        }

                        if (!TryGetFloatFromCell(cellData[rowNumber, ROUNDED_TIME_COLUMN], out float time))
                        {
                            continue;
                        }
                        if (time < 0.1f)
                        {
                            continue;
                        }

                        if (!TryGetIntFromCell(cellData[rowNumber, EMP_NUMBER_COLUMN], out int employeeNumber))
                        {
                            Log("Couldn't get employee number", true);
                            continue;
                        }

                        Company company = employeeNumber == 1734 ? Company.VALLEY_BUS_COACHES : Company.VALLEY_BUS_LLC;
                        Shift shift = new(company);
                        shift.ShiftTime = time;
                        shift.Date = date;

                        shift.WeekNumber = date.CompareTo(FirstDayWeek2) < 0 ? 1 : 2;
                        if (date.AddDays(7).CompareTo(FirstDayWeek2) < 0 || date.AddDays(-7).CompareTo(FirstDayWeek2) > 0)
                        {
                            Log("ERROR: Date of shift: " + date.ToShortDateString() + " is not within 7 days of FirstDayWeek2 ( " + FirstDayWeek2.ToShortDateString() + ")", true);
                        }

                        TryParseTimeSpan(cellData[rowNumber, PUNCH_IN_COLUMN], out shift.ClockIn);
                        TryParseTimeSpan(cellData[rowNumber, PUNCH_OUT_COLUMN], out shift.ClockOut);

                        //if (!EmployeeDictionary.ContainsKey(employeeNumber))
                        //{
                        //    string name = null == cellData[rowNumber, EMP_NAME_COLUMN] ? "" : (null == cellData[rowNumber, EMP_NAME_COLUMN].ToString() ? "" : new string((cellData[rowNumber, EMP_NAME_COLUMN].ToString())));
                        //    EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, name) { IsPartialEntry = true });
                        //}

                        Employee employee = EmployeeDictionary[employeeNumber];
                        if (EmployeeIdsToIgnore.Contains(employeeNumber))
                        {
                            continue;
                        }
                        if (employee.IsPartialEntry)
                        {
                            Log("In Timesheets, Employee " + employeeNumber + " (" + employee.Name + ") was not found.", true);
                        }

                        if (TryGetIntFromCell(cellData[rowNumber, JOB_TYPE_COLUMN], out shift.JobInt))
                        {
                            shift.JobType = GetJobTypeFromCode(shift.JobInt);
                            if (employeeNumber == 1983/*chris clark*/ && (Jobs)shift.JobInt == Jobs.MECHANIC)
                            {
                                shift.JobInt = (int)Jobs.DRIVER_SCHOOL;
                            }
                            if (shift.JobType == Jobs.DRIVER_SCHOOL && !employee.PayRates.ContainsKey(shift.JobType))
                            {
                                shift.JobType = Jobs.NON_CDL_DRIVER;
                                if (!employee.IsSalaried && employee.IsANonCdlDriver())
                                {
                                    NonCdlDrivers.Add(employee);
                                }
                            }
                            if (shift.JobType == Jobs.VACATION && employee.IdNumber != 1734/*Brad Rust*/)
                            {
                                shift.CompanyName = Company.VALLEY_BUS_LLC;
                            }
                            else if (shift.JobType == Jobs.VACATION)
                            {
                                shift.CompanyName = Company.VALLEY_BUS_COACHES;
                            }
                        }
                        else
                        {
                            Log("Problem getting Job Code for code: " + cellData[rowNumber, JOB_TYPE_COLUMN].ToString(), true);
                        }

                        CheckForManualEntry(shift, employee);

                        TryGetStringFromCell(cellData[rowNumber, NOTES_COLUMN], out shift.Notes);
                        if (StringSearch(shift.Notes, "TNT"))
                        {
                            shift.JobType = Jobs.DRIVER_CHARTER;
                        }
                        if (StringSearch(shift.Notes, "per diem") || StringSearch(shift.Notes, "perdiem"))
                        {
                            shift.PerDiem = 45;
                        }
                        if (StringSearch(shift.Notes, "wf"))
                        { //typically, west fargo is determined as the location by job code (20), this is just for WF Paras //update: Wf para has it's own job code now, but I will just leave this because it shouldn't hurt anything.
                            shift.ShiftLocation = Location.WEST_FARGO;
                        }

                        if (null != cellData[rowNumber, BUS_NUMBER_COLUMN] && StringSearch(cellData[rowNumber, BUS_NUMBER_COLUMN].ToString(), "old"))
                        {
                            string? busNumberString = cellData[rowNumber, BUS_NUMBER_COLUMN].ToString();
                            if (null != busNumberString)
                            {
                                busNumberString = busNumberString.Replace("old", "");
                                busNumberString = busNumberString.Replace("Old", "");
                                busNumberString = busNumberString.Replace(" ", "");
                                if (int.TryParse(busNumberString, out shift.BusNumber))
                                {
                                }
                            }
                        }
                        else if (TryGetIntFromCell(cellData[rowNumber, BUS_NUMBER_COLUMN], out shift.BusNumber))
                        {
                            shift.IsAGrandForksShift = shift.BusNumber >= GF_MIN_BUS && shift.BusNumber <= GF_MAX_BUS;
                        }
                        else if (long.TryParse(cellData[rowNumber, BUS_NUMBER_COLUMN].ToString(), out long weirdBusNumber))
                        {
                            Log("Weird bus number encountered");
                            shift.BusNumber = Shift.WEST_FARGO_BUS_PLACE_HOLDER;
                        }
                        else
                        {
                            if (null == cellData[rowNumber, BUS_NUMBER_COLUMN] || !StringSearch(cellData[rowNumber, BUS_NUMBER_COLUMN].ToString(), "N/A"))
                            {
                                if (null != cellData[rowNumber, BUS_NUMBER_COLUMN] && null != cellData[rowNumber, BUS_NUMBER_COLUMN].ToString())
                                {
                                    if (!BusProblems.Contains(cellData[rowNumber, BUS_NUMBER_COLUMN].ToString()))
                                    {
                                        Log("Problem getting bus number for busName: " + cellData[rowNumber, BUS_NUMBER_COLUMN].ToString(), true);
                                        BusProblems.Add(cellData[rowNumber, BUS_NUMBER_COLUMN].ToString());
                                    }
                                }
                                else if (null != cellData[rowNumber, BUS_NUMBER_COLUMN - 1 /*bus name column*/])
                                {
                                    Log("Problem getting bus number for MobileID: " + cellData[rowNumber, BUS_NUMBER_COLUMN - 1].ToString());
                                }
                            }

                            if (shift.JobInt == 20 || shift.JobInt == 23)
                            {
                                shift.BusNumber = Shift.WEST_FARGO_BUS_PLACE_HOLDER;
                            }
                        }

                        if (shift.IsASchoolRouteShift() && !StringSearch(shift.Notes, "training"))
                        {
                            if (shift.JobInt == 20 || shift.JobInt == 23)
                            {
                                shift.ShiftLocation = Location.WEST_FARGO;
                            }
                            else if (shift.IsAGrandForksShift || employee.IsAGrandForksEmployee)
                            {
                                shift.ShiftLocation = Location.GRAND_FORKS;
                            }
                            else
                            {
                                shift.ShiftLocation = Location.FARGO;
                            }
                            if (shift.JobType == Jobs.DRIVER_SCHOOL)
                            {
                                Shift.DailySchoolRouteCounter[(int)shift.ShiftLocation, date.Day] += 1;
                            }
                        }

                        //if (shift.IsASchoolRouteShift())
                        //{
                        //    CheckShiftAgainstSchedule(shift, employee, employeeScheduleData);
                        //}

                        employee.Shifts.Add(shift);
                    }
                }
            }

            workBook.Close();
            excelApp.Quit();

            LogSchedulingData();
            //Marshal.ReleaseComObject(workBook);
            //Marshal.ReleaseComObject(excelApp);
        }

        static HashSet<int> LoggedEmployees = new();
        static Dictionary<string, List<string>> SchedulingLogMessages = new();
        void LogSchedulingData()
        {
            foreach (var kvp in SchedulingLogMessages)
            {
                Log("");
                foreach (var message in kvp.Value)
                {
                    Log(message);
                }
            }
        }
        static HashSet<string> EarlyOutSignals = new();
        static void CheckShiftAgainstSchedule(Shift shift, Employee employee, Dictionary<int, Dictionary<RouteTimeContext, TimeSpan>> employeeScheduleData)
        {
            if (shift.IsASummerRoute())
            {
                return;
            }
            if (!shift.IsASchoolRouteShift())
            {
                return;
            }
            foreach (var kvp in employeeScheduleData)
            {
                if (kvp.Key == employee.IdNumber)
                {
                    if (!kvp.Value.ContainsKey(shift.TimeContext()))
                    {
                        continue;
                    }

                    TimeSpan earliestClockIn = kvp.Value[shift.TimeContext()];
                    if (employee.ScheduleExceptions.ContainsKey(shift.Date.DayOfWeek) && employee.ScheduleExceptions[shift.Date.DayOfWeek].ContainsKey(shift.TimeContext()))
                    {
                        earliestClockIn = employee.ScheduleExceptions[shift.Date.DayOfWeek][shift.TimeContext()];
                    }
                    if (shift.ClockIn.CompareTo(earliestClockIn) < 0)
                    {
                        if (shift.ClockOut.CompareTo(earliestClockIn) < 0)
                        {
                            //something is going on but it's probably not an early punch in
                            continue;
                        }
                        var originalClockIn = shift.ClockIn;
                        TimeSpan difference = earliestClockIn - shift.ClockIn;

                        if (difference.CompareTo(new TimeSpan(1, 30, 0)) > 0 && shift.Date.DayOfWeek == DayOfWeek.Wednesday && shift.IsAGrandForksShift)
                        {
                            Log("Skipping shift because it is probably an early out on Wednesday in GF");
                            continue;
                        }

                        if (difference.CompareTo(new TimeSpan(1, 30, 0)) > 0 && shift.TimeContext() == RouteTimeContext.AFTERNOON)
                        {
                            if (EarlyOutSignals.Contains(shift.Date.ToShortDateString()))
                            {
                                Log("Was there an early out on this day: " + shift.Date.ToShortDateString(), true);
                            }
                            EarlyOutSignals.Add(shift.Date.ToShortDateString());
                        }

                        if (shift.BusNumber != 0 && shift.BusNumber != Shift.WEST_FARGO_BUS_PLACE_HOLDER)
                        {
                            int totalShiftsForContext = employee.BusShiftTotals[shift.Date.DayOfWeek][shift.TimeContext()];
                            if (employee.ShiftsByBusNumber[shift.Date.DayOfWeek][shift.TimeContext()][shift.BusNumber] < totalShiftsForContext / 2)
                            {
                                Log("Not docking time for " + employee.Name + " because this bus was not used a majority of the time for this context.\nOriginal clock in time: " + originalClockIn.ToString() + "\nNew time: " + earliestClockIn.ToString());
                                continue;
                            }
                        }

                        shift.ModifyClockIn(earliestClockIn);

                        if (originalClockIn.CompareTo(shift.ClockIn) == 0)
                        {
                            Log("Shift invalidated for " + employee.Name + ".\nOriginal clock in time: " + originalClockIn.ToString() + "\nNew time: " + shift.ClockIn.ToString());
                        }
                        if (difference.CompareTo(new TimeSpan(0, 25, 0)) > 0)
                        {
                            if (shift.TimeContext() == RouteTimeContext.MORNING)
                            {
                                continue;
                            }

                            //Log("Modifying clock in for " + employee.Name + " by " + difference.ToString() + "\nOriginal clock in time: " + originalClockIn.ToString() + "\nNew time: " + shift.ClockIn.ToString(), false/*!LoggedEmployees.Contains(employee.IdNumber)*/);
                            LoggedEmployees.Add(employee.IdNumber);
                        }

                        if (difference.CompareTo(new TimeSpan(0, 15, 0)) > 0)
                        {
                            if (!SchedulingLogMessages.ContainsKey(employee.Name))
                            {
                                SchedulingLogMessages.Add(employee.Name, new());
                            }
                            string message = "For " + employee.Name + " on " + shift.Date.ToShortDateString() + " changing clock in time from " + originalClockIn.ToString() + " to " + shift.ClockIn.ToString();
                            SchedulingLogMessages[employee.Name].Add(message);
                            LoggedEmployees.Add(employee.IdNumber);
                        }
                    }
                }
            }
        }

        public void ReadCoachesPayroll()
        {
            const int EMP_NAME_COLUMN = 1;
            const int EMP_NUMBER_COLUMN = 3;
            const int DATE_RANGE_COLUMN = 5;
            const int DOLLARS_COLUMN = 9;
            const int PER_DIEM_COLUMN = 11;
            const int BONUS_COLUMN = 13;
            const int BUS_NUMBER_COLUMN = 15;
            const int HOURS_COLUMN = 17;

            if (!CheckForExcelFileOnDesktop("CoachesPayroll.xlsx", out string filePath))
            {
                return;
            }
            var lastModified = File.GetLastWriteTime(filePath);
            if (new DateTime(lastModified.Year, lastModified.Month, lastModified.Day).CompareTo(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)) < 0)
            {
                if (PrintForm.InputBool("CoachesPayroll is old. Would you like to skip CoachesPayroll?"))
                {
                    return;
                }
            }
            Excel.Application excelApp = new();
            var fInfo = new FileInfo(filePath);
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            int employeeNumber = 0; //employee number persists for multiple rows
            bool[] bCompanyWasFound = new bool[2];
            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Company company = StringSearch(sheet.Name, "Coaches") ? Company.VALLEY_BUS_COACHES : Company.VALLEY_BUS_LLC;
                bCompanyWasFound[(int)company] = true;
                Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["z1000"]];
                var cellData = (Object[,])range.Value2;
                for (int rowNumber = 2; rowNumber < cellData.GetLength(0); ++rowNumber)
                {
                    if (null != cellData[rowNumber, DATE_RANGE_COLUMN])
                    {
                        string? date = cellData[rowNumber, DATE_RANGE_COLUMN].ToString();
                        if (date == "date" || date == "Date")
                        {
                            //header row
                            continue;
                        }
                        if (date == null || date == "" || date == " ")
                        {
                            Log("date == nothing for row: " + rowNumber);
                            continue;
                        }

                        if (!TryGetIntFromCell(cellData[rowNumber, EMP_NUMBER_COLUMN], out int eNumber) && employeeNumber == 0)
                        {
                            Log("Couldn't get employee number", true);
                            continue;
                        }

                        if (eNumber != 0)
                        {
                            employeeNumber = eNumber;
                        }

                        if (!EmployeeDictionary.ContainsKey(employeeNumber))
                        {
                            string name = null == cellData[rowNumber, EMP_NAME_COLUMN] ? "" : (null == cellData[rowNumber, EMP_NAME_COLUMN].ToString() ? "" : new string((cellData[rowNumber, EMP_NAME_COLUMN].ToString())));
                            Log("In Coaches Payroll, Employee " + employeeNumber + " (" + name + ") was not found.", true);
                            EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, name));
                        }
                        Employee employee = EmployeeDictionary[employeeNumber];
                        //todo: check employee name to find data entry errors

                        List<DateTime> dates = GetDatesFromCoachesDateRange(date);

                        TryGetFloatFromCell(cellData[rowNumber, DOLLARS_COLUMN], out float dollars);
                        TryGetFloatFromCell(cellData[rowNumber, BONUS_COLUMN], out float bonus);
                        TryGetFloatFromCell(cellData[rowNumber, PER_DIEM_COLUMN], out float perDiem);
                        TryGetIntFromCell(cellData[rowNumber, BUS_NUMBER_COLUMN], out int busNumber);
                        TryGetFloatFromCell(cellData[rowNumber, HOURS_COLUMN], out float hours);



                        float payRate = 0f;
                        if (hours > 0.001f)
                        {
                            payRate = (float)Math.Round((bonus + dollars) / hours, 2);
                            if (Math.Abs(payRate - T_AND_J_CHARTER_RATE) < 0.01f)
                            {
                                payRate = T_AND_J_CHARTER_RATE;
                            }
                            dollars = 0f;
                            bonus = 0f;
                        }
                        else if (busNumber > 699 && busNumber < 800 && bonus > 0.01f) //if bonus is 0, then the employee probably got the benefit of a minimum guarantee so they don't get overtime.
                        {
                            hours = (dollars + bonus) / T_AND_J_CHARTER_RATE;
                            payRate = T_AND_J_CHARTER_RATE;
                            dollars = 0f;
                            bonus = 0f;
                        }

                        List<Shift> shifts = new();

                        int coachTripDays = hours < 0.01f && dollars + bonus > 0 ? Math.Max(1, dates.Count) : 0;

                        if (dates.Count > 1 && dates[0].CompareTo(FirstDayWeek2) != dates[^1].CompareTo(FirstDayWeek2) && dollars + hours > 0)
                        { //multiple shifts and different weeks.


                            shifts = dates.Select(date =>
                            {
                                Shift shift = new(company)
                                {
                                    Date = date,
                                    WeekNumber = date.CompareTo(FirstDayWeek2) < 0 ? 1 : 2,
                                    JobType = Jobs.DRIVER_COACH,
                                    DollarAmount = dollars / dates.Count,
                                    BonusDollars = date.Equals(dates[0]) ? bonus : 0,
                                    PerDiem = date.Equals(dates[0]) ? perDiem : 0,
                                    ShiftTime = hours / dates.Count,
                                    PayRate = payRate > 1 ? payRate : null,
                                    BusNumber = busNumber,
                                    CoachTripDays = 1
                                };
                                return shift;
                            }).ToList();
                        }
                        else
                        {
                            shifts.Add(new(company)
                            {
                                WeekNumber = dates[^1].CompareTo(FirstDayWeek2) < 0 ? 1 : 2,
                                JobType = Jobs.DRIVER_COACH,
                                DollarAmount = dollars,
                                BonusDollars = bonus,
                                PerDiem = perDiem,
                                ShiftTime = hours,
                                PayRate = payRate > 1 ? payRate : null,
                                BusNumber = busNumber,
                                CoachTripDays = coachTripDays
                            });

                        }

                        shifts.ForEach(shift => employee.Shifts.Add(shift));
                    }
                }
            }
            for (int i = 0; i < 2; i++)
            {
                if (!bCompanyWasFound[i])
                {
                    DelayedLog("ERROR: Couldn't find company " + ((Company)i).ToString() + " in Coaches Payroll. Please make sure one sheets contains the word 'Coaches' and the other does not.", true);
                }
            }

            excelApp.Quit();

            //Marshal.ReleaseComObject(workBook);
            //Marshal.ReleaseComObject(excelApp);
        }

        public Dictionary<int, Dictionary<RouteTimeContext, TimeSpan>> LoadEmployeeScheduleData()
        {
            if (!PrintForm.InputBool("Please check that Driver-Para-Schedule.xlsx is synced with OneDrive.", "Okay", "Skip"))
            {
                return new();
            }
            Log("Please check that Driver-Para-Schedule.xlsx is synced with OneDrive.", true);
            Dictionary<int, Dictionary<RouteTimeContext, TimeSpan>> employeeScheduleData = new();

            Excel.Application xlApp = new();
            String path = "C:/Users/User/valleybusllc.com/Admin Team - Payroll - Payroll/Payroll/Driver-Para-Schedule.xlsx";
            Excel.Workbook workBook = xlApp.Workbooks.Open(path);

            foreach (Excel.Worksheet workSheet in workBook.Worksheets)
            {
                Excel.Range range = workSheet.Range[workSheet.Range["A1"], workSheet.Range["Z400"]];
                var cellData = (Object[,])range.Value2;
                int rows = range.Value2.GetLength(0) + 1;
                HashSet<int> employeesWhoseIdsHaveBeenChecked = new();
                if (TryGetStringFromCell(cellData[400, 26], out String cellString))
                {
                    employeesWhoseIdsHaveBeenChecked = cellString
                        .Split(',')
                        .Select(int.Parse)
                        .ToHashSet();
                }


                for (int row = 1; row < rows; ++row)
                {
                    int employeeNameColumn = 1;
                    int employeeNumberColumn = 2;
                    int exceptionColumn = 6;
                    int columnOffset = employeeNumberColumn + 1;
                    if (TryGetIntFromCell(cellData[row, employeeNumberColumn], out int employeeNumber))
                    {
                        if (TryGetStringFromCell(cellData[row, employeeNameColumn], out string nameFromSheet))
                        {
                            if (!EmployeeDictionary.ContainsKey(employeeNumber))
                            {
                                Log("Problem in attendance schedule! , " + employeeNumber.ToString() + " isn't an employee");
                                continue;
                            }
                            if (!employeesWhoseIdsHaveBeenChecked.Contains(employeeNumber))
                            {
                                Log("From attendance schedule, " + nameFromSheet + "(" + EmployeeDictionary[employeeNumber].Name + ")");
                                employeesWhoseIdsHaveBeenChecked.Add(employeeNumber);
                            }
                        }
                        for (int column = 0; column <= (int)RouteTimeContext.AFTERNOON; column++)
                        {
                            if (TryGetDateFromCell(cellData[row, column + columnOffset], out DateTime dateTime))
                            {
                                TimeSpan timeSpan = dateTime.TimeOfDay;
                                if ((RouteTimeContext)column == RouteTimeContext.AFTERNOON && timeSpan.CompareTo(new TimeSpan(11, 59, 59)) < 0)
                                {
                                    //time wasn't put in as 24 hour
                                    timeSpan = timeSpan.Add(new TimeSpan(12, 0, 0));
                                }
                                if (!employeeScheduleData.ContainsKey(employeeNumber))
                                {
                                    employeeScheduleData.Add(employeeNumber, new());
                                }
                                employeeScheduleData[employeeNumber][(RouteTimeContext)column] = timeSpan;
                            }
                        }

                        if (TryGetStringFromCell(cellData[row, exceptionColumn], out string exceptionInstructions))
                        {
                            Employee employee = EmployeeDictionary[employeeNumber];
                            columnOffset = exceptionColumn + 1;
                            DayOfWeek dayOfWeek = DayOfWeek.Sunday;
                            if (StringSearch(exceptionInstructions, "wed"))
                            {
                                dayOfWeek = DayOfWeek.Wednesday;
                            }
                            else
                            {
                                Log("Can't determine day for exception: " + exceptionInstructions, true);
                            }
                            for (int column = 0; column <= (int)RouteTimeContext.AFTERNOON; column++)
                            {
                                if (TryGetDateFromCell(cellData[row, column + columnOffset], out DateTime dateTime))
                                {
                                    TimeSpan timeSpan = dateTime.TimeOfDay;
                                    if ((RouteTimeContext)column == RouteTimeContext.AFTERNOON && timeSpan.CompareTo(new TimeSpan(11, 59, 59)) < 0)
                                    {
                                        //time wasn't put in as 24 hour
                                        timeSpan = timeSpan.Add(new TimeSpan(12, 0, 0));
                                    }
                                    if (!employee.ScheduleExceptions.ContainsKey(dayOfWeek))
                                    {
                                        employee.ScheduleExceptions.Add(dayOfWeek, new());
                                    }
                                    if (!employee.ScheduleExceptions[dayOfWeek].ContainsKey((RouteTimeContext)column))
                                    {
                                        employee.ScheduleExceptions[dayOfWeek].Add((RouteTimeContext)column, new());
                                    }
                                    employee.ScheduleExceptions[dayOfWeek][(RouteTimeContext)column] = timeSpan;
                                }
                            }
                        }


                        //Object[,] employeesWhoseIdsHaveBeenCheckedObject = new String[1, 1];
                        //employeesWhoseIdsHaveBeenCheckedObject[0, 0] = String.Join(",", employeesWhoseIdsHaveBeenChecked);
                        //workSheet.Range["Z" + 400].Value = employeesWhoseIdsHaveBeenCheckedObject;

                        //SaveWorkBook(workBook, path);
                    }
                }
            }
            workBook.Close();
            xlApp.Quit();

            return employeeScheduleData;
        }

        public void WritePayrollImports()
        {
            Excel.Application xlApp = new();
            xlApp.DisplayAlerts = false;
            object misValue = System.Reflection.Missing.Value;

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            List<Employee> SortedEmployees = (from c in EmployeeDictionary
                                              orderby c.Key
                                              select c.Value).ToList();

            for (int company = (int)Company.VALLEY_BUS_LLC; company <= (int)Company.VALLEY_BUS_COACHES; ++company)
            {
                const int ROW_COUNT = 5000;
                object[,] matrix = new object[ROW_COUNT, 26];

                Excel.Workbook? workBook = null;
                Excel.Worksheet? workSheet = null;
                if ((Company)company == Company.VALLEY_BUS_LLC)
                {
                    string filePath = DesktopPath() + "Timesheets.xlsx";
                    var fInfo = new FileInfo(filePath);
                    if (fInfo.Exists)
                    {
                        workBook = xlApp.Workbooks.Open(filePath);
                    }
                    if (null == workBook)
                    {
                        //create new workbook
                        workBook = xlApp.Workbooks.Add(misValue);
                    }
                    workSheet = workBook.Worksheets.Add(misValue);
                }
                else
                {
                    string filePath = DesktopPath() + "MotorCoach_TimeCardImport.xlsx";
                    var fInfo = new FileInfo(filePath);
                    fInfo = new FileInfo(filePath);
                    if (fInfo.Exists)
                    {
                        workBook = xlApp.Workbooks.Open(filePath);
                    }
                    if (null == workBook)
                    {
                        //create new workbook
                        workBook = xlApp.Workbooks.Add(misValue);
                    }
                    workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                }

                WriteHeadersForTimeCardImport(workSheet);
                int rowCounter = 0;
                foreach (var emp in SortedEmployees)
                {
                    if (emp.IdNumber == 1768 && emp.Shifts.Count > 0)
                    {
                        //Log("Eddie Peltier is being ignored by payroll", true);
                        //continue;
                    }
                    if (emp != null)
                    {
                        if (!Program.EmployeeIdsToIgnore.Contains(emp.IdNumber) && emp.Shifts.Count > 0)
                        {
                            //if (!emp.HasAnActiveDirectDepositAccount)
                            //{
                            //    DelayedLog("Employee: " + emp.Name + " (" + emp.IdNumber + ") has no active DD account. Phone: " + emp.PhoneNumber);
                            //}
                            if (emp.IsPartialEntry)
                            {
                                if (!emp.WasReportedForPartialEntry)
                                {
                                    emp.WasReportedForPartialEntry = true;
                                    Log("Employee: " + emp.Name + " (" + emp.IdNumber + ") was not found in payroll or on the employee export.", true);
                                }
                                continue;
                            }
                            if (null == emp.SocialSecurityNumber || emp.SocialSecurityNumber == "")
                            {
                                if (!emp.WasReportedForPartialEntry)
                                {
                                    emp.WasReportedForPartialEntry = true;
                                    Log(emp.Name + " (" + emp.IdNumber + ") is not getting paid because they do not have a social security number in workbright.", true);
                                }
                                continue;
                            }
                            for (int shiftType = 0; shiftType < 3; ++shiftType)
                            {
                                if (null != emp.ShiftTotals[company, shiftType])
                                {
                                    foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                                    {
                                        foreach (var shifts in pair.Values)
                                        {
                                            foreach (Shift shift in shifts)
                                            {
                                                if (!shift.IsValid(emp))
                                                {
                                                    continue;
                                                }
                                                if (shift.CompanyName == (Company)company && shift.IsValid(emp))
                                                {
                                                    if (shift.ShiftTime + shift.DollarAmount + shift.BonusDollars + shift.PerDiem > 0f)
                                                    {
                                                        if (shift.JobType == Jobs.VACATION)
                                                        {
                                                            WriteToMatrix(emp, shift, shift.ShiftTime, 0f, TimeCardImportColumns.VACATION_HOURS, TimeCardImportColumns.VACATION_WEEK, null, ref rowCounter, matrix);
                                                        }
                                                        else if (shift.JobType == Jobs.HOLIDAY)
                                                        {
                                                            WriteToMatrix(emp, shift, shift.ShiftTime, 0f, TimeCardImportColumns.HOLIDAY_HOURS, TimeCardImportColumns.HOLIDAY_WEEK, null, ref rowCounter, matrix);
                                                        }
                                                        else
                                                        {
                                                            WriteToMatrix(emp, shift, shift.ShiftTime, shift.DollarAmount, TimeCardImportColumns.REGULAR_HOURS, TimeCardImportColumns.REGULAR_HOURS_WEEK, TimeCardImportColumns.REGULAR_DOLLARS, ref rowCounter, matrix);
                                                        }
                                                    }
                                                    if (shift.MinimumGuaranteeHours > 0f)
                                                    {
                                                        WriteToMatrix(emp, shift, shift.MinimumGuaranteeHours, 0f, TimeCardImportColumns.MG_HOURS, TimeCardImportColumns.MG_WEEK, null, ref rowCounter, matrix);
                                                    }
                                                    if (shift.SummerGuaranteeHours > 0f)
                                                    {
                                                        WriteToMatrix(emp, shift, shift.SummerGuaranteeHours, 0f, TimeCardImportColumns.SUMMER_BONUS_HOURS, TimeCardImportColumns.SUMMER_BONUS_WEEK, null, ref rowCounter, matrix);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            for (int weekNumber = 1; weekNumber < 3; ++weekNumber)
                            {
                                if (emp.OverTimeHours[company, weekNumber] > 0f)
                                {
                                    matrix[rowCounter, (int)TimeCardImportColumns.EMP_NUMBER] = emp.IdNumber;
                                    matrix[rowCounter, (int)TimeCardImportColumns.JOB_CODE] = "OT";
                                    matrix[rowCounter, (int)TimeCardImportColumns.OT_HOURS] = Math.Round(emp.OverTimeHours[company, weekNumber], 2);
                                    matrix[rowCounter, (int)TimeCardImportColumns.OT_WEEK] = weekNumber;
                                    ++rowCounter;
                                }
                            }
                        }
                    }
                }

                //TODO: check that our matrix doesn't get cut off - it shouldn't because we are using 5000 instead of a dynamic number;
                Excel.Range range = workSheet.Range[workSheet.Range["A2"], workSheet.Range["Z" + ROW_COUNT]];

                range.Value = matrix;

                if ((Company)company == Company.VALLEY_BUS_LLC)
                {
                    SaveWorkBook(workBook, DesktopPath() + "Timesheets1.xlsx");
                    ((Excel.Worksheet)workBook.Worksheets.get_Item(2)).Delete();
                    SaveWorkBook(workBook, DesktopPath() + "VB_TimeCardImport.xlsx");

                }
                else
                {
                    SaveWorkBook(workBook, DesktopPath() + "MotorCoach_TimeCardImport.xlsx");
                }

                workBook.Close(true, misValue, misValue);
                //Marshal.ReleaseComObject(workSheet);
                //Marshal.ReleaseComObject(workBook);
            }

            xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);

            var p = new Process
            {
                StartInfo = new ProcessStartInfo(DesktopPath() + "VB_TimeCardImport.xlsx")
                {
                    UseShellExecute = true
                }
            };
            p.Start();
            p = new Process
            {
                StartInfo = new ProcessStartInfo(DesktopPath() + "MotorCoach_TimeCardImport.xlsx")
                {
                    UseShellExecute = true
                }
            };
            p.Start();
        }

        public void WriteWfnPayrollImports()
        {
            List<Employee> sortedEmployees = EmployeeDictionary
                .OrderBy(kvp => kvp.Key)
                .Select(kvp => kvp.Value)
                .ToList();

            string batchId = MakeWfnBatchId();

            for (int company = (int)Company.VALLEY_BUS_LLC; company <= (int)Company.VALLEY_BUS_COACHES; ++company)
            {
                List<Dictionary<string, string>> rows = new();

                foreach (Employee emp in sortedEmployees)
                {
                    if (emp == null || EmployeeIdsToIgnore.Contains(emp.IdNumber) || emp.Shifts.Count == 0)
                    {
                        continue;
                    }
                    if (emp.IsPartialEntry)
                    {
                        if (!emp.WasReportedForPartialEntry)
                        {
                            emp.WasReportedForPartialEntry = true;
                            Log("Employee: " + emp.Name + " (" + emp.IdNumber + ") was not found in payroll or on the employee export.", true);
                        }
                        continue;
                    }
                    if (string.IsNullOrWhiteSpace(emp.SocialSecurityNumber))
                    {
                        if (!emp.WasReportedForPartialEntry)
                        {
                            emp.WasReportedForPartialEntry = true;
                            Log(emp.Name + " (" + emp.IdNumber + ") is not getting paid because they do not have a social security number in workbright.", true);
                        }
                        continue;
                    }

                    bool employeeHasWfnPayRowsForThisCompany = false;
                    bool bEmployeeHasShownVacation = false;
                    for (int weekNumber = 1; weekNumber < 3; ++weekNumber)
                    {
                        int firstRowForThisEmployeeCompanyWeek = rows.Count;
                        float regularHourCapMultiplier = CalculateWfnRegularHourCapMultiplier(emp, (Company)company, weekNumber);

                        for (int shiftType = 0; shiftType < 3; ++shiftType)
                        {
                            if (emp.ShiftTotals[company, shiftType] == null)
                            {
                                continue;
                            }

                            float vacationHours = emp.VacationHours + emp.NetVacationChangeForPayPeriod;
                            if (vacationHours != 0f && !bEmployeeHasShownVacation)
                            {
                                bEmployeeHasShownVacation = true;
                                Company com = emp.IdNumber == 1734 || emp.IdNumber == 123 ? Company.VALLEY_BUS_COACHES : Company.VALLEY_BUS_LLC;
                                Dictionary<string, string> row = MakeBaseWfnRow(emp, com, batchId);
                                //row["VBL"] = vacationHours.ToString();
                            }

                            foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                            {
                                foreach (var weekEntry in pair)
                                {
                                    if (weekEntry.Key != weekNumber)
                                    {
                                        continue;
                                    }
                                    foreach (Shift shift in weekEntry.Value)
                                    {
                                        if (shift.CompanyName != (Company)company || !shift.IsValid(emp))
                                        {
                                            continue;
                                        }

                                        if (shift.ShiftTime + shift.DollarAmount + shift.BonusDollars + shift.PerDiem > 0f)
                                        {
                                            if (shift.JobType == Jobs.VACATION)
                                            {
                                                AddWfnHours3Row(rows, emp, (Company)company, batchId, shift, shift.ShiftTime, "VAC");
                                            }
                                            else if (shift.JobType == Jobs.HOLIDAY)
                                            {
                                                AddWfnHours3Row(rows, emp, (Company)company, batchId, shift, shift.ShiftTime, "HOL");
                                            }
                                            else
                                            {
                                                AddWfnRegularRow(rows, emp, (Company)company, batchId, shift, shift.ShiftTime, shift.DollarAmount, regularHourCapMultiplier);
                                            }
                                            employeeHasWfnPayRowsForThisCompany = true;
                                        }

                                        if (shift.MinimumGuaranteeHours > 0f)
                                        {
                                            AddWfnHours3Row(rows, emp, (Company)company, batchId, shift, shift.MinimumGuaranteeHours, "MNG");
                                            employeeHasWfnPayRowsForThisCompany = true;
                                        }
                                        if (shift.SummerGuaranteeHours > 0f)
                                        {
                                            AddWfnHours3Row(rows, emp, (Company)company, batchId, shift, shift.SummerGuaranteeHours, "MNG");
                                            employeeHasWfnPayRowsForThisCompany = true;
                                        }

                                        if (!shift.ExtrasWereWrittenToExport)
                                        {
                                            shift.ExtrasWereWrittenToExport = true;
                                            if (shift.BonusDollars > 0f)
                                            {
                                                AddWfnEarnings3Row(rows, emp, (Company)company, batchId, shift, shift.BonusDollars, "BON");
                                                employeeHasWfnPayRowsForThisCompany = true;
                                            }
                                            if (shift.PerDiem > 0f)
                                            {
                                                AddWfnAdjustmentRow(rows, emp, (Company)company, batchId, shift, shift.PerDiem, "PDM");
                                                employeeHasWfnPayRowsForThisCompany = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (emp.OverTimeHours[company, weekNumber] > 0f)
                        {
                            AdjustFinalRegularRowToExactly40Hours(rows, firstRowForThisEmployeeCompanyWeek);
                            AddWfnOvertimeRow(rows, emp, (Company)company, batchId, emp.OverTimeHours[company, weekNumber], weekNumber);
                            employeeHasWfnPayRowsForThisCompany = true;
                        }
                    }

                    if (employeeHasWfnPayRowsForThisCompany && emp.IsSalaried && emp.PrimaryCompany == (Company)company)
                    {
                        AddWfnSalaryRows(rows, emp, (Company)company, batchId);
                    }
                }

                //manual entries always go at the end of the .csv file
                AddManualEntryWfnRows(rows, (Company)company, batchId);

                string path = DesktopPath() + ((Company)company == Company.VALLEY_BUS_LLC ? "EPIMMFAA.csv" : "EPIMKZAA.csv");
                WriteCsv(path, WfnPayrollImportHeaders, rows);

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(path)
                    {
                        UseShellExecute = true
                    }
                };
                p.Start();
            }
        }

        private static readonly List<string> WfnPayrollImportHeaders = new()
        {
            "Co Code",
            "Batch ID",
            "File #",
            "Reg Hours",
            "Reg Earnings",
            "O/T Hours",
            "Hours 3 Code",
            "Hours 3 Amount",
            "Earnings 3 Code",
            "Earnings 3 Amount",
            "Adjust Ded Code",
            "Adjust Ded Amount",
            "FLSA Workweek",
            "Temp Dept",
            "Temp Rate",
            //"VBL"
        };

        private static string GetWfnCompanyCode(Company company)
        {
            return company == Company.VALLEY_BUS_COACHES ? "MKZ" : "MMF";
        }

        private static string MakeWfnBatchId()
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            Random random = new();
            return new string(Enumerable.Range(0, 8).Select(_ => chars[random.Next(chars.Length)]).ToArray());
        }

        private static Dictionary<string, string> MakeBaseWfnRow(Employee emp, Company company, string batchId)
        {
            Dictionary<string, string> row = WfnPayrollImportHeaders.ToDictionary(header => header, _ => "");
            row["Co Code"] = GetWfnCompanyCode(company);
            row["Batch ID"] = batchId;
            row["File #"] = emp.IdNumber.ToString();
            return row;
        }

        private static void SetWfnTempDepartment(Dictionary<string, string> row, Shift? shift, bool bIsOvertime = false)
        {
            if (null == shift && !bIsOvertime)
            {
                Log("SetWfnTempDepartment was called with a null shift and bIsOvertime is false.", true);
                return;
            }
            string departmentCode = bIsOvertime ? "000029" : Shift.GetDepartmentCode(shift.JobType);
            if (string.IsNullOrWhiteSpace(departmentCode))
            {
                Log("string.IsNullOrWhiteSpace(departmentCode)", true);
            }
            row["Temp Dept"] = departmentCode.PadLeft(6, '0');
        }
        private static void SetWfnFlsaWorkweek(Dictionary<string, string> row, Shift shift)
        {
            row["FLSA Workweek"] = shift.WeekNumber.ToString();
        }

        private static void AddWfnRegularRow(List<Dictionary<string, string>> rows, Employee emp, Company company, string batchId, Shift shift, float hours, float dollars, float regularHourCapMultiplier)
        {
            Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
            SetWfnTempDepartment(row, shift);
            SetWfnFlsaWorkweek(row, shift);

            float cappedHours = hours * regularHourCapMultiplier;
            float cappedDollars = dollars * regularHourCapMultiplier;

            if (cappedHours > 0.001f)
            {
                row["Reg Hours"] = FormatWfnNumber(cappedHours);
            }
            if (cappedDollars > 0.001f)
            {
                row["Reg Earnings"] = FormatWfnNumber(cappedDollars);
            }
            else if (shift.PayRate > 0f)
            {
                row["Temp Rate"] = FormatWfnNumber(shift.PayRate.Value);
            }
            else if (cappedHours > 0.01f)
            {
                Log("No payrate or dollar amount found for WFN regular row for " + emp.Name + ".", true);
            }
            rows.Add(row);
        }

        private static float CalculateWfnRegularHourCapMultiplier(Employee emp, Company company, int weekNumber)
        {
            float regularHours = 0f;

            for (int shiftType = 0; shiftType < 3; ++shiftType)
            {
                if (emp.ShiftTotals[(int)company, shiftType] == null)
                {
                    continue;
                }

                foreach (var pair in emp.ShiftTotals[(int)company, shiftType].Values)
                {
                    foreach (var weekEntry in pair)
                    {
                        if (weekEntry.Key != weekNumber)
                        {
                            continue;
                        }

                        foreach (Shift shift in weekEntry.Value)
                        {
                            if (shift.CompanyName != company || !shift.IsValid(emp))
                            {
                                continue;
                            }

                            if (shift.JobType == Jobs.VACATION || shift.JobType == Jobs.HOLIDAY)
                            {
                                continue;
                            }

                            regularHours += shift.ShiftTime;
                        }
                    }
                }
            }

            if (regularHours <= 40.0001f)
            {
                return 1f;
            }

            return 40f / regularHours;
        }

        private static void AdjustFinalRegularRowToExactly40Hours(
    List<Dictionary<string, string>> rows,
    int firstRowIndexForEmployeeCompanyWeek)
        {
            float totalRegHours = 0f;
            int finalRegRowIndex = -1;

            for (int i = firstRowIndexForEmployeeCompanyWeek; i < rows.Count; i++)
            {
                if (float.TryParse(rows[i].GetValueOrDefault("Reg Hours", ""), out float regHours) && regHours > 0.001f)
                {
                    totalRegHours += regHours;
                    finalRegRowIndex = i;
                }
            }

            if (finalRegRowIndex < 0)
            {
                return;
            }

            float difference = (float)Math.Round(40f - totalRegHours, 2);

            if (Math.Abs(difference) >= 0.01f)
            {
                float currentFinalHours = float.Parse(rows[finalRegRowIndex]["Reg Hours"]);
                rows[finalRegRowIndex]["Reg Hours"] = FormatWfnNumber(currentFinalHours + difference);
            }
        }

        private static void AddWfnOvertimeRow(List<Dictionary<string, string>> rows, Employee emp, Company company, string batchId, float hours, int weekNumber)
        {
            if (hours < 0.001f)
            {
                return;
            }

            Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
            row["O/T Hours"] = FormatWfnNumber(hours);

            float blendedRegularRate = CalculateWfnBlendedRegularRate(emp, company, weekNumber);
            if (blendedRegularRate > 0.001f)
            {
                row["Temp Rate"] = FormatWfnNumber(blendedRegularRate);
            }
            SetWfnTempDepartment(row, null, true);

            row["FLSA Workweek"] = weekNumber.ToString();
            rows.Add(row);
        }

        private static float CalculateWfnBlendedRegularRate(Employee emp, Company company, int weekNumber)
        {
            float regularHours = 0f;
            float regularEarnings = 0f;
            float tempDepartmentShiftHours = 0f;

            for (int shiftType = 0; shiftType < 3; ++shiftType)
            {
                if (emp.ShiftTotals[(int)company, shiftType] == null)
                {
                    continue;
                }

                foreach (var pair in emp.ShiftTotals[(int)company, shiftType].Values)
                {
                    foreach (var weekEntry in pair)
                    {
                        if (weekEntry.Key != weekNumber)
                        {
                            continue;
                        }

                        foreach (Shift shift in weekEntry.Value)
                        {
                            if (shift.CompanyName != company || !shift.IsValid(emp))
                            {
                                continue;
                            }

                            if (shift.JobType == Jobs.VACATION || shift.JobType == Jobs.HOLIDAY)
                            {
                                continue;
                            }

                            // Blended overtime only uses regular hours and regular earnings.
                            // Do not include minimum guarantee hours, summer guarantee hours, vacation, holiday,
                            // bonus dollars, per diem, or any other non-regular compensation.
                            float shiftRegularHours = shift.ShiftTime;
                            if (shiftRegularHours < 0.001f)
                            {
                                continue;
                            }

                            float shiftRegularEarnings = 0f;
                            if (shift.DollarAmount > 0.001f)
                            {
                                shiftRegularEarnings = shift.DollarAmount;
                            }
                            else if (shift.PayRate > 0f)
                            {
                                shiftRegularEarnings = shiftRegularHours * shift.PayRate.Value;
                            }
                            else
                            {
                                Log("Cannot calculate blended regular rate for " + emp.Name + " because a regular shift has hours but no payrate or dollar amount.", true);
                                continue;
                            }

                            regularHours += shiftRegularHours;
                            regularEarnings += shiftRegularEarnings;
                        }
                    }
                }
            }

            if (regularHours < 0.001f)
            {
                Log("Cannot calculate blended regular rate for " + emp.Name + " because regular hours are zero.", true);
                return 0f;
            }

            return (float)Math.Round(regularEarnings / regularHours, 2);
        }

        private static void AddWfnHours3Row(List<Dictionary<string, string>> rows, Employee emp, Company company, string batchId, Shift shift, float hours, string code)
        {
            if (hours < 0.001f)
            {
                return;
            }
            Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
            SetWfnTempDepartment(row, shift);
            row["Hours 3 Amount"] = FormatWfnNumber(hours);
            row["Hours 3 Code"] = code;
            if (shift.PayRate > 0f)
            {
                row["Temp Rate"] = FormatWfnNumber(shift.PayRate.Value);
            }
            SetWfnFlsaWorkweek(row, shift);
            rows.Add(row);
        }

        private static void AddWfnEarnings3Row(List<Dictionary<string, string>> rows, Employee emp, Company company, string batchId, Shift shift, float dollars, string code)
        {
            if (dollars < 0.001f)
            {
                return;
            }
            Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
            SetWfnTempDepartment(row, shift);
            row["Earnings 3 Amount"] = FormatWfnNumber(dollars);
            row["Earnings 3 Code"] = code;
            SetWfnFlsaWorkweek(row, shift);
            rows.Add(row);
        }

        private static void AddWfnAdjustmentRow(List<Dictionary<string, string>> rows, Employee emp, Company company, string batchId, Shift shift, float dollars, string code)
        {
            if (dollars < 0.001f)
            {
                Log("Why is dollars < 0.001f?", true);
                return;
            }

            Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
            SetWfnTempDepartment(row, shift);
            row["Adjust Ded Amount"] = FormatWfnNumber(dollars);
            row["Adjust Ded Code"] = code;
            rows.Add(row);
        }

        private static void AddWfnSalaryRows(List<Dictionary<string, string>> rows, Employee emp, Company company, string batchId)
        {
            if (emp.AnnualSalaryAmount <= 0.001f)
            {
                Log("Cannot add WFN salary rows for salaried employee " + emp.Name + " (" + emp.IdNumber + ") because AnnualSalaryAmount is missing.", true);
                return;
            }

            float weeklySalaryAmount = RoundWfnMoneyUp((emp.AnnualSalaryAmount / 26f) / 2f);
            for (int weekNumber = 1; weekNumber <= 2; weekNumber++)
            {
                Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
                row["Reg Hours"] = "40";
                row["Reg Earnings"] = FormatWfnNumber(weeklySalaryAmount);
                row["FLSA Workweek"] = weekNumber.ToString();
                rows.Add(row);
            }
        }

        /// <summary>
        /// Writes all manual entries for the given company to the end of the WFN rows.
        /// Most manual entries never get an FLSA Workweek; regular hours are the exception.
        /// </summary>
        private static void AddManualEntryWfnRows(List<Dictionary<string, string>> rows, Company company, string batchId)
        {
            foreach (ManualEntry entry in ManualEntriesTracker.GetInstance().GetEntries())
            {
                if (entry.Company != company || entry.Employee == null)
                {
                    continue;
                }

                Employee emp = entry.Employee;
                if (string.IsNullOrWhiteSpace(emp.SocialSecurityNumber))
                {
                    if (!emp.WasReportedForPartialEntry)
                    {
                        emp.WasReportedForPartialEntry = true;
                        Log(emp.Name + " (" + emp.IdNumber + ") is not getting paid because they do not have a social security number in workbright.", true);
                    }
                    continue;
                }

                float vacationHours = entry.VacationHours + entry.RoundUpVacationHours;
                if (vacationHours > 0.001f)
                {
                    Dictionary<string, string> row = MakeManualEntryWfnRow(emp, company, batchId, Jobs.VACATION);
                    row["Hours 3 Code"] = "VAC";
                    row["Hours 3 Amount"] = FormatWfnNumber(vacationHours);
                    if (entry.VacationPayRate > 0.001f)
                    {
                        row["Temp Rate"] = FormatWfnNumber(entry.VacationPayRate);
                    }
                    rows.Add(row);
                }

                if (entry.RegularHours > 0.001f && entry.Jobtype.HasValue && entry.WeekNumber >= 1 && entry.WeekNumber <= 2)
                {
                    Dictionary<string, string> row = MakeManualEntryWfnRow(emp, company, batchId, entry.GetResolvedJobType(emp));
                    row["Reg Hours"] = FormatWfnNumber(entry.RegularHours);
                    row["FLSA Workweek"] = entry.WeekNumber.ToString();
                    if (entry.JobPayRate > 0.001f)
                    {
                        row["Temp Rate"] = FormatWfnNumber(entry.JobPayRate);
                    }
                    rows.Add(row);
                }

                if (entry.MgHours > 0.001f && entry.Jobtype.HasValue)
                {
                    Dictionary<string, string> row = MakeManualEntryWfnRow(emp, company, batchId, entry.GetResolvedJobType(emp));
                    row["Hours 3 Code"] = "MNG";
                    row["Hours 3 Amount"] = FormatWfnNumber(entry.MgHours);
                    if (entry.JobPayRate > 0.001f)
                    {
                        row["Temp Rate"] = FormatWfnNumber(entry.JobPayRate);
                    }
                    rows.Add(row);
                }

                if (entry.BonusDollars > 0.001f)
                {
                    Dictionary<string, string> row = MakeManualEntryWfnRow(emp, company, batchId, entry.Jobtype.HasValue ? entry.GetResolvedJobType(emp) : Jobs.ADMIN);
                    row["Earnings 3 Code"] = "BON";
                    row["Earnings 3 Amount"] = FormatWfnNumber(entry.BonusDollars);
                    rows.Add(row);
                }

                if (entry.BackpayHours > 0.001f && entry.Jobtype.HasValue)
                {
                    Dictionary<string, string> row = MakeManualEntryWfnRow(emp, company, batchId, entry.GetResolvedJobType(emp));
                    row["Hours 3 Code"] = "BPY";
                    row["Hours 3 Amount"] = FormatWfnNumber(entry.BackpayHours);
                    if (entry.JobPayRate > 0.001f)
                    {
                        row["Temp Rate"] = FormatWfnNumber(entry.JobPayRate);
                    }
                    rows.Add(row);
                }

                if (entry.Expense > 0.001f)
                {
                    Dictionary<string, string> row = MakeManualEntryWfnRow(emp, company, batchId, entry.Jobtype.HasValue ? entry.GetResolvedJobType(emp) : Jobs.ADMIN);
                    row["Adjust Ded Code"] = "EXP";
                    row["Adjust Ded Amount"] = FormatWfnNumber(entry.Expense);
                    rows.Add(row);
                }
            }
        }

        private static Dictionary<string, string> MakeManualEntryWfnRow(Employee emp, Company company, string batchId, Jobs jobType)
        {
            Dictionary<string, string> row = MakeBaseWfnRow(emp, company, batchId);
            string departmentCode = Shift.GetDepartmentCode(jobType);
            if (string.IsNullOrWhiteSpace(departmentCode))
            {
                Log("No department code found for manual entry job type " + jobType + " for " + emp.Name + ".", true);
            }
            else
            {
                row["Temp Dept"] = departmentCode.PadLeft(6, '0');
            }
            //FLSA Workweek is intentionally left blank for manual entries
            return row;
        }

        private static float RoundWfnMoneyUp(float value)
        {
            float rounded = (float)(Math.Ceiling(value * 100f) / 100f);
            if (rounded != value)
            {
                Log("salary rounded up from " + value + " to " + rounded);
            }
            return rounded;
        }

        private static string FormatWfnNumber(float value)
        {
            return Math.Round(value, 2).ToString("0.##");
        }

        private static void WriteCsv(string path, List<string> headers, List<Dictionary<string, string>> rows)
        {
            using StreamWriter writer = new(path, false, new System.Text.UTF8Encoding(false));
            writer.WriteLine(string.Join(",", headers.Select(EscapeCsv)));
            foreach (Dictionary<string, string> row in rows)
            {
                writer.WriteLine(string.Join(",", headers.Select(header => EscapeCsv(row.GetValueOrDefault(header, "")))));
            }
        }

        private static string EscapeCsv(string? value)
        {
            value ??= "";
            if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        //private static void WriteCsv(string path, List<string> headers, List<Dictionary<string, string>> rows)
        //{
        //    using StreamWriter writer = new(path, false, new UTF8Encoding(true));
        //    writer.WriteLine(string.Join(",", headers.Select(EscapeCsv)));
        //    foreach (var row in rows)
        //    {
        //        writer.WriteLine(string.Join(",", headers.Select(header => EscapeCsv(row.GetValueOrDefault(header, "")))));
        //    }
        //}

        //private static string EscapeCsv(string value)
        //{
        //    value ??= "";
        //    return "\"" + value.Replace("\"", "\"\"") + "\"";
        //}

        private static string GetImportField(ImportedEmployee importedEmployee, string fieldName)
        {
            return importedEmployee.ImportFields.TryGetValue(fieldName, out object? value) && value != null ? value.ToString() ?? "" : "";
        }

        private static string GetEmployeeNumberAsSixDigits(int employeeNumber)
        {
            return employeeNumber.ToString("D6");
        }

        private static string GetAdpFederalFilingStatus(string isolvedStatus)
        {
            return isolvedStatus switch
            {
                "FDS2" => "D",
                "FDH" => "H",
                "FDM2" => "J",
                _ => ""
            };
        }

        private static string GetAdpStateMaritalStatus(string isolvedStatus)
        {
            return isolvedStatus switch
            {
                "NDS" => "S",
                "NDH" => "H",
                "NDM" => "M",
                _ => ""
            };
        }

        private static string GetAdpHomeDepartment(string jobString)
        {
            if (StringSearch(jobString, "driver") || StringSearch(jobString, "drvr"))
            {
                return "000001";
            }
            if (StringSearch(jobString, "aid") || StringSearch(jobString, "para"))
            {
                return "000025";
            }
            if (StringSearch(jobString, "mech"))
            {
                return "000007";
            }
            Log("Problem determining ADP home department for job string: " + jobString, true);
            return "";
        }

        private static string GetAdpStateCode(string state)
        {
            string result = state switch
            {
                "ND" => "38",
                "MN" => "27",
                _ => ""
            };
            if (result == "")
            {
                Log("Problem determining ADP state code for state: " + state, true);
            }
            return result;
        }

        private static string GetAdpJobTitle(ImportedEmployee importedEmployee)
        {
            string job = GetImportField(importedEmployee, "Job");
            if (int.TryParse(job, out int jobNumber) && Enum.IsDefined(typeof(Jobs), jobNumber)
                && JobTitleMapper.TryMapJobToImportCode((Jobs)jobNumber, out string codeFromEnum)
                && !JobTitleMapper.IsNeverImported(codeFromEnum))
            {
                return codeFromEnum;
            }
            if (job == "1" || StringSearch(job, "Driver") || StringSearch(job, "drvr"))
            {
                return JobTitleMapper.DriverDailySchool;
            }
            if (job == "25" || StringSearch(job, "aid") || StringSearch(job, "para"))
            {
                return JobTitleMapper.Para;
            }
            if (job == "7" || StringSearch(job, "mech") || StringSearch(job, "000007"))
            {
                return JobTitleMapper.Mechanic;
            }
            if (StringSearch(job, "admin"))
            {
                return JobTitleMapper.Admin;
            }
            if (StringSearch(job, "wash") || StringSearch(job, "wshby"))
            {
                return JobTitleMapper.WashBay;
            }
            if (StringSearch(job, "body") || StringSearch(job, "bdyshp"))
            {
                return JobTitleMapper.BodyShop;
            }
            if (StringSearch(job, "non") && StringSearch(job, "cdl"))
            {
                return JobTitleMapper.NonCdlDriver;
            }
            if (JobTitleMapper.IsNeverImported(job))
            {
                Log("Refusing to import blocked job title '" + job + "' for "
                    + GetImportField(importedEmployee, "FirstName") + " "
                    + GetImportField(importedEmployee, "LastName") + "; defaulting to "
                    + JobTitleMapper.DriverDailySchool + ".", true);
                return JobTitleMapper.DriverDailySchool;
            }
            Log("Couldn't find a job title for this employee, defaulting to job title of " + job + " for " + GetImportField(importedEmployee, "FirstName") + " " + GetImportField(importedEmployee, "LastName") + ", job field: " + job, true);
            return job;
        }

        private static bool PositionIsFullTime(ImportedEmployee importedEmployee)
        {
            string JobTitle = GetAdpJobTitle(importedEmployee);
            return JobTitle != "PARA" && JobTitle != "DRDLYSC";
        }

        // Isolved ImportFields rate keys -> ADP_NewHireImport.csv rate column headers
        private static readonly Dictionary<string, string> IsolvedRateFieldToAdpHeader = new()
        {
            ["Rate_Admin"] = "Rate - Admin",
            ["Rate_AidDlySchool"] = "Rate - Para School",
            ["Rate_AidDlyChrter"] = "Rate - Para Charter",
            ["Rate_Body Shop"] = "Rate - Body Shop",
            ["Rate_Cleaning"] = "Rate - Cleaning",
            ["Rate_DrvrDlySchool"] = "Rate - Driver School",
            ["Rate_DrvrSchoolChrtr"] = "Rate - Driver Charter",
            ["Rate_Mechanic"] = "Rate - Mechanic",
            ["Rate_Wash Bay"] = "Rate - Wash Bay"
        };

        private static bool HasAdpRateImportFields(ImportedEmployee importedEmployee)
        {
            foreach (string isolvedField in IsolvedRateFieldToAdpHeader.Keys)
            {
                if (!string.IsNullOrWhiteSpace(GetImportField(importedEmployee, isolvedField)))
                {
                    return true;
                }
            }
            return false;
        }

        private static void AddAdpRateFields(Dictionary<string, string> adpFields, ImportedEmployee importedEmployee)
        {
            foreach (var mapping in IsolvedRateFieldToAdpHeader)
            {
                string rateValue = GetImportField(importedEmployee, mapping.Key);
                if (!string.IsNullOrWhiteSpace(rateValue))
                {
                    adpFields[mapping.Value] = rateValue;
                }
            }
        }

        private static string GetDdAccountField(Dictionary<string, object> account, string fieldName)
        {
            return account.TryGetValue(fieldName, out object? value) && value != null ? value.ToString()?.Trim() ?? "" : "";
        }

        /// <summary>
        /// Builds ADP bank-deposit columns for one DD account.
        /// Returns false (and logs) when the account must be discarded (e.g. routing number is not 9 digits).
        /// </summary>
        private static bool TryAddAdpBankDepositFields(Dictionary<string, string> adpFields, Dictionary<string, object> account, int positionNumber, Employee employee)
        {
            string routingNumber = GetDdAccountField(account, "RoutingNumber");
            string routingDigits = new string(routingNumber.Where(char.IsDigit).ToArray());
            if (routingDigits.Length != 9)
            {
                Log("Discarding DD account #" + positionNumber + " for " + employee.Name + " (" + employee.IdNumber + "): routing number is not 9 digits (" + routingNumber + ").", true);
                return false;
            }

            string accountNumber = GetDdAccountField(account, "AccountNumber");
            if (string.IsNullOrWhiteSpace(accountNumber))
            {
                Log("Discarding DD account #" + positionNumber + " for " + employee.Name + " (" + employee.IdNumber + "): missing account number.", true);
                return false;
            }

            string amount = GetDdAccountField(account, "Amount");
            string percent = GetDdAccountField(account, "Percent");
            string sequence = GetDdAccountField(account, "Sequence");
            bool isFullDeposit = sequence == "0"
                || (string.IsNullOrWhiteSpace(amount) && string.IsNullOrWhiteSpace(percent));

            string percentNet = "";
            string deductionAmount = "";
            string partialNet = "";
            string fullDepositFlag = "N";

            if (isFullDeposit)
            {
                fullDepositFlag = "Y";
            }
            else if (!string.IsNullOrWhiteSpace(percent) && float.TryParse(percent, out float percentValue))
            {
                if (percentValue >= 100f)
                {
                    fullDepositFlag = "Y";
                }
                else
                {
                    percentNet = percent;
                }
            }
            else if (!string.IsNullOrWhiteSpace(amount))
            {
                deductionAmount = amount;
            }

            //adpFields["Bank Deposit Position Number"] = positionNumber.ToString();
            adpFields["Bank Deposit Deduction Code"] = "CK" + positionNumber;
            adpFields["Bank Deposit Routing Number"] = routingDigits;
            adpFields["Bank Deposit Account Number"] = accountNumber;
            //adpFields["Bank Deposit Partial Net"] = partialNet;
            adpFields["Bank Full Deposit Flag"] = fullDepositFlag;
            adpFields["Bank Deposit Deduction Amount"] = deductionAmount;
            adpFields["Bank Deposit Prenote Code"] = "O";
            adpFields["Bank Deposit Prenote Date"] = DateTime.Today.ToShortDateString();
            adpFields["Bank Deposit Percent Net"] = percentNet;
            return true;
        }

        private static List<Dictionary<string, object>> GetValidDdAccountsForAdp(ImportedEmployee importedEmployee)
        {
            List<Dictionary<string, object>> accounts = new();
            foreach (Dictionary<string, object> account in importedEmployee.DDAccounts)
            {
                if (!string.IsNullOrWhiteSpace(GetDdAccountField(account, "AccountNumber"))
                    || !string.IsNullOrWhiteSpace(GetDdAccountField(account, "RoutingNumber")))
                {
                    accounts.Add(account);
                }
            }
            return accounts;
        }
        
        private static readonly List<string> AdpNewHireImportHeaders = new()
        {
            "Associate ID",
            "Position ID",
            "Change Effective On",
            "Is Paid By WFN",
            //"Position Uses Time",
            "Tax ID Type",
            "Tax ID Number",
            "First Name",
            "Middle Name",
            "Last Name",
            "Gender",
            "Hire Date",
            "Position Start Date",
            "Rate Type",
            "Rate 1 Amount",
            "Pay Frequency Code",
            "Employee Status",
            "Termination Date",
            "Termination Reason",
            "Rehire Date",
            "Rehire Reason",
            "Standard Hours",
            "Worker Category or Employee Type",
            "Home Department",
            "Birth Date",
            "Federal Tax Form Year",
            "Federal Tax Filing Status",
            "Federal Tax Multiple Jobs",
            "Federal Tax Dependents Amount",
            "Federal Tax Other Income Amount",
            "Federal Tax Deductions Amount",
            "Federal Tax Additional Amount",
            "Federal Tax Exemptions",
            "Worked State Tax Code",
            //"State Marital Status",
            "SUI/SDI Tax Jurisdiction Code",
            "Assign Onboarding Experience",
            "Job Title",
            "Address 1 Line 1",
            "Address 1 Line 2",
            "Address 1 City",
            "Address 1 State Postal Code",
            "Address 1 Zip Code",
            "FLSA OVERTIME",
            "Mobile Phone Number", //missing
            "Personal Email Address", //missing
            "Bank Deposit Deduction Code", //mandatory for each DD account, entry is CK1 for first account, CK2 for second, etc.
            "Bank Deposit Routing Number", //mandatory for each DD account, if routing number is not 9 digits, discard entire account and log an error.
            "Bank Deposit Account Number", //mandatory for each DD account
            "Bank Full Deposit Flag", //Y if the full net pay is to be deposited into this account, otherwise "N"
            "Bank Deposit Deduction Amount", //number if a specific amount is to be deposited into this account, otherwise leave blank
            "Bank Deposit Prenote Code", //always "O"
            "Bank Deposit Prenote Date", //always today's date
            "Bank Deposit Percent Net", //number if a percentage of net pay is to be deposited into this account, otherwise leave blank (special note: if this value is >= 100, leave this blank and set the Bank Full Deposit Flag to "Y" instead)


            "Rate - Admin",
            "Rate - Para School",
            "Rate - Para Charter",
            "Rate - Body Shop",
            "Rate - Cleaning",
            "Rate - Driver School",
            "Rate - Driver Charter",
            "Rate - Mechanic",
            "Rate - Wash Bay"
            //"Will Worker Complete Form I-9"
        };

        private object[,] BuildAdpNewHireImportMatrix()
        {
            List<Dictionary<string, string>> dataRows = new();

            foreach (var employeeEntry in ImportEmployees)
            {
                Employee employee = EmployeeDictionary[employeeEntry.Key];
                ImportedEmployee importedEmployee = employeeEntry.Value;

                bool isRaiseUpdate = employee.WasAlreadyInPayroll && employee.NeedsUpdateInPayroll && HasAdpRateImportFields(importedEmployee);
                if (employee.WasAlreadyInPayroll && !isRaiseUpdate)
                {
                    continue;
                }
                if (!isRaiseUpdate && employee.Shifts.Count == 0 && !employee.HadManualEntry)
                {
                    continue;
                }
                if (importedEmployee.ImportFields.Count == 0 || string.IsNullOrWhiteSpace(employee.SocialSecurityNumber))
                {
                    continue;
                }

                string sixDigitEmployeeNumber = GetEmployeeNumberAsSixDigits(employee.IdNumber);
                string hireDate = GetImportField(importedEmployee, "HireDate");
                string changeEffectiveOn = isRaiseUpdate ? FirstDayWeek2.AddDays(-7).ToShortDateString() : hireDate;
                string state = GetImportField(importedEmployee, "State");
                if (string.IsNullOrWhiteSpace(state))
                {
                    state = "ND";
                }

                string mobilePhone = GetImportField(importedEmployee, "HomePhone");
                if (string.IsNullOrWhiteSpace(mobilePhone))
                {
                    mobilePhone = employee.PhoneNumber ?? "";
                }

                Dictionary<string, string> adpFields;
                if (isRaiseUpdate)
                {
                    // Existing employees: only identity + rate columns (mirrors isolved rate-only import rows)
                    adpFields = new()
                    {
                        ["Position ID"] = "MMF" + sixDigitEmployeeNumber,
                        ["Change Effective On"] = changeEffectiveOn,
                        ["Tax ID Type"] = "SSN",
                        ["Tax ID Number"] = GetImportField(importedEmployee, "SSN")
                    };
                }
                else
                {
                    adpFields = new()
                    {
                        ["Associate ID"] = "",
                        ["Position ID"] = "MMF" + sixDigitEmployeeNumber,
                        ["Change Effective On"] = changeEffectiveOn,
                        ["Is Paid By WFN"] = "Y",
                        //["Position Uses Time"] = "N",
                        ["Tax ID Type"] = "SSN",
                        ["Tax ID Number"] = GetImportField(importedEmployee, "SSN"),
                        ["First Name"] = GetImportField(importedEmployee, "FirstName"),
                        ["Middle Name"] = GetImportField(importedEmployee, "MiddleName"),
                        ["Last Name"] = GetImportField(importedEmployee, "LastName"),
                        ["Hire Date"] = hireDate,
                        ["Position Start Date"] = hireDate,
                        ["Gender"] = employee.IsMale ? "M" : "F",
                        ["Rate Type"] = "",
                        ["Rate 1 Amount"] = "",
                        ["Pay Frequency Code"] = "B",
                        //["Employee Status"] = "A",
                        ["Standard Hours"] = "",
                        ["Worker Category or Employee Type"] = PositionIsFullTime(importedEmployee) ? "F" : "P",
                        ["Home Department"] = GetAdpHomeDepartment(GetImportField(importedEmployee, "Organization")),
                        ["Birth Date"] = GetImportField(importedEmployee, "BirthDate"),
                        ["Federal Tax Form Year"] = DateTime.Today.Year.ToString(),
                        ["Federal Tax Filing Status"] = GetAdpFederalFilingStatus(GetImportField(importedEmployee, "FedFilingStatus")),
                        ["Federal Tax Multiple Jobs"] = "N",
                        ["Federal Tax Dependents Amount"] = GetImportField(importedEmployee, "FedDependentsAmt"),
                        ["Federal Tax Other Income Amount"] = "0",
                        ["Federal Tax Deductions Amount"] = GetImportField(importedEmployee, "FedDeductions"),
                        ["Federal Tax Additional Amount"] = GetImportField(importedEmployee, "FedAddlAmount"),
                        ["Federal Tax Exemptions"] = GetImportField(importedEmployee, "FedExemptions"),
                        ["Worked State Tax Code"] = "ND",
                        //["State Marital Status"] = GetAdpStateMaritalStatus(GetImportField(importedEmployee, "StateFilingStatus")),
                        ["SUI/SDI Tax Jurisdiction Code"] = "92",
                        ["Assign Onboarding Experience"] = "",
                        ["Job Title"] = GetAdpJobTitle(importedEmployee),
                        ["Personal Email Address"] = GetImportField(importedEmployee, "SelfServiceEmail"),
                        ["Address 1 Line 1"] = GetImportField(importedEmployee, "Address1"),
                        ["Address 1 Line 2"] = GetImportField(importedEmployee, "Address2"),
                        ["Address 1 City"] = GetImportField(importedEmployee, "City"),
                        ["Address 1 State Postal Code"] = state,
                        ["Address 1 Zip Code"] = GetImportField(importedEmployee, "ZipCode"),
                        ["FLSA OVERTIME"] = "Y",
                        ["Mobile Phone Number"] = mobilePhone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", ""),
                        //["Will Worker Complete Form I-9"] = "Y"
                    };
                }

                AddAdpRateFields(adpFields, importedEmployee);

                if (isRaiseUpdate)
                {
                    dataRows.Add(adpFields);
                    continue;
                }

                List<Dictionary<string, object>> ddAccounts = GetValidDdAccountsForAdp(importedEmployee);
                if (ddAccounts.Count == 0)
                {
                    dataRows.Add(adpFields);
                    continue;
                }

                int positionNumber = 1;
                bool wroteAnyBankRow = false;
                foreach (Dictionary<string, object> account in ddAccounts)
                {
                    Dictionary<string, string> rowFields = new(adpFields);
                    if (TryAddAdpBankDepositFields(rowFields, account, positionNumber, employee))
                    {
                        dataRows.Add(rowFields);
                        wroteAnyBankRow = true;
                        positionNumber++;
                    }
                }

                if (!wroteAnyBankRow)
                {
                    dataRows.Add(adpFields);
                }
            }

            AddEmploymentStatusImportRows(dataRows);

            object[,] adpMatrix = new string[dataRows.Count + 1, AdpNewHireImportHeaders.Count];
            for (int columnNumber = 0; columnNumber < AdpNewHireImportHeaders.Count; columnNumber++)
            {
                adpMatrix[0, columnNumber] = AdpNewHireImportHeaders[columnNumber];
            }

            for (int rowNumber = 0; rowNumber < dataRows.Count; rowNumber++)
            {
                Dictionary<string, string> adpFields = dataRows[rowNumber];
                for (int columnNumber = 0; columnNumber < AdpNewHireImportHeaders.Count; columnNumber++)
                {
                    string header = AdpNewHireImportHeaders[columnNumber];
                    if (adpFields.TryGetValue(header, out string? value))
                    {
                        adpMatrix[rowNumber + 1, columnNumber] = value;
                    }
                }
            }

            return adpMatrix;
        }

        private void AddEmploymentStatusImportRows(List<Dictionary<string, string>> dataRows)
        {
            string changeEffectiveOn = FirstDayWeek2.AddDays(-7).ToShortDateString();
            DateTime payDate = FirstDayWeek2.AddDays(12);

            foreach (Employee employee in EmployeeDictionary.Values)
            {
                if (string.IsNullOrWhiteSpace(employee.SocialSecurityNumber))
                {
                    continue;
                }

                string taxId = employee.SocialSecurityNumber;
                string employeeNumber = GetEmployeeNumberAsSixDigits(employee.IdNumber);
                if (PayrollHistory.PartTimeEmployeesNeedingFullTimeStatus.Contains(employee.IdNumber))
                {
                    dataRows.Add(new()
                    {
                        ["Position ID"] = "MMF" + employeeNumber,
                        ["Change Effective On"] = changeEffectiveOn,
                        ["Tax ID Type"] = "SSN",
                        ["Tax ID Number"] = taxId,
                        ["Worker Category or Employee Type"] = "F"
                    });
                }

                IEnumerable<Company> companies = employee.ActiveCompanies?.Any() == true
                    ? employee.ActiveCompanies
                    : new[] { employee.PrimaryCompany };
                if (PayrollHistory.EmployeesNeedingTermination.Contains(employee.IdNumber))
                {
                    string terminationDate = PayrollHistory.GetTerminationDate(employee.IdNumber, payDate).ToShortDateString();
                    foreach (Company company in companies)
                    {
                        dataRows.Add(new()
                        {
                            ["Position ID"] = GetPositionId(company, employeeNumber),
                            ["Change Effective On"] = changeEffectiveOn,
                            ["Tax ID Type"] = "SSN",
                            ["Tax ID Number"] = taxId,
                            ["Employee Status"] = "T",
                            ["Termination Date"] = terminationDate,
                            ["Termination Reason"] = "N"
                        });
                    }
                }

                if (PayrollHistory.EmployeesNeedingRehire.Contains(employee.IdNumber))
                {
                    foreach (Company company in companies)
                    {
                        dataRows.Add(new()
                        {
                            ["Position ID"] = GetPositionId(company, employeeNumber),
                            ["Change Effective On"] = changeEffectiveOn,
                            ["Tax ID Type"] = "SSN",
                            ["Tax ID Number"] = taxId,
                            ["Employee Status"] = "A",
                            ["Rehire Date"] = changeEffectiveOn,
                            ["Rehire Reason"] = "CURR"
                        });
                    }
                }

                if (!string.IsNullOrWhiteSpace(employee.RecommendedJobTitleCode)
                    && !JobTitleMapper.IsNeverImported(employee.RecommendedJobTitleCode))
                {
                    dataRows.Add(new()
                    {
                        ["Position ID"] = "MMF" + employeeNumber,
                        ["Change Effective On"] = changeEffectiveOn,
                        ["Tax ID Type"] = "SSN",
                        ["Tax ID Number"] = taxId,
                        ["Job Title"] = employee.RecommendedJobTitleCode
                    });
                }
            }
        }

        private static string GetPositionId(Company company, string employeeNumber) =>
            (company == Company.VALLEY_BUS_COACHES ? "MKZ" : "MMF") + employeeNumber;

        public void WritePayrollHistory()
        {
            if (IsPrimaryPayrollRun)
            {
                PayrollHistory.WriteCurrentPayPeriod(FirstDayWeek2.AddDays(12), EmployeeDictionary.Values);
            }
        }

        public void WriteEmployeeImports()
        {
            object[,] employeeMatrix = new string[ImportEmployees.Count + 1, 52];
            object[,] raiseMatrix = new string[ImportEmployees.Count + 1, 52];
            object[,] directDepositMatrix = new string[ImportEmployees.Count + 1, 52];

            for (int columnNumber = 0; columnNumber < ImportedEmployee.EmployeeImportHeaders.Count; columnNumber++)
            {
                employeeMatrix[0, columnNumber] = ImportedEmployee.EmployeeImportHeaders[columnNumber];
            }
            for (int columnNumber = 0; columnNumber < ImportedEmployee.DDImportHeaders.Count; columnNumber++)
            {
                directDepositMatrix[0, columnNumber] = ImportedEmployee.DDImportHeaders[columnNumber];
            }

            int employeeRowNumber = 0;
            int raisesRowNumber = 0;
            int accountRowNumber = 0;
            foreach (var employeeEntry in ImportEmployees)
            {
                var employee = EmployeeDictionary[employeeEntry.Key];

                if ((!employee.HasAnyDirectDepositAccount || employee.needsDDImported) && (employee.Shifts.Count > 0 || employee.WasAlreadyInPayroll || employee.HadManualEntry))
                {
                    if (employeeEntry.Value.ImportFields.Count > 0)
                    {
                        foreach (var accountInfo in employeeEntry.Value.DDAccounts)
                        {
                            if (accountInfo.Count > 0)
                            {
                                if (employee.WasAlreadyInPayroll)
                                {
                                    Log("Adding Direct Deposit for " + employee.Name + "(" + employee.IdNumber + ") and employee.WasAlreadyInPayroll == true");
                                }
                                for (int columnNumber = 0; columnNumber < ImportedEmployee.DDImportHeaders.Count; columnNumber++)
                                {
                                    if (accountInfo.ContainsKey(ImportedEmployee.DDImportHeaders[columnNumber]))
                                    {
                                        directDepositMatrix[accountRowNumber + 1, columnNumber] = accountInfo[ImportedEmployee.DDImportHeaders[columnNumber]];
                                    }
                                }
                                accountRowNumber++;
                            }
                        }
                    }
                }
                if (employee.Shifts.Count == 0 && !employee.WasAlreadyInPayroll && !employee.HadManualEntry)
                {
                    //don't update employees who aren't currently active.
                    continue;
                }
                if (employee.WasAlreadyInPayroll && !employee.NeedsUpdateInPayroll)
                {
                    continue;
                }
                if (employeeEntry.Value.ImportFields.Count > 0 && null != employee.SocialSecurityNumber && employee.SocialSecurityNumber != "")
                {
                    if (employeeEntry.Value.ImportFields.ContainsKey("SSN"))
                    {
                        for (int columnNumber = 0; columnNumber < ImportedEmployee.EmployeeImportHeaders.Count; columnNumber++)
                        {
                            string fieldName = ImportedEmployee.EmployeeImportHeaders[columnNumber];
                            if (employee.WasAlreadyInPayroll)
                            {
                                if (!FieldsToInputEvenIfTheEmployeeWasAlreadyInPayroll.Contains(fieldName))
                                {
                                    if (!StringSearch(fieldName, "rate"))
                                    {
                                        continue;
                                    }
                                }
                            }
                            if (employeeEntry.Value.ImportFields.ContainsKey(fieldName))
                            {
                                employeeMatrix[employeeRowNumber + 1, columnNumber] = employeeEntry.Value.ImportFields[fieldName];
                            }
                        }
                        employeeRowNumber++;
                    }
                    else
                    {
                        Log("WARNING: This section shouldn't be active.", true);
                        int columnNumber = 0;
                        for (int headerNumber = 0; headerNumber < ImportedEmployee.EmployeeImportHeaders.Count; headerNumber++)
                        {
                            if (employeeEntry.Value.ImportFields.ContainsKey(ImportedEmployee.EmployeeImportHeaders[headerNumber]))
                            {
                                raiseMatrix[raisesRowNumber + 1, columnNumber] = employeeEntry.Value.ImportFields[ImportedEmployee.EmployeeImportHeaders[headerNumber]];
                                raiseMatrix[0, columnNumber] = ImportedEmployee.EmployeeImportHeaders[headerNumber];
                                columnNumber++;
                            }
                        }
                        raisesRowNumber++;
                    }
                }
            }
            Excel.Application xlApp = new();
            xlApp.DisplayAlerts = false;
            object misValue = System.Reflection.Missing.Value;
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            object[,] adpNewHireMatrix = BuildAdpNewHireImportMatrix();

            List<string> paths = new()
            {
                //{ DesktopPath() + "EmployeeImport.xlsx" },
                //{ DesktopPath() + "RaiseImport.xlsx" },
                { DesktopPath() + "DirectDepositImport.xlsx" },
                { DesktopPath() + "ADP_NewHireImport.csv" }
            };
            List<object[,]> matricis = new()
            {
                //{ employeeMatrix },
                //{raiseMatrix },
                { directDepositMatrix },
                { adpNewHireMatrix }
            };
            for (int i = 0; i < matricis.Count; i++)
            {
                Excel.Workbook? workBook = null;
                var fInfo = new FileInfo(paths[i]);
                if (fInfo.Exists)
                {
                    workBook = xlApp.Workbooks.Open(paths[i]);
                }
                if (null == workBook)
                {
                    //create new workbook
                    workBook = xlApp.Workbooks.Add(misValue);
                }
                Excel.Worksheet workSheet = workBook.Worksheets.Add(misValue);
                //Excel.Worksheet workSheet2 = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                ((Excel.Worksheet)workBook.Worksheets.get_Item(2)).Delete();
                //Marshal.ReleaseComObject(workSheet2);

                int columnCount = matricis[i].GetLength(1);
                string lastColumn = GetExcelColumnName(columnCount);
                Excel.Range range = workSheet.Range[workSheet.Range["A1"], workSheet.Range[lastColumn + matricis[i].GetLength(0)]];
                range.Value = matricis[i];

                SaveWorkBook(workBook, paths[i]);

                workBook.Close(true, misValue, misValue);
                //Marshal.ReleaseComObject(workSheet);
                //Marshal.ReleaseComObject(workBook);

                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(paths[i])
                    {
                        UseShellExecute = true
                    }
                };
                p.Start();
            }

            xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);
        }

        public void WriteBirthDates()
        {
            Excel.Application xlApp = new();
            xlApp.DisplayAlerts = false;
            object misValue = System.Reflection.Missing.Value;
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook? workBook = null;
            String path = DesktopPath() + "BirthDates.xlsx";
            var fInfo = new FileInfo(path);
            if (fInfo.Exists)
            {
                workBook = xlApp.Workbooks.Open(path);
            }
            if (null == workBook)
            {
                //create new workbook
                workBook = xlApp.Workbooks.Add(misValue);
            }
            if (workBook.Worksheets.Count < 2)
            {
                workBook.Worksheets.Add(misValue);
            }


            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(2);
            HashSet<int> activeEmployees = EmployeeDictionary
            .Where(kvp => kvp.Value.Shifts.Count > 0)
            .Select(kvp => kvp.Value.IdNumber)
            .ToHashSet();

            foreach (var kvp in EmployeeDictionary)
            {
                if (kvp.Value.IsSalaried && !kvp.Value.IsTerminated)
                {
                    activeEmployees.Add(kvp.Key);
                }
            }
            Object[,] activeEmployeesObject = new String[1, 1];
            activeEmployeesObject[0, 0] = String.Join(",", activeEmployees);


            Excel.Range range = workSheet.Range[workSheet.Range["A1"], workSheet.Range["B15"]];
            var cellData = (Object[,])range.Value2;

            DateTime payDate = FirstDayWeek2.AddDays(12);
            bool bDateWasFound = false;
            bool bShouldWriteBirthDates = true;
            for (int row = 1; row < 10; ++row)
            {
                if (TryGetStringFromCell(cellData[row, 1], out String dateString) && (DateTime.TryParse(dateString, out DateTime dateOfData) || TryGetDateFromCell(cellData[row, 1], out dateOfData)))
                {
                    if (TryGetStringFromCell(cellData[row, 2], out String cellString))
                    {
                        HashSet<int> numberSet = cellString
                            .Split(',')
                            .Select(int.Parse)
                            .ToHashSet();
                        activeEmployees.UnionWith(numberSet);
                    }

                    if (dateOfData.Year == payDate.Year && dateOfData.Month == payDate.Month)
                    {
                        if (dateOfData.Day == payDate.Day)
                        {
                            bDateWasFound = true; //this means we are just overwriting this value

                            if (Math.Abs(payDate.Day - dateOfData.Day) > 10)
                            {
                                bShouldWriteBirthDates = false;
                            }

                            break;
                        }
                    }
                }
            }

            if (bShouldWriteBirthDates)
            {

                for (int row = 1; row < 10; ++row)
                {
                    if (TryGetStringFromCell(cellData[row, 1], out String dateString) && (DateTime.TryParse(dateString, out DateTime dateOfData) || TryGetDateFromCell(cellData[row, 1], out dateOfData)))
                    {
                        if (dateOfData.AddMonths(3).CompareTo(payDate) >= 1 && (!bDateWasFound || dateOfData.Date != payDate.Date))
                        {
                            continue;
                        }
                    }
                    else if (bDateWasFound)
                    {
                        continue;
                    }

                    Object[,] now = new string[1, 1];
                    now[0, 0] = payDate.ToShortDateString();
                    workSheet.Range["A" + row].Value = now;
                    workSheet.Range["B" + row].Value = activeEmployeesObject;
                    break;
                }

                object[,] employeeMatrix = new string[Program.EmployeeDictionary.Count + 1, 52];

                Dictionary<DateTime, List<string>>[] birthDates = new Dictionary<DateTime, List<string>>[2] { new(), new() };
                foreach (var employeeEntry in EmployeeDictionary)
                {
                    var employee = employeeEntry.Value;
                    if (employee.IsTerminated || !activeEmployees.Contains(employee.IdNumber))
                    {
                        continue;
                    }

                    int index = employee.BirthDate.Month == DateTime.Now.Month ? 0 : employee.BirthDate.Month == DateTime.Now.AddMonths(1).Month ? 1 : 2;
                    if (index > 1)
                    {
                        continue;
                    }
                    if (!birthDates[index].ContainsKey(employee.BirthDate))
                    {
                        birthDates[index].Add(employee.BirthDate, new());
                    }
                    birthDates[index][employee.BirthDate].Add(employee.Name);
                }
                Dictionary<DateTime, List<string>>[] sortedBirthDates = new Dictionary<DateTime, List<string>>[2];
                sortedBirthDates[0] = birthDates[0]
                    .OrderBy(kvp => kvp.Key.Day)
                    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                sortedBirthDates[1] = birthDates[1]
                    .OrderBy(kvp => kvp.Key.Day)
                    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                int employeeRowNumber = 1;
                for (int i = 0; i < 2; i++)
                {
                    employeeMatrix[employeeRowNumber++, 0] = DateTime.Now.AddMonths(i).Month.ToString("MMMM") + ":";
                    foreach (var kvp in sortedBirthDates[i])
                    {
                        foreach (var employeeName in kvp.Value)
                        {
                            employeeMatrix[employeeRowNumber, 0] = employeeName;
                            employeeMatrix[employeeRowNumber, 1] = kvp.Key.ToShortDateString();
                            employeeRowNumber++;
                        }
                    }
                }

                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                //Marshal.ReleaseComObject(workSheet2);

                range = workSheet.Range[workSheet.Range["A1"], workSheet.Range["AZ" + employeeMatrix.GetLength(0)]];
                range.Value = employeeMatrix;

                SaveWorkBook(workBook, path);
            }

            workBook.Close(true, misValue, misValue);
            //Marshal.ReleaseComObject(workSheet);
            //Marshal.ReleaseComObject(workBook);

            //var p = new Process
            //{
            //    StartInfo = new ProcessStartInfo(path)
            //    {
            //        UseShellExecute = true
            //    }
            //};
            //p.Start();

            xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);
        }

        public void WriteOverTimeReport()
        {
            Excel.Application xlApp = new();
            xlApp.DisplayAlerts = false;
            object misValue = System.Reflection.Missing.Value;
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook? workBook = null;
            String ShortDateString = FirstDayWeek2.AddDays(12).ToShortDateString().Replace("/", "-");
            String path = "C:\\Users\\User\\valleybusllc.com\\PayrollExceptionMonitoring - PayrollMonitoring\\OvertimeReport_" + ShortDateString + ".xlsx";
            var fInfo = new FileInfo(path);
            //if (fInfo.Exists)
            //{
            //    fInfo.Delete();
            //}
            if (null == workBook)
            {
                //create new workbook
                workBook = xlApp.Workbooks.Add(misValue);
            }

            List<Employee> filteredEmployees = EmployeeDictionary.Values
            //.Where(emp => emp.OverTimeHours.Any(hours => hours >= 5))
            .Where(emp => emp.OverTimeHoursForAllCompaniesForWeek(1) + emp.OverTimeHoursForAllCompaniesForWeek(2) >= 5)
            .OrderByDescending(emp => emp.OverTimeHoursForAllCompaniesForWeek(1) + emp.OverTimeHoursForAllCompaniesForWeek(2) >= 5)
            .ToList();

            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

            workSheet.Cells[1, 1] = "Employee";
            workSheet.Cells[1, 2] = "OT Week 1";
            workSheet.Cells[1, 3] = "OT Week 2";
            workSheet.Cells[1, 4] = "OT Total";

            const int ROW_COUNT = 1500;
            object[,] matrix = new object[ROW_COUNT, 4];
            int rowCounter = 0;
            foreach (Employee employee in filteredEmployees)
            {
                matrix[rowCounter, 0] = employee.Name;
                matrix[rowCounter, 1] = Math.Round(employee.OverTimeHoursForAllCompaniesForWeek(1), 2).ToString();
                matrix[rowCounter, 2] = Math.Round(employee.OverTimeHoursForAllCompaniesForWeek(2), 2).ToString();
                matrix[rowCounter, 3] = Math.Round(employee.OverTimeHoursForAllCompaniesForWeek(1) + employee.OverTimeHoursForAllCompaniesForWeek(2), 2).ToString();
                ++rowCounter;
            }


            Excel.Range range = workSheet.Range[workSheet.Range["A2"], workSheet.Range["D" + ROW_COUNT]];
            range.Value = matrix;
            var cellData = (Object[,])range.Value2;
            SaveWorkBook(workBook, path);

            workBook.Close(true, misValue, misValue);
            xlApp.Quit();
            var p = new Process
            {
                StartInfo = new ProcessStartInfo("C:\\Users\\User\\valleybusllc.com\\PayrollExceptionMonitoring - PayrollMonitoring\\OvertimeReport_" + ShortDateString + ".xlsx")
                {
                    UseShellExecute = true
                }
            };
            p.Start();
        }

        private void WriteHeadersForTimeCardImport(Excel.Worksheet workSheet)
        {
            workSheet.Cells[1, TimeCardImportColumns.EMP_NUMBER + 1] = "Key";
            workSheet.Cells[1, TimeCardImportColumns.REGULAR_HOURS + 1] = "E_Hourly Regular_Hours";
            workSheet.Cells[1, TimeCardImportColumns.OT_HOURS + 1] = "E_Blended Overtim_Hours";
            workSheet.Cells[1, TimeCardImportColumns.MG_HOURS + 1] = "E_Min Guaran_Hours";
            workSheet.Cells[1, TimeCardImportColumns.HOLIDAY_HOURS + 1] = "E_Holiday_Hours";
            workSheet.Cells[1, TimeCardImportColumns.VACATION_HOURS + 1] = "E_Vacation_Hours";
            workSheet.Cells[1, TimeCardImportColumns.REGULAR_HOURS_WEEK + 1] = "E_Hourly Regular_WeekNumber";
            workSheet.Cells[1, TimeCardImportColumns.OT_WEEK + 1] = "E_Blended Overtim_WeekNumber";
            workSheet.Cells[1, TimeCardImportColumns.REGULAR_DOLLARS + 1] = "E_Hourly Regular_Dollars";
            workSheet.Cells[1, TimeCardImportColumns.PER_DIEM_DOLLARS_COLUMN + 1] = "E_Per Diem_Dollars";
            workSheet.Cells[1, TimeCardImportColumns.JOB_CODE + 1] = "LaborValue1";
            workSheet.Cells[1, TimeCardImportColumns.VACATION_WEEK + 1] = "E_Vacation_WeekNumber";
            workSheet.Cells[1, TimeCardImportColumns.HOLIDAY_WEEK + 1] = "E_Holiday_WeekNumber";
            workSheet.Cells[1, TimeCardImportColumns.SUMMER_BONUS_HOURS + 1] = "E_Summer Bonus_Hours";
            workSheet.Cells[1, TimeCardImportColumns.SUMMER_BONUS_WEEK + 1] = "E_Summer Bonus_WeekNumber";
            workSheet.Cells[1, TimeCardImportColumns.MG_WEEK + 1] = "E_Min Guaran_WeekNumber";
            workSheet.Cells[1, TimeCardImportColumns.BONUS_DOLLARS_COLUMN + 1] = "E_Bonus_Dollars";
            workSheet.Cells[1, TimeCardImportColumns.PAY_RATE_COLUMN + 1] = "E_*_ORRate";
        }

        private enum TimeCardImportColumns
        {
            EMP_NUMBER = 0, JOB_CODE, REGULAR_HOURS, REGULAR_HOURS_WEEK, MG_HOURS, MG_WEEK, HOLIDAY_HOURS, HOLIDAY_WEEK, VACATION_HOURS, VACATION_WEEK, PAY_RATE_COLUMN, REGULAR_DOLLARS, OT_HOURS, OT_WEEK, BONUS_DOLLARS_COLUMN, PER_DIEM_DOLLARS_COLUMN, SUMMER_BONUS_HOURS, SUMMER_BONUS_WEEK
        }

        private void CheckEmployeeNumberWithSocialSecurityNumber(Employee employee)
        {
            if (employee.SocialSecurityNumber.Equals(""))
            {
                return;
            }
            foreach (var employeeEntry in EmployeeDictionary)
            {
                if (StringSearch(employeeEntry.Value.SocialSecurityNumber, employee.SocialSecurityNumber.ToString()))
                {
                    if (employeeEntry.Key != employee.IdNumber)
                    {
                        Log("Employee number mismatch:\n" + employee.Name + ": " + employee.IdNumber.ToString() + "\n" + employeeEntry.Value.Name + ": " + employeeEntry.Key.ToString() + "\nJust make sure their number gets updated in payroll and Synovia and this payroll should be just fine.", true);
                        employeeEntry.Value.fixIdNumber(employee.IdNumber);
                    }
                    break;
                }
            }
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        private void SaveWorkBook(Excel.Workbook workBook, string filePath)
        {
            try
            {
                if (filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    workBook.SaveAs(
                        filePath,
                        Excel.XlFileFormat.xlCSV
                    );
                }
                else
                {
                    workBook.SaveAs(filePath);
                }
                //object misValue = System.Reflection.Missing.Value;
                //workBook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception e)
            {
                Log("Error saving workbook " + filePath + ". Please make sure the file is not open and run the process again.", true);
            }
        }

        private bool TryGetStringFromCell(Object cellData, out string outString)
        {
            outString = "";
            if (null != cellData)
            {
                string? str = cellData.ToString();
                if ("" != str && null != str)
                {
                    outString = str;
                    return true;
                }
            }
            return false;
        }

        private bool TryGetDateFromCell(Object cellData, out DateTime date)
        {
            date = DateTime.MinValue;
            if (null != cellData)
            {
                string? str = cellData.ToString();
                if (str != null && str != "")
                {
                    if (!double.TryParse(str, out double d))
                    {
                        return false;
                    }
                    date = DateTime.FromOADate(d);
                    return true;
                }
            }
            return false;
        }

        private bool TryGetFloatFromCell(Object cellData, out float outFloat)
        {
            outFloat = 0f;
            if (null != cellData)
            {
                string? numberString = cellData.ToString();
                if (float.TryParse(numberString, out outFloat))
                {
                    return true;
                }
            }
            return false;
        }

        private bool TryGetIntFromCell(Object cellData, out int outInt)
        {
            outInt = 0;
            if (null != cellData)
            {
                string? numberString = cellData.ToString();
                if (int.TryParse(numberString, out outInt))
                {
                    return true;
                }
            }
            return false;
        }

        private void WriteToMatrix(Employee emp, Shift shift, float time, float dollarAmount, TimeCardImportColumns timeColumn, TimeCardImportColumns weekColumn, TimeCardImportColumns? dollarColumn, ref int rowCounter, object[,] matrix)
        {
            matrix[rowCounter, (int)weekColumn] = shift.WeekNumber;
            matrix[rowCounter, (int)TimeCardImportColumns.EMP_NUMBER] = emp.IdNumber.ToString();
            matrix[rowCounter, (int)TimeCardImportColumns.JOB_CODE] = shift.GetLaborCode(false);
            if (time > 0.001f)
            {
                if (dollarAmount > 0f)
                {
                    Log("In WriteToMatrix(): time > 0f && dollarAmount > 0f", true);
                }
                matrix[rowCounter, (int)timeColumn] = Math.Round(time, 2);
            }
            if (dollarAmount > 0f)
            {
                if (shift.PayRate > 0f)
                {
                    Log("In WriteToMatrix(): dollarAmount > 0f && shift.PayRate > 0f", true);
                }
                if (null == dollarColumn)
                {
                    Log("In WriteToMatrix(): dollarAmount > 0f but null == dollarColumn.", true);
                }
                else
                {
                    matrix[rowCounter, (int)dollarColumn] = Math.Round(dollarAmount, 2);
                }
            }
            else if (shift.PayRate > 0f)
            {
                matrix[rowCounter, (int)TimeCardImportColumns.PAY_RATE_COLUMN] = Math.Round(shift.PayRate.Value, 2);
            }
            else
            {
                if (shift.AllHours(false) > 0.01)
                { //some shifts are just a per diem or bonus entry
                    Log("No Payrate or Dollar amount found for shift for " + emp.Name + ".", true);
                }
            }
            if (!shift.ExtrasWereWrittenToExport)
            {
                shift.ExtrasWereWrittenToExport = true;
                if (shift.PerDiem > 0f)
                {
                    matrix[rowCounter, (int)TimeCardImportColumns.PER_DIEM_DOLLARS_COLUMN] = Math.Round(shift.PerDiem, 2);
                }
                if (shift.BonusDollars > 0)
                {
                    matrix[rowCounter, (int)TimeCardImportColumns.BONUS_DOLLARS_COLUMN] = Math.Round(shift.BonusDollars, 2);
                }
            }
            if (dollarAmount + shift.PerDiem + shift.BonusDollars + time < 0.001f)
            {
                Log("How did shift with no time or dollar amount make it here?", true);
            }
            rowCounter++;
        }

        private Jobs GetJobTypeFromCode(int code)
        {
            switch (code)
            {
                case 1:
                case 20:
                    return Jobs.DRIVER_SCHOOL;
                case 18:
                case 21:
                    DelayedLog("Warning: Jobcode " + code + " is being used.");
                    goto case 2;
                case 2:
                    return Jobs.DRIVER_CHARTER;
                case 23:
                case 25:
                    return Jobs.AIDE_SCHOOL;
            }
            return (Jobs)code;
        }

        private int RegisterJobColumn(Dictionary<Jobs, int> columns, Jobs job, int columnNumber)
        {
            columns.Add(job, columnNumber);
            return columnNumber;
        }

        private bool TryParseTimeSpan(Object cellData, out TimeSpan timeSpan)
        {
            timeSpan = new TimeSpan();
            if (cellData != null && cellData.ToString() != null)
            {
                DateTime dt;
                if (DateTime.TryParse(cellData.ToString(), out dt))
                {
                    timeSpan = dt.TimeOfDay;
                    return true;
                }

                double oaDate;
                if (double.TryParse(cellData.ToString(), out oaDate))
                {
                    timeSpan = TimeSpan.FromHours(oaDate);
                    TimeSpan t2 = DateTime.FromOADate(oaDate).TimeOfDay;
                    DelayedLog("Check time span parsing.", true);
                    return true;
                }

                DelayedLog("Warning: Couldn't parse TimeSpan for " + cellData.ToString());
            }
            return false;
        }


        private bool CheckForExcelFileOnDesktop(string fileName, out string filePath)
        {
            filePath = DesktopPath() + fileName;
            if (!File.Exists(filePath))
            {
                Log("ERROR: Please make sure there is an excel spreadsheet on your desktop named " + fileName, true);
                return false;
            }
            return true;
        }

        private List<DateTime> GetDatesFromCoachesDateRange(string cellText)
        {
            List<DateTime> dates = new();

            if (double.TryParse(cellText, out double dateDouble))
            {
                DateTime conv = DateTime.FromOADate(dateDouble);
                dates.Add(conv);
                return dates;
            }

            int[] day = new int[2];
            int[] month = new int[2];
            int[] year = new int[2];
            cellText = cellText.Replace(" ", String.Empty);
            if (StringSearch(cellText, "-"))
            {
                string[] stringSplit = cellText.Split('-');
                if (stringSplit.Length == 2)
                {
                    for (int i = 0; i < stringSplit.Length; ++i)
                    {
                        string[] split2 = stringSplit[i].Split("/");
                        if (split2.Length > 1)
                        {
                            if (split2.Length == 3)
                            {
                                if (!int.TryParse(split2[0], out month[i]))
                                {
                                    Log("Problem getting date ranges for coaches, problem 4", true);
                                }
                                if (!int.TryParse(split2[1], out day[i]))
                                {
                                    Log("Problem getting date ranges for coaches, problem 3", true);
                                }
                                string yearString = split2[2];
                                if (yearString.Length == 2)
                                {
                                    string currentYear = DateTime.Now.Year.ToString();
                                    yearString = currentYear[..2] + yearString;
                                }
                                if (!int.TryParse(yearString, out year[i]))
                                {
                                    Log("Problem getting date ranges for coaches, problem 5", true);
                                }
                            }
                            else
                            {
                                if (i == 0)
                                {
                                    if (!int.TryParse(split2[0], out month[i]))
                                    {
                                        Log("Problem getting date ranges for coaches", true);
                                    }
                                    if (!int.TryParse(split2[1], out day[i]))
                                    {
                                        Log("Problem getting date ranges for coaches", true);
                                    }
                                }
                                else
                                {
                                    month[i] = month[0];
                                    if (!int.TryParse(split2[0], out day[i]))
                                    {
                                        Log("Problem getting date ranges for coaches", true);
                                    }
                                    string yearString = split2[1];
                                    if (yearString.Length == 2)
                                    {
                                        string currentYear = DateTime.Now.Year.ToString();
                                        yearString = currentYear[..2] + yearString;
                                    }
                                    if (!int.TryParse(yearString, out year[i]))
                                    {
                                        Log("Problem getting date ranges for coaches", true);
                                    }
                                }
                            }
                            if (i > 0)
                            {
                                if (Math.Abs(month[i] - month[0]) > 1)
                                {
                                    if (Math.Abs(month[i] - month[0]) != 11)
                                    {
                                        Log("Problem getting date ranges for coaches, problem 7", true);
                                    }
                                    if (month[0] > month[i])
                                    {
                                        year[i] = year[0] + 1;
                                    }
                                    else
                                    {
                                        Log("Problem getting date ranges for coaches", true);
                                    }
                                }
                                else
                                {
                                    year[0] = year[i];
                                }
                                if (Math.Abs(year[0] - year[i]) > 1)
                                {
                                    Log("Problem getting date ranges for coaches, problem 6", true);
                                }
                            }
                        }
                        else
                        {
                            Log("Problem getting date ranges for coaches, problem 2", true);
                        }
                    }
                }
                else
                {
                    Log("Problem getting date ranges for coaches, problem 1", true);
                }
            }

            DateTime firstDay = new DateTime(year[0], month[0], day[0]);
            dates.Add(firstDay);
            DateTime lastDay = new DateTime(year[1], month[1], day[1]);
            for (int i = 1; i < 14; ++i)
            {
                DateTime nextDay = firstDay.AddDays(i);
                if (nextDay.CompareTo(lastDay) > 0)
                {
                    break;
                }
                else
                {
                    dates.Add(nextDay);
                }
            }

            return dates;
        }

        private void CheckForManualEntry(Shift shift, Employee employee)
        {
            Dictionary<int, Jobs> employeesToCheck = new()
            {
                { 1523, Jobs.DRIVER_SCHOOL }
            };
            if (!employeesToCheck.ContainsKey(employee.IdNumber) || employeesToCheck[employee.IdNumber] != shift.JobType)
            {
                return;
            }
            if (shift.ClockIn.Seconds == 0 || shift.ClockOut.Seconds == 0)
            {
                //Log("Possible manual entry detected for " + employee.Name + " on " + shift.Date.ToShortDateString(), true);
            }
        }
    }

    public class ImportedEmployee
    {
        public static List<string> EmployeeImportHeaders = new()
            {
                "HireDate",
                "EmployeeNumber",
                "TimeClockID",
                "SSN",
                "FirstName",
                "MiddleName",
                "LastName",
                "SelfServiceEnabled",
                "SelfServiceEmail",
                "Address1",
                "Address2",
                "ZipCode",
                "City",
                "State",
                "BirthDate",
                "HomePhone",
                "I9Completed",
                "I9CompletedDate",
                "Citizenship",
                "Gender",
                "PayType",
                "Frequency",
                "NormalHours",
                "Job",
                "Organization",
                "ResidentLocation",
                "WorkLocation",
                "FedFilingStatus",
                "StateFilingStatus",
                "FedExemptions",
                "StateExemptions",
                "FedBlockTax",
                "StateBlockTax",
                "FedDependentsAmt",
                "FedAddlAmount",
                "EmploymentCategory",
                "Rate_Training",
                "Rate_AidDlySchool",
                "Rate_DrvrDlySchool",
                "Rate_DrvrSchoolChrtr",
                "Rate_AidDlyChrter",
                "Rate_Admin",
                "Rate_Wash Bay",
                "Rate_Body Shop",
                "Rate_Mechanic",
                "Rate_Cleaning",
                "OrganizationValue2"
            };
        public Dictionary<string, object> ImportFields = new();

        public static List<string> DDImportHeaders = new()
            {
                "Key",
                "Status",
                "AccountType",
                "Sequence",
                "Amount",
                "Percent",
                "RoutingNumber",
                "AccountNumber"
            };
        public bool WasOnImployeeExportSheet = false;
        public List<Dictionary<string, object>> DDAccounts = new()
            {
                new(), new()
            };

        public int LatestAccountIndex = 1;
    }
}
