using System.Data;
using System.Diagnostics;
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
        public DateTime FirstDayWeek2;
        public ExcelWorker()
        {
            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            if (PrintForm.InputDateTime("Would you like to manually enter the first day of week 2 (as opposed to auto-detecting the date)?", out DateTime dateTime))
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


            //manual override
            //FirstDayWeek2 = new DateTime(2023, 9, 24);
            //Log("FirstDayWeek2 override is active.", true);
        }

        public void Read501394()
        {
            if (!CheckForExcelFileOnDesktop("503194-01.xlsx", out string filePath))
            {
                return;
            }
            Excel.Application excelApp = new Excel.Application();
            var fInfo = new FileInfo(filePath);
            Excel.Workbook workBook = excelApp.Workbooks.Open(filePath);

            Dictionary<Jobs, int> payColumns = new();
            const int SSN_COLUMN = 2;
            const int EMP_LAST_NAME_COLUMN = 3;
            const int EMP_FIRST_NAME_COLUMN = 4;
            const int PHONE_NUMBER_COLUMN = 15;
            const int SALARY_COLUMN = 28;
            int ADMIN_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.ADMIN, 30);
            int AIDE_SCHOOL_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.AIDE_SCHOOL, 31);
            int AIDE_CHARTER_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.AIDE_CHARTER, 32);
            int BODY_SHOP_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.BODY_SHOP, 34);
            int CLEANING_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.CLEANING, 35);
            int DRIVER_SCHOOL_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.DRIVER_SCHOOL, 36);
            int DRIVER_CHARTER_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.DRIVER_CHARTER, 37);
            int MECHANIC_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.MECHANIC, 38);
            int WASH_BAY_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.WASH_BAY, 40);
            const int EMP_NUMBER_COLUMN = 41;
            int TRAINING_PAY_COLUMN = RegisterJobColumn(payColumns, Jobs.TRAINING, 42);
            const int ORGANIZATION_TAG_COLUMN = 44;
            const int YEARS_OF_SERVICE_COLUMN = 45;
            const int EMPLOYMENT_CATEGORY_COLUMN = 47;
            const int DD_ACCOUNT_1 = 48;

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
                        if (!Program.EmployeeDictionary.ContainsKey(employeeNumber))
                        {
                            string? employeeName = cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString() + " " + cellData[rowNumber, EMP_LAST_NAME_COLUMN].ToString();
                            Program.EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, employeeName));
                        }
                        Employee employee = Program.EmployeeDictionary[employeeNumber];
                        TryGetStringFromCell(cellData[rowNumber, PHONE_NUMBER_COLUMN], out employee.PhoneNumber);
                        TryGetStringFromCell(cellData[rowNumber, SSN_COLUMN], out employee.SocialSecurityNumber);
                        foreach (KeyValuePair<Jobs, int> entry in payColumns)
                        {
                            if (TryGetFloatFromCell(cellData[rowNumber, entry.Value], out float payRate))
                            {
                                employee.SetPayRate(entry.Key, Math.Max(payRate, employee.PayRates.GetValueOrDefault(entry.Key, 0f)));
                            }
                        }
                        if (TryGetFloatFromCell(cellData[rowNumber, SALARY_COLUMN], out float salary) && salary > 50)
                        {
                            employee.IsSalaried = true;
                        }
                        if (!employee.IsGrandForksEmployee)
                        {
                            if (TryGetStringFromCell(cellData[rowNumber, ORGANIZATION_TAG_COLUMN], out string tag))
                            {
                                employee.IsGrandForksEmployee = StringSearch(tag, "Grand Forks") || StringSearch(tag, "GF");
                            }
                        }
                        if (TryGetIntFromCell(cellData[rowNumber, YEARS_OF_SERVICE_COLUMN], out int yearsOfService))
                        {
                            employee.YearsOfService = Math.Max(yearsOfService, employee.YearsOfService);
                        }
                        if (employee.EmploymentCategory != "ACAFT")
                        {
                            TryGetStringFromCell(cellData[rowNumber, EMPLOYMENT_CATEGORY_COLUMN], out employee.EmploymentCategory);
                        }
                        if (!employee.HasADirectDepositAccount)
                        {
                            for (int i = 0; i < 6; i++)
                            {
                                if (TryGetStringFromCell(cellData[rowNumber, DD_ACCOUNT_1 + i], out string accountStatus))
                                {
                                    if ((i == 5 && accountStatus != "") || accountStatus == "Active")
                                    {
                                        employee.HasADirectDepositAccount = true;
                                        break;
                                    }
                                }
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

        public void PreCheckTimeSheets()
        {
            const int EMP_NUMBER_COLUMN = 2;
            const int EMP_NAME_COLUMN = 3;
            const int DAY_COLUMN = 6;
            const int PUNCH_IN_COLUMN = 8;
            const int PUNCH_OUT_COLUMN = 10;
            const int ROUNDED_TIME_COLUMN = 12;
            const int JOB_TYPE_COLUMN = 13;
            const int NOTES_COLUMN = 16;
            const int BUS_NUMBER_COLUMN = 32;

            List<int> employeeIdsToIgnore = new() { 503/*John Mc*/};

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
                range = sheet.Range[sheet.Range["A1"], sheet.Range["AG" + rows]];
                var cellData = (Object[,])range.Value2;
                rows = cellData.GetLength(0);
                for (int rowNumber = 6; rowNumber <= rows; ++rowNumber)
                {
                    if (null != cellData[rowNumber, DAY_COLUMN])
                    {
                        string? date = cellData[rowNumber, DAY_COLUMN].ToString();
                        if (date == null || date == "" || date == " ")
                        {
                            Log("date == nothing for row: " + rowNumber);
                            continue;
                        }
                        else
                        {
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

                            if (employeeIdsToIgnore.Contains(employeeNumber))
                            {
                                continue;
                            }

                            if (EmployeeDictionary.ContainsKey(employeeNumber))
                            {
                                EmployeeDictionary[employeeNumber].ShouldBeConsideredForRaises = true;
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
            if (!CheckForExcelFileOnDesktop("Employee Export.xlsx", out string filePath))
            {
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

            List<string> headers = new() { "Start Date", "Employee #", "Employee #", "SSN", "First Name", "Middle Name", "Last Name", "Email", "Street", "Apt/Suite/Unit", "Zip", "City", "State", "Birthdate", "Phone", "Date Received (Form I-9)", "Citizenship Designation (Form I-9)", "Gender", "Position", "Zip", "Filing Status (W4)", "Deductions (W4)", "Total Dependents Withholding (W4)", "Extra Withholding (W4)", "Exempt Status (W4)", "Account 1", "Account 1 - $ Specific Deposit Amount", "Account 1 - % Net Amount", "Account 1 - Account Number", "Account 1 - Deposit Instructions", "Account 1 - Routing Number", "Account 1 - Type", "Account 2", "Account 2 - $ Specific Deposit Amount", "Account 2 - % Net Amount", "Account 2 - Account Number", "Account 2 - Deposit Instructions", "Account 2 - Routing Number", "Account 2 - Type", "Account 3", "Account 3 - Account Number", "Account 3 - Routing Number", "Account 3 - Type", "Employee Groups" };

            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["B2"]].CurrentRegion;
                //Excel.Range excelRange = (Excel.Range)sheet.Range[sheet.Range["A1"], sheet.Range["P36"]];
                var cellData = (Object[,])range.Value2;
                int rows = cellData.GetLength(0);
                for (int rowNumber = 2; rowNumber <= rows; ++rowNumber)
                {
                    //Log("cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString() == " + cellData[rowNumber, EMP_FIRST_NAME_COLUMN].ToString());
                    if (TryGetIntFromCell(cellData[rowNumber, EMPLOYEE_NUMBER], out int employeeNumber))
                    {
                        bool bEmpWasAlreadyInPayroll = false;
                        ImportedEmployee importedEmployee = new();
                        importedEmployee.WasOnImployeeExportSheet = true;
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
                            Program.EmployeeDictionary[employeeNumber].WasCreatedFromEmployeeExport = true;
                        }
                        else
                        {
                            bEmpWasAlreadyInPayroll = true;
                            //Log("Employee: " + Program.EmployeeDictionary[employeeNumber].Name + " from employee export is already in payroll.", true);
                            //continue;
                        }
                        Employee employee = Program.EmployeeDictionary[employeeNumber];

                        foreach (var header in headers)
                        {
                            Object? cell = null;
                            for (int i = 0; i < headers.Count; i++)
                            {
                                if (null != cellData[1, i + 1] && header == cellData[1, i + 1].ToString())
                                {
                                    cell = cellData[rowNumber, i + 1];
                                    break;
                                }
                            }
                            if (cell != null)
                            {
                                TryGetStringFromCell(cell, out string cellString);
                                switch (header)
                                {
                                    case "Start Date":
                                        double d = double.Parse(cellString);
                                        importedEmployee.ImportFields["HireDate"] = DateTime.FromOADate(d).ToShortDateString();
                                        break;
                                    case "Employee #":
                                        break;
                                    case "SSN":
                                        importedEmployee.ImportFields["SSN"] = cellString;
                                        employee.SocialSecurityNumber = cellString;
                                        break;
                                    case "First Name":
                                        importedEmployee.ImportFields["FirstName"] = cellString;
                                        break;
                                    case "Middle Name":
                                        importedEmployee.ImportFields["MiddleName"] = cellString;
                                        break;
                                    case "Last Name":
                                        importedEmployee.ImportFields["LastName"] = cellString;
                                        break;
                                    case "Email":
                                        importedEmployee.ImportFields["SelfServiceEnabled"] = bEmpWasAlreadyInPayroll ? "N" : "Y";
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
                                        d = double.Parse(cellString);
                                        importedEmployee.ImportFields["BirthDate"] = DateTime.FromOADate(d).ToShortDateString();
                                        break;
                                    case "Phone":
                                        importedEmployee.ImportFields["HomePhone"] = cellString;
                                        employee.PhoneNumber = cellString;
                                        break;
                                    case "Date Received (Form I-9)":
                                        importedEmployee.ImportFields["I9Completed"] = cellString == "" ? "N" : "Y";
                                        if (cellString != "")
                                        {
                                            d = double.Parse(cellString);
                                            importedEmployee.ImportFields["I9CompletedDate"] = DateTime.FromOADate(d).ToShortDateString();
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
                                        importedEmployee.ImportFields["Gender"] = StringSearch(cellString, "Female") ? "F" : "M";
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
                                            importedEmployee.ImportFields["Job"] = ((int)Jobs.AIDE_SCHOOL).ToString();
                                            importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.AIDE_SCHOOL, false);
                                        }
                                        else if (StringSearch(cellString, "driver"))
                                        {
                                            importedEmployee.ImportFields["Job"] = ((int)Jobs.DRIVER_SCHOOL).ToString();
                                            importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.DRIVER_SCHOOL, false);
                                        }
                                        else if (StringSearch(cellData[rowNumber, EMPLOYEE_GROUPS].ToString(), "driver"))
                                        {
                                            importedEmployee.ImportFields["Job"] = ((int)Jobs.DRIVER_SCHOOL).ToString();
                                            importedEmployee.ImportFields["Organization"] = Shift.GetLaborCode(Jobs.DRIVER_SCHOOL, false);
                                        }
                                        else if (StringSearch(cellData[rowNumber, EMPLOYEE_GROUPS].ToString(), "para") || StringSearch(cellData[rowNumber, EMPLOYEE_GROUPS].ToString(), "aide"))
                                        {
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
                                            importedEmployee.ImportFields["FedFilingStatus"] = "FDS";
                                            importedEmployee.ImportFields["StateFilingStatus"] = "NDS";
                                        }
                                        else if (StringSearch(cellString, "household"))
                                        {
                                            importedEmployee.ImportFields["FedFilingStatus"] = "FDH";
                                            importedEmployee.ImportFields["StateFilingStatus"] = "NDH";
                                        }
                                        else
                                        {
                                            importedEmployee.ImportFields["FedFilingStatus"] = "FDM";
                                            importedEmployee.ImportFields["StateFilingStatus"] = "NDM";
                                        }
                                        break;
                                    case "Deductions (W4)":
                                        importedEmployee.ImportFields["FedExemptions"] = cellString;
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
                                        employee.HasADirectDepositAccount = true;
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
                                        importedEmployee.ImportFields["Rate_Training"] = TRAINING_RATE.ToString();
                                        employee.PayRates[Jobs.TRAINING] = TRAINING_RATE;
                                        if (StringSearch(cellString, "GF"))
                                        {
                                            employee.IsGrandForksEmployee = true;
                                            importedEmployee.ImportFields["OrganizationValue2"] = "GF";
                                        }
                                        if (StringSearch(cellString, "para") || StringSearch(cellString, "aide"))
                                        {
                                            float payRate = employee.IsGrandForksEmployee ? GrandForksDefaultRates[Jobs.AIDE_SCHOOL] : FargoDefaultRates[Jobs.AIDE_SCHOOL];
                                            importedEmployee.ImportFields["Rate_AidDlySchool"] = payRate.ToString();
                                            employee.PayRates[Jobs.AIDE_SCHOOL] = payRate;

                                            payRate = employee.IsGrandForksEmployee ? GrandForksDefaultRates[Jobs.AIDE_CHARTER] : FargoDefaultRates[Jobs.AIDE_CHARTER];
                                            importedEmployee.ImportFields["Rate_AidDlyChrter"] = payRate.ToString();
                                            employee.PayRates[Jobs.AIDE_CHARTER] = payRate;
                                        }
                                        else if (StringSearch(cellString, "driver"))
                                        {
                                            float payRate = employee.IsGrandForksEmployee ? GrandForksDefaultRates[Jobs.DRIVER_SCHOOL] : FargoDefaultRates[Jobs.DRIVER_SCHOOL];
                                            importedEmployee.ImportFields["Rate_DrvrDlySchool"] = payRate.ToString();
                                            employee.PayRates[Jobs.DRIVER_SCHOOL] = payRate;

                                            payRate = employee.IsGrandForksEmployee ? GrandForksDefaultRates[Jobs.DRIVER_CHARTER] : FargoDefaultRates[Jobs.DRIVER_CHARTER];
                                            importedEmployee.ImportFields["Rate_DrvrSchoolChrtr"] = payRate.ToString();
                                            employee.PayRates[Jobs.DRIVER_CHARTER] = payRate;
                                        }
                                        break;
                                }
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

        private HashSet<string> BusProblems = new();
        public void ReadTimeSheets()
        {
            const int EMP_NUMBER_COLUMN = 2;
            const int EMP_NAME_COLUMN = 3;
            const int DAY_COLUMN = 6;
            const int PUNCH_IN_COLUMN = 8;
            const int PUNCH_OUT_COLUMN = 10;
            const int ROUNDED_TIME_COLUMN = 12;
            const int JOB_TYPE_COLUMN = 13;
            const int NOTES_COLUMN = 16;
            const int BUS_NUMBER_COLUMN = 32;

            List<int> employeeIdsToIgnore = new() { 503/*John Mc*/};

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

            foreach (Excel.Worksheet sheet in workBook.Worksheets)
            {
                Excel.Range range = sheet.Range[sheet.Range["A6"], sheet.Range["B8"]];
                range = range.CurrentRegion;
                int rows = range.Value2.GetLength(0) + 6;
                //Excel.Range range = sheet.Range[sheet.Range["A6"]].CurrentRegion;
                range = sheet.Range[sheet.Range["A1"], sheet.Range["AG" + rows]];
                var cellData = (Object[,])range.Value2;
                rows = cellData.GetLength(0);
                for (int rowNumber = 6; rowNumber <= rows; ++rowNumber)
                {
                    if (null != cellData[rowNumber, DAY_COLUMN])
                    {
                        string? date = cellData[rowNumber, DAY_COLUMN].ToString();
                        if (date == null || date == "" || date == " ")
                        {
                            Log("date == nothing for row: " + rowNumber);
                            continue;
                        }
                        else
                        {
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

                            if (employeeIdsToIgnore.Contains(employeeNumber))
                            {
                                continue;
                            }

                            Shift shift = new(Company.VALLEY_BUS_LLC);
                            shift.ShiftTime = time;
                            double d = double.Parse(date);
                            DateTime conv = DateTime.FromOADate(d);
                            shift.Date = conv;

                            shift.WeekNumber = conv.CompareTo(FirstDayWeek2) < 0 ? 1 : 2;
                            if (conv.AddDays(7).CompareTo(FirstDayWeek2) < 0 || conv.AddDays(-7).CompareTo(FirstDayWeek2) > 0)
                            {
                                Log("ERROR: Date of shift: " + conv.ToShortDateString() + " is not within 7 days of FirstDayWeek2 ( " + FirstDayWeek2.ToShortDateString() + ")", true);
                            }

                            TryParseTimeSpan(cellData[rowNumber, PUNCH_IN_COLUMN], out shift.ClockIn);
                            TryParseTimeSpan(cellData[rowNumber, PUNCH_OUT_COLUMN], out shift.ClockOut);

                            if (!EmployeeDictionary.ContainsKey(employeeNumber))
                            {
                                string name = null == cellData[rowNumber, EMP_NAME_COLUMN] ? "" : (null == cellData[rowNumber, EMP_NAME_COLUMN].ToString() ? "" : new string((cellData[rowNumber, EMP_NAME_COLUMN].ToString())));
                                DelayedLog("In Timesheets, Employee " + employeeNumber + " (" + name + ") was not found.", true);
                                EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, name));
                            }
                            Employee employee = EmployeeDictionary[employeeNumber];

                            if (TryGetIntFromCell(cellData[rowNumber, JOB_TYPE_COLUMN], out int jobTypeInt))
                            {
                                shift.JobType = GetJobTypeFromCode(jobTypeInt);
                                if (shift.JobType == Jobs.DRIVER_SCHOOL && !employee.PayRates.ContainsKey(shift.JobType))
                                {
                                    shift.JobType = Jobs.NON_CDL_DRIVER;
                                    if (!employee.IsSalaried)
                                    {
                                        NonCdlDrivers.Add(employee);
                                    }
                                }
                            }
                            else
                            {
                                Log("Problem getting Job Code for code: " + cellData[rowNumber, JOB_TYPE_COLUMN].ToString(), true);
                            }

                            if (TryGetStringFromCell(cellData[rowNumber, NOTES_COLUMN], out shift.Notes))
                            {
                                if (StringSearch(shift.Notes, "bonus"))
                                {
                                    shift.IsABusStartingShift = true;
                                }
                            }

                            if (TryGetIntFromCell(cellData[rowNumber, BUS_NUMBER_COLUMN], out shift.BusNumber))
                            {
                                shift.IsAGrandForksShift = shift.BusNumber >= GF_MIN_BUS && shift.BusNumber <= GF_MAX_BUS;
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

                                if (jobTypeInt == 20 || jobTypeInt == 23)
                                {
                                    shift.BusNumber = Shift.WEST_FARGO_BUS_PLACE_HOLDER;
                                }
                            }

                            if ((shift.JobType == Jobs.DRIVER_SCHOOL || shift.JobType == Jobs.AIDE_SCHOOL) && !StringSearch(shift.Notes, "training"))
                            {
                                if (jobTypeInt == 20 || jobTypeInt == 23)
                                {
                                    shift.ShiftLocation = Location.WEST_FARGO;
                                }
                                else if (shift.IsAGrandForksShift || employee.IsGrandForksEmployee)
                                {
                                    shift.ShiftLocation = Location.GRAND_FORKS;
                                }
                                else
                                {
                                    shift.ShiftLocation = Location.FARGO;
                                }
                                if (shift.JobType == Jobs.DRIVER_SCHOOL)
                                {
                                    Shift.DailySchoolRouteCounter[(int)shift.ShiftLocation, conv.Day] += 1;
                                }
                            }

                            employee.Shifts.Add(shift);
                        }
                    }
                }
            }

            workBook.Close();
            excelApp.Quit();

            //Marshal.ReleaseComObject(workBook);
            //Marshal.ReleaseComObject(excelApp);
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

            if (!CheckForExcelFileOnDesktop("Coaches Payroll.xlsx", out string filePath))
            {
                return;
            }
            var lastModified = File.GetLastWriteTime(filePath);
            if (new DateTime(lastModified.Year, lastModified.Month, lastModified.Day).CompareTo(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)) < 0)
            {
                Log("Coaches Payroll is old.", true);
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
                            DelayedLog("In Coaches Payroll, Employee " + employeeNumber + " (" + name + ") was not found.", true);
                            EmployeeDictionary.Add(employeeNumber, new Employee(employeeNumber, name));
                        }
                        Employee employee = EmployeeDictionary[employeeNumber];
                        //todo: check employee name to find data entry errors

                        List<DateTime> dates = GetDatesFromCoachesDateRange(date);

                        TryGetFloatFromCell(cellData[rowNumber, DOLLARS_COLUMN], out float dollars);
                        TryGetFloatFromCell(cellData[rowNumber, BONUS_COLUMN], out float bonus);
                        TryGetFloatFromCell(cellData[rowNumber, PER_DIEM_COLUMN], out float perDiem);
                        TryGetFloatFromCell(cellData[rowNumber, HOURS_COLUMN], out float hours);
                        TryGetIntFromCell(cellData[rowNumber, BUS_NUMBER_COLUMN], out int busNumber);

                        List<Shift> shifts = new();
                        if (hours > 0/*week number doesn't really matter if there's no hours*/ &&
                            dates.Count > 1 && dates[0].CompareTo(FirstDayWeek2) != dates[^1].CompareTo(FirstDayWeek2))
                        { //multiple shifts and different weeks.


                            shifts = dates.Select(date =>
                            {
                                //TODO: test this
                                Shift shift = new(company)
                                {
                                    Date = date,
                                    WeekNumber = date.CompareTo(FirstDayWeek2) < 0 ? 1 : 2,
                                    JobType = Jobs.DRIVER_COACH,
                                    DollarAmount = dollars / dates.Count,
                                    BonusDollars = bonus / dates.Count,
                                    PerDiem = perDiem,
                                    ShiftTime = hours,
                                    BusNumber = busNumber
                                };
                                return shift;
                            }).ToList();

                            //for (int i = 0; i < 2; ++i)
                            //{
                            //    shifts.Add(new(company)
                            //    {
                            //        WeekNumber = (i == 0 ? dates[0] : dates[^1]).CompareTo(FirstDayWeek2) < 0 ? 1 : 2,
                            //        Date = i == 0 ? dates[0] : dates[^1],
                            //        JobType = Jobs.DRIVER_COACH,
                            //        DollarAmount = dollars * 0.5f,
                            //        BonusDollars = bonus * 0.5f,
                            //        PerDiem = perDiem * 0.5f,
                            //        ShiftTime = hours * 0.5f,
                            //        BusNumber = busNumber

                            //    });
                            //}
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
                                BusNumber = busNumber

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
                    if (emp != null)
                    {
                        if (emp.Shifts.Count > 0)
                        {
                            if (!emp.HasADirectDepositAccount)
                            {
                                DelayedLog("Employee: " + emp.Name + " (" + emp.IdNumber + ") has no active DD account. Phone: " + emp.PhoneNumber);
                            }
                            for (int shiftType = 0; shiftType < 3; ++shiftType)
                            {
                                if (null != emp.ShiftTotals[company, shiftType])
                                {
                                    foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                                    {
                                        foreach (Shift shift in pair.Values)
                                        {
                                            if (shift.IsValid(emp) && shift.CompanyName == (Company)company)
                                            {
                                                if (shift.ShiftTime + shift.DollarAmount + shift.BonusDollars + shift.PerDiem > 0f)
                                                {
                                                    if (shift.JobType == Jobs.VACATION)
                                                    {
                                                        WriteToMatrix(emp, shift, shift.ShiftTime, 0f, TimeCardImportColumns.VACATION_HOURS, TimeCardImportColumns.VACATION_WEEK, TimeCardImportColumns.VACATION_DOLLARS, ref rowCounter, matrix);
                                                    }
                                                    else if (shift.JobType == Jobs.HOLIDAY)
                                                    {
                                                        WriteToMatrix(emp, shift, shift.ShiftTime, 0f, TimeCardImportColumns.HOLIDAY_HOURS, TimeCardImportColumns.HOLIDAY_WEEK, TimeCardImportColumns.HOLIDAY_DOLLARS, ref rowCounter, matrix);
                                                    }
                                                    else
                                                    {
                                                        WriteToMatrix(emp, shift, shift.ShiftTime, shift.DollarAmount, TimeCardImportColumns.REGULAR_HOURS, TimeCardImportColumns.REGULAR_HOURS_WEEK, TimeCardImportColumns.REGULAR_DOLLARS, ref rowCounter, matrix);
                                                    }
                                                }
                                                if (shift.MinimumGuaranteeHours > 0f)
                                                {
                                                    WriteToMatrix(emp, shift, shift.MinimumGuaranteeHours, shift.MgDollars, TimeCardImportColumns.MG_HOURS, TimeCardImportColumns.MG_WEEK, TimeCardImportColumns.MG_DOLLARS, ref rowCounter, matrix);
                                                }
                                                if (shift.SummerGuaranteeHours > 0f)
                                                {
                                                    WriteToMatrix(emp, shift, shift.SummerGuaranteeHours, 0f, TimeCardImportColumns.SUMMER_BONUS_HOURS, TimeCardImportColumns.SUMMER_BONUS_WEEK, TimeCardImportColumns.SUMMER_BONUS_DOLLARS, ref rowCounter, matrix);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if ((Company)company == Company.VALLEY_BUS_LLC)
                            {
                                for (int weekNumber = 1; weekNumber < 3; ++weekNumber)
                                {
                                    if (emp.OverTimeHours[weekNumber] > 0f)
                                    {
                                        matrix[rowCounter, (int)TimeCardImportColumns.EMP_NUMBER] = emp.IdNumber;
                                        matrix[rowCounter, (int)TimeCardImportColumns.JOB_CODE] = "OT";
                                        matrix[rowCounter, (int)TimeCardImportColumns.OT_HOURS] = Math.Round(emp.OverTimeHours[weekNumber], 2);
                                        matrix[rowCounter, (int)TimeCardImportColumns.OT_WEEK] = weekNumber;
                                        ++rowCounter;
                                    }
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

            var employeeList = ImportEmployees.Values.ToList();
            int employeeRowNumber = 0;
            int raisesRowNumber = 0;
            int accountRowNumber = 0;
            foreach (var employeeEntry in ImportEmployees)
            {
                if (employeeEntry.Value.ImportFields.Count > 0)
                {
                    foreach (var accountInfo in employeeEntry.Value.DDAccounts)
                    {
                        if (accountInfo.Count > 0)
                        {
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
                var employee = EmployeeDictionary[employeeEntry.Key];
                if (!employee.WasCreatedFromEmployeeExport && (employee.Shifts.Count == 0/* || employeeEntry.Value.ImportFields.ContainsKey("SSN")*/))
                {
                    continue;
                }
                if (employeeEntry.Value.ImportFields.Count > 0)
                {
                    if (employeeEntry.Value.ImportFields.ContainsKey("SSN"))
                    {
                        for (int columnNumber = 0; columnNumber < ImportedEmployee.EmployeeImportHeaders.Count; columnNumber++)
                        {
                            if (employeeEntry.Value.ImportFields.ContainsKey(ImportedEmployee.EmployeeImportHeaders[columnNumber]))
                            {
                                employeeMatrix[employeeRowNumber + 1, columnNumber] = employeeEntry.Value.ImportFields[ImportedEmployee.EmployeeImportHeaders[columnNumber]];
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
            List<string> paths = new()
            {
                { DesktopPath() + "EmployeeImport.xlsx" },
                //{ DesktopPath() + "RaiseImport.xlsx" },
                { DesktopPath() + "DirectDepositImport.xlsx" }
            };
            List<object[,]> matricis = new()
            {
                {employeeMatrix },
                //{raiseMatrix },
                {directDepositMatrix }
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

                Excel.Range range = workSheet.Range[workSheet.Range["A1"], workSheet.Range["AZ" + matricis[0].GetLength(0)]];
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

        private void WriteHeadersForTimeCardImport(Excel.Worksheet workSheet)
        {
            workSheet.Cells[1, 1] = "Key";
            workSheet.Cells[1, 2] = "E_Hourly Regular_Hours";
            workSheet.Cells[1, 3] = "E_Blended Overtim_Hours";
            workSheet.Cells[1, 4] = "E_Min Guaran_Hours";
            workSheet.Cells[1, 5] = "E_Holiday_Hours";
            workSheet.Cells[1, 6] = "E_Vacation_Hours";
            workSheet.Cells[1, 7] = "E_Hourly Regular_WeekNumber";
            workSheet.Cells[1, 8] = "E_Blended Overtim_WeekNumber";
            workSheet.Cells[1, 9] = "E_Vacation_Dollars";
            workSheet.Cells[1, 10] = "E_Hourly Regular_Dollars";
            workSheet.Cells[1, 11] = "E_Per Diem_Dollars";
            workSheet.Cells[1, 12] = "LaborValue1";
            workSheet.Cells[1, 13] = "E_Holiday_Dollars";
            workSheet.Cells[1, 14] = "E_Vacation_WeekNumber";
            workSheet.Cells[1, 15] = "E_Holiday_WeekNumber";
            workSheet.Cells[1, 16] = "E_Summer Bonus_Hours";
            workSheet.Cells[1, 17] = "E_Summer Bonus_WeekNumber";
            workSheet.Cells[1, 18] = "E_Min Guaran_WeekNumber";
            workSheet.Cells[1, 19] = "E_Bonus_Dollars";
            workSheet.Cells[1, 20] = "E_Min Guaran_Dollars";
            workSheet.Cells[1, 21] = "E_Summer Bonus_Dollars";
        }

        private enum TimeCardImportColumns
        {
            EMP_NUMBER = 0, REGULAR_HOURS = 1, OT_HOURS = 2, MG_HOURS = 3, HOLIDAY_HOURS = 4, VACATION_HOURS = 5, REGULAR_HOURS_WEEK = 6, OT_WEEK = 7, VACATION_DOLLARS = 8, REGULAR_DOLLARS = 9, PER_DIEM_DOLLARS_COLUMN = 10, JOB_CODE = 11, HOLIDAY_DOLLARS = 12, VACATION_WEEK = 13, HOLIDAY_WEEK = 14, SUMMER_BONUS_HOURS = 15, SUMMER_BONUS_WEEK = 16, MG_WEEK = 17, BONUS_DOLLARS_COLUMN = 18, MG_DOLLARS = 19, SUMMER_BONUS_DOLLARS = 20
        }

        private void SaveWorkBook(Excel.Workbook workBook, string filePath)
        {
            try
            {
                object misValue = System.Reflection.Missing.Value;
                workBook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
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

        private void WriteToMatrix(Employee emp, Shift shift, float time, float dollarAmount, TimeCardImportColumns timeColumn, TimeCardImportColumns weekColumn, TimeCardImportColumns dollarColumn, ref int rowCounter, object[,] matrix)
        {
            matrix[rowCounter, (int)weekColumn] = shift.WeekNumber;
            matrix[rowCounter, (int)TimeCardImportColumns.EMP_NUMBER] = emp.IdNumber.ToString();
            matrix[rowCounter, (int)TimeCardImportColumns.JOB_CODE] = shift.GetLaborCode(false);
            if (time > 0.001f)
            {
                matrix[rowCounter, (int)timeColumn] = Math.Round(time, 2);
                if (shift.HasSpecialPayRate(emp) && dollarAmount < 0.01f)
                {
                    //todo: test this
                    matrix[rowCounter, (int)dollarColumn] = Math.Round(shift.SpecialRate(emp) * time, 2);
                }
            }
            if (dollarAmount > 0f)
            {
                matrix[rowCounter, (int)dollarColumn] = Math.Round(dollarAmount, 2);
            }
            if (shift.PerDiem > 0f)
            {
                matrix[rowCounter, (int)TimeCardImportColumns.PER_DIEM_DOLLARS_COLUMN] = Math.Round(shift.PerDiem, 2);
            }
            if (shift.BonusDollars > 0)
            {
                matrix[rowCounter, (int)TimeCardImportColumns.BONUS_DOLLARS_COLUMN] = Math.Round(shift.BonusDollars, 2);
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
                case 3:
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

        private void TryParseTimeSpan(Object cellData, out TimeSpan timeSpan)
        {
            timeSpan = new TimeSpan();
            if (cellData != null && cellData.ToString() != null)
            {
                DateTime dt;
                if (DateTime.TryParse(cellData.ToString(), out dt))
                {
                    timeSpan = dt.TimeOfDay;
                    return;
                }

                double oaDate;
                if (double.TryParse(cellData.ToString(), out oaDate))
                {
                    timeSpan = TimeSpan.FromHours(oaDate);
                    TimeSpan t2 = DateTime.FromOADate(oaDate).TimeOfDay;
                    DelayedLog("Check time span parsing.", true);
                    return;
                }

                DelayedLog("Warning: Couldn't parse TimeSpan for " + cellData.ToString());
            }
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
            bool bLastDayWasReached = false;
            for (int i = 1; i < 14; ++i)
            {
                DateTime nextDay = firstDay.AddDays(i);
                if (nextDay.CompareTo(lastDay) > 0)
                {
                    bLastDayWasReached = true;
                    break;
                }
                else
                {
                    dates.Add(nextDay);
                }
            }

            return dates;
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
