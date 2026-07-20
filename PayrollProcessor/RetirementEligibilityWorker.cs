using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class RetirementEligibilityWorker
    {
        private const int FirstYear = 2016;
        private const int LastYear = 2026;
        private const double FullEligibilityHours = 1000.0;
        private const double LtptHours = 500.0;

        private readonly Dictionary<int, RetirementEmployee> Employees = new();
        private readonly List<string> SourceHeaders = new();

        public void Run()
        {
            Excel.Application excelApp = new();
            excelApp.DisplayAlerts = false;

            try
            {
                for (int year = FirstYear; year <= LastYear; year++)
                {
                    string path = DesktopPath() + year + ".xlsx";

                    if (!File.Exists(path))
                    {
                        Log("Missing file: " + path, true);
                        continue;
                    }

                    ReadYearReport(excelApp, path, year);
                }

                foreach (RetirementEmployee employee in Employees.Values)
                {
                    DetermineEligibility(employee);
                }
                WriteReport(excelApp);

                Log("401(k) eligibility report created on desktop: 401k Eligibility Review.xlsx", true);
            }
            finally
            {
                excelApp.Quit();
            }
        }

        private void ReadYearReport(Excel.Application excelApp, string path, int year)
        {
            Excel.Workbook workbook = excelApp.Workbooks.Open(path);

            try
            {
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    Excel.Range range = sheet.Range[sheet.Range["A1"], sheet.Range["Z5000"]].CurrentRegion;
                    object[,] cellData = (object[,])range.Value2;

                    int rows = cellData.GetLength(0);
                    int cols = cellData.GetLength(1);

                    int headerRow = FindHeaderRow(cellData, rows, cols);
                    if (headerRow == 0)
                    {
                        Log("Could not find header row in " + Path.GetFileName(path), true);
                        continue;
                    }

                    List<string> headers = ReadHeaders(cellData, headerRow, cols);

                    foreach (string header in headers)
                    {
                        if (!SourceHeaders.Contains(header))
                        {
                            SourceHeaders.Add(header);
                        }
                    }

                    int empNumberCol = FindColumn(headers,
                        "employee number", "employee #", "employee id", "emp number", "emp #", "id");

                    int hoursCol = FindColumn(headers,
                        " All Hours Compensated", "hours", "total hours", "hours worked", "worked hours", "regular hours");

                    int matchCol = FindColumn(headers,
                        "401K Match");

                    if (empNumberCol == -1 || hoursCol == -1)
                    {
                        Log("Missing required employee number or hours column in " + Path.GetFileName(path), true);
                        continue;
                    }

                    for (int row = headerRow + 1; row <= rows; row++)
                    {
                        if (!TryGetInt(cellData[row, empNumberCol + 1], out int employeeNumber))
                        {
                            continue;
                        }

                        if (!TryGetDouble(cellData[row, hoursCol + 1], out double hours))
                        {
                            hours = 0;
                        }

                        if (!Employees.ContainsKey(employeeNumber))
                        {
                            Employees[employeeNumber] = new RetirementEmployee(employeeNumber);
                        }

                        RetirementEmployee employee = Employees[employeeNumber];


                        if (TryGetDouble(cellData[row, matchCol + 1], out double match) && match > 0)
                        {
                            employee.IsEligible = true;
                            employee.EligibilityType = "Fully Eligible";
                        }

                        employee.HoursByYear[year] =
                            employee.HoursByYear.GetValueOrDefault(year, 0) + hours;

                        for (int i = 0; i < headers.Count; i++)
                        {
                            string header = headers[i];
                            string value = CellString(cellData[row, i + 1]);

                            if (i == hoursCol)
                            {
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                employee.Fields[header] = value;
                            }
                            else if (!employee.Fields.ContainsKey(header))
                            {
                                employee.Fields[header] = "";
                            }
                        }
                    }
                }
            }
            finally
            {
                workbook.Close(false);
            }
        }

        private void DetermineEligibility(RetirementEmployee employee)
        {
            if (employee.IsEligible)
            {
                return;
            }

            foreach (var entry in employee.HoursByYear.OrderBy(x => x.Key))
            {
                if (entry.Key == 2026)
                {
                    continue; //not using 2026 for eligibility
                }
                if (entry.Value >= FullEligibilityHours)
                {
                    Log("Employee eligible: " + employee.EmployeeNumber);
                    employee.IsEligible = true;
                    employee.EligibilityType = "Fully Eligible";
                    employee.EligibilityYear = entry.Key;
                    return;
                }
            }

            for (int year = FirstYear + 1; year <= LastYear; year++)
            {
                if (year == 2026)
                {
                    continue; //not using 2026 for eligibility
                }

                double previous = employee.HoursByYear.GetValueOrDefault(year - 1, 0);
                double current = employee.HoursByYear.GetValueOrDefault(year, 0);

                if (previous >= LtptHours && current >= LtptHours)
                {
                    employee.IsEligible = true;
                    employee.EligibilityType = "LTPT Eligible";
                    employee.EligibilityYear = year;
                    employee.LtptYears = (year - 1) + " & " + year;
                    return;
                }
            }
        }

        private void WriteReport(Excel.Application excelApp)
        {
            string path = DesktopPath() + "401k Eligibility Review.xlsx";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            Excel.Workbook workbook = excelApp.Workbooks.Add();

            WriteEligibilitySheet(workbook, "Fully Eligible",
                Employees.Values.Where(e => e.IsEligible && e.EligibilityType == "Fully Eligible"));

            WriteEligibilitySheet(workbook, "LTPT Eligible",
                Employees.Values.Where(e => e.IsEligible && e.EligibilityType == "LTPT Eligible"));

            WriteEligibilitySheet(workbook, "InEligible",
                Employees.Values.Where(e => !e.IsEligible));

            //while (workbook.Worksheets.Count > 2)
            //{
            //    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
            //}

            workbook.SaveAs(path);
            workbook.Close(true);
        }

        private void WriteEligibilitySheet(Excel.Workbook workbook, string sheetName, IEnumerable<RetirementEmployee> employees)
        {
            Excel.Worksheet sheet;

            if (workbook.Worksheets.Count == 1 && ((Excel.Worksheet)workbook.Worksheets[1]).UsedRange.Count == 1)
            {
                sheet = workbook.Worksheets[1];
            }
            else
            {
                sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            }

            sheet.Name = sheetName;

            List<string> headers = new()
            {
                "Eligibility Type",
                "Eligibility Year",
                "LTPT Years"
            };

            headers.AddRange(SourceHeaders.Where(h => !IsHoursHeader(h)));

            for (int year = FirstYear; year <= LastYear; year++)
            {
                headers.Add(year + " Hours");
            }

            List<RetirementEmployee> employeeList = employees
                .OrderBy(e => e.EligibilityYear)
                .ThenBy(e => e.Fields.GetValueOrDefault("Last Name", ""))
                .ThenBy(e => e.Fields.GetValueOrDefault("First Name", ""))
                .ThenBy(e => e.EmployeeNumber)
                .ToList();

            object[,] output = new object[employeeList.Count + 1, headers.Count];

            for (int col = 0; col < headers.Count; col++)
            {
                output[0, col] = headers[col];
            }

            for (int row = 0; row < employeeList.Count; row++)
            {
                RetirementEmployee employee = employeeList[row];
                int col = 0;

                output[row + 1, col++] = employee.EligibilityType;
                output[row + 1, col++] = employee.EligibilityYear;
                output[row + 1, col++] = employee.LtptYears;

                foreach (string header in SourceHeaders.Where(h => !IsHoursHeader(h)))
                {
                    output[row + 1, col++] = employee.Fields.GetValueOrDefault(header, "");
                }

                for (int year = FirstYear; year <= LastYear; year++)
                {
                    output[row + 1, col++] = Math.Round(employee.HoursByYear.GetValueOrDefault(year, 0), 2);
                }
            }

            Excel.Range range = sheet.Range[
                sheet.Cells[1, 1],
                sheet.Cells[employeeList.Count + 1, headers.Count]
            ];

            range.Value2 = output;
            range.Columns.AutoFit();
        }

        private static List<string> ReadHeaders(object[,] cellData, int headerRow, int cols)
        {
            List<string> headers = new();

            for (int col = 1; col <= cols; col++)
            {
                string header = CellString(cellData[headerRow, col]);

                if (string.IsNullOrWhiteSpace(header))
                {
                    header = "Column " + col;
                }

                while (headers.Contains(header))
                {
                    header += "_Duplicate";
                }

                headers.Add(header);
            }

            return headers;
        }

        private static int FindHeaderRow(object[,] cellData, int rows, int cols)
        {
            for (int row = 1; row <= Math.Min(rows, 20); row++)
            {
                bool hasEmployee = false;
                bool hasHours = false;

                for (int col = 1; col <= cols; col++)
                {
                    string value = CellString(cellData[row, col]).ToLower();

                    if (value.Contains("employee") || value.Contains("emp"))
                    {
                        hasEmployee = true;
                    }

                    if (value.Contains("hour"))
                    {
                        hasHours = true;
                    }
                }

                if (hasEmployee && hasHours)
                {
                    return row;
                }
            }

            return 0;
        }

        private static int FindColumn(List<string> headers, params string[] possibleHeaders)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                string header = headers[i].Trim().ToLower();

                foreach (string possibleHeader in possibleHeaders)
                {
                    if (header == possibleHeader.ToLower())
                    {
                        return i;
                    }
                }
            }

            for (int i = 0; i < headers.Count; i++)
            {
                string header = headers[i].Trim().ToLower();

                foreach (string possibleHeader in possibleHeaders)
                {
                    if (header.Contains(possibleHeader.ToLower()))
                    {
                        return i;
                    }
                }
            }

            Log("Couldn't find column header for possibleHeaders: " + possibleHeaders, true);
            return -1;
        }

        private static bool IsHoursHeader(string header)
        {
            return header.Contains("hour", StringComparison.OrdinalIgnoreCase);
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

            string text = CellString(cell)
                .Replace(",", "")
                .Replace(".0", "");

            return int.TryParse(text, out value);
        }

        private static bool TryGetDouble(object? cell, out double value)
        {
            value = 0;

            if (cell == null)
            {
                return false;
            }

            if (cell is double d)
            {
                value = d;
                return true;
            }

            string text = CellString(cell).Replace(",", "");

            return double.TryParse(text, out value);
        }

        private class RetirementEmployee
        {
            public int EmployeeNumber { get; }
            public Dictionary<string, string> Fields { get; } = new();
            public Dictionary<int, double> HoursByYear { get; } = new();

            public bool IsEligible { get; set; }
            public string EligibilityType { get; set; } = "";
            public int EligibilityYear { get; set; }
            public string LtptYears { get; set; } = "";

            public RetirementEmployee(int employeeNumber)
            {
                EmployeeNumber = employeeNumber;
            }
        }
    }
}