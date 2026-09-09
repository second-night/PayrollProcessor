using System.Diagnostics;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace PayrollProcessor
{
    internal sealed class PayrollHistoryParser
    {
        public void Run()
        {
            Application.Run(new PayrollHistoryParserForm(new PayrollHistoryCatalog()));
        }
    }

    internal static class PayrollHistoryReportWriter
    {
        public static string Write(PayrollHistoryEmployee employee, List<PayrollHistoryPeriod> periods,
            bool lastSixPayPeriods, DateTime startDate, DateTime endDate)
        {
            string safeName = string.Concat((employee.LastName + employee.FirstName)
                .Where(character => char.IsLetterOrDigit(character)));
            if (safeName == "")
            {
                safeName = "Employee";
            }
            string fileName = "PayrollHistory_" + employee.EmployeeNumber + "_" + safeName + "_"
                + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            string path = CreateReportPath(fileName);

            List<DateTime> payDates = periods.Select(period => period.PayDate.Date).Distinct().OrderBy(date => date).ToList();
            Dictionary<DateTime, Dictionary<Company, PayrollHistoryPeriod>> byDate = new();
            foreach (PayrollHistoryPeriod period in periods)
            {
                if (!byDate.TryGetValue(period.PayDate.Date, out Dictionary<Company, PayrollHistoryPeriod>? byCompany))
                {
                    byCompany = new();
                    byDate[period.PayDate.Date] = byCompany;
                }
                byCompany[period.Company] = period;
            }

            Excel.Application excelApp = new()
            {
                DisplayAlerts = false
            };
            Excel.Workbook? workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Add();
                Excel.Worksheet summary = (Excel.Worksheet)workbook.Worksheets[1];
                summary.Name = "Employee";
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[2]).Delete();
                }
                Excel.Worksheet history = (Excel.Worksheet)workbook.Worksheets.Add(After: summary);
                history.Name = "Pay History";
                Excel.Worksheet detail = (Excel.Worksheet)workbook.Worksheets.Add(After: history);
                detail.Name = "Period Detail";

                WriteSummary(summary, employee, periods, lastSixPayPeriods, startDate, endDate, payDates);
                WriteHistory(history, byDate, payDates);
                WriteDetail(detail, periods);

                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                workbook.SaveAs(path);
                workbook.Close(true);
                workbook = null;
            }
            finally
            {
                workbook?.Close(false);
                excelApp.Quit();
            }

            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            return path;
        }

        public static string WriteHousingRequest(PayrollHistoryEmployee employee, List<PayrollHistoryPeriod> lastSixPeriods)
        {
            string safeName = string.Concat((employee.LastName + employee.FirstName)
                .Where(character => char.IsLetterOrDigit(character)));
            if (safeName == "")
            {
                safeName = "Employee";
            }
            string fileName = "HousingRequest_" + employee.EmployeeNumber + "_" + safeName + "_"
                + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            string path = CreateReportPath(fileName);

            List<DateTime> lastSixDates = lastSixPeriods
                .Select(period => period.PayDate.Date)
                .Distinct()
                .OrderBy(date => date)
                .ToList();
            int weekCount = Math.Max(1, lastSixDates.Count * 2);
            List<PayrollHistoryPeriod> llcSix = lastSixPeriods
                .Where(period => period.Company == Company.VALLEY_BUS_LLC)
                .ToList();
            List<PayrollHistoryPeriod> coachesSix = lastSixPeriods
                .Where(period => period.Company == Company.VALLEY_BUS_COACHES)
                .ToList();
            float totalHours = lastSixPeriods.Sum(period => period.TotalHours);
            float overtimeHours = lastSixPeriods.Sum(period => period.OvertimeHours);
            float llcHours = llcSix.Sum(period => period.TotalHours);
            float coachesHours = coachesSix.Sum(period => period.TotalHours);
            float llcOvertime = llcSix.Sum(period => period.OvertimeHours);
            float coachesOvertime = coachesSix.Sum(period => period.OvertimeHours);
            float averageHoursPerWeek = lastSixDates.Count == 0 ? 0f : totalHours / weekCount;
            float llcHoursPerWeek = lastSixDates.Count == 0 ? 0f : llcHours / weekCount;
            float coachesHoursPerWeek = lastSixDates.Count == 0 ? 0f : coachesHours / weekCount;
            float anticipatedWeeklyOvertime = lastSixDates.Count == 0 ? 0f : overtimeHours / weekCount;
            float llcWeeklyOvertime = lastSixDates.Count == 0 ? 0f : llcOvertime / weekCount;
            float coachesWeeklyOvertime = lastSixDates.Count == 0 ? 0f : coachesOvertime / weekCount;

            int currentYear = DateTime.Today.Year;
            int[] years = { currentYear, currentYear - 1, currentYear - 2 };
            string rangeNote = lastSixDates.Count == 0
                ? "No pay periods found"
                : "Last " + lastSixDates.Count + " pay period(s) / " + weekCount
                    + " weeks, Valley Bus LLC and Valley Bus Coaches combined";

            Excel.Application excelApp = new()
            {
                DisplayAlerts = false
            };
            Excel.Workbook? workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Add();
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets[1];
                sheet.Name = "Housing Request";
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[2]).Delete();
                }

                object[,] values = new object[28, 7];
                values[0, 0] = "Housing Request Employment History";
                values[2, 0] = "Employee Number";
                values[2, 1] = employee.EmployeeNumber;
                values[3, 0] = "Name";
                values[3, 1] = employee.DisplayName;
                values[4, 0] = "Employment Status";
                values[4, 1] = string.IsNullOrWhiteSpace(employee.EmploymentStatus) ? "Unknown" : employee.EmploymentStatus;
                values[5, 0] = "Start Date";
                values[5, 1] = FormatDate(employee.CurrentStartDate);
                values[6, 0] = "End Date";
                values[6, 1] = employee.EmploymentStatus == "Terminated"
                    ? FormatDate(employee.TerminationDate)
                    : "Present";
                values[7, 0] = "Current Gross Base Pay";
                values[7, 1] = FormatBasePay(employee);
                values[8, 0] = "Average Hours Per Week";
                values[8, 1] = Round(averageHoursPerWeek);
                values[8, 2] = rangeNote;
                values[9, 0] = "Valley Bus LLC Hours Per Week";
                values[9, 1] = Round(llcHoursPerWeek);
                values[10, 0] = "Valley Bus Coaches Hours Per Week";
                values[10, 1] = Round(coachesHoursPerWeek);
                values[11, 0] = "Anticipated Weekly Overtime Hours";
                values[11, 1] = Round(anticipatedWeeklyOvertime);
                values[11, 2] = "From overtime hours in the last 6 pay periods, both companies combined";
                values[12, 0] = "Valley Bus LLC Weekly Overtime";
                values[12, 1] = Round(llcWeeklyOvertime);
                values[13, 0] = "Valley Bus Coaches Weekly Overtime";
                values[13, 1] = Round(coachesWeeklyOvertime);

                values[15, 0] = "Year";
                values[15, 1] = "Company";
                values[15, 2] = "Base Pay";
                values[15, 3] = "Overtime";
                values[15, 4] = "Commissions / Tips";
                values[15, 5] = "Bonus";
                values[15, 6] = "Total";
                int row = 16;
                foreach (int year in years)
                {
                    WriteYearRow(values, row++, year, "Valley Bus LLC",
                        SumYear(employee.Periods.Values, year, Company.VALLEY_BUS_LLC));
                    WriteYearRow(values, row++, year, "Valley Bus Coaches",
                        SumYear(employee.Periods.Values, year, Company.VALLEY_BUS_COACHES));
                    WriteYearRow(values, row++, year, "Combined",
                        SumYear(employee.Periods.Values, year, null));
                }

                Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[28, 7]];
                range.Value2 = values;
                sheet.Range["A1"].Font.Bold = true;
                sheet.Range["A1"].Font.Size = 14;
                sheet.Range["A16:G16"].Font.Bold = true;
                sheet.Range["B9:B14"].NumberFormat = "0.00";
                sheet.Range["C17:G25"].NumberFormat = "$#,##0.00";
                sheet.Columns[1].ColumnWidth = 38;
                sheet.Columns[2].ColumnWidth = 22;
                sheet.Columns[3].ColumnWidth = 22;
                sheet.Columns[4].ColumnWidth = 20;
                sheet.Columns[5].ColumnWidth = 20;
                sheet.Columns[6].ColumnWidth = 14;
                sheet.Columns[7].ColumnWidth = 14;

                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                workbook.SaveAs(path);
                workbook.Close(true);
                workbook = null;
            }
            finally
            {
                workbook?.Close(false);
                excelApp.Quit();
            }

            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            return path;
        }

        internal static string OutputFolder =>
            Path.Combine(Path.GetTempPath(), "PayrollHistoryParser");

        internal static void CleanupOutputFolder()
        {
            try
            {
                if (!Directory.Exists(OutputFolder))
                {
                    return;
                }

                foreach (string path in Directory.EnumerateFiles(OutputFolder))
                {
                    try
                    {
                        File.Delete(path);
                    }
                    catch (IOException)
                    {
                    }
                    catch (UnauthorizedAccessException)
                    {
                    }
                }
            }
            catch (IOException)
            {
            }
        }

        private static string CreateReportPath(string fileName)
        {
            Directory.CreateDirectory(OutputFolder);
            return Path.Combine(OutputFolder, fileName);
        }

        private static void WriteSummary(Excel.Worksheet sheet, PayrollHistoryEmployee employee,
            List<PayrollHistoryPeriod> periods, bool lastSixPayPeriods, DateTime startDate, DateTime endDate,
            List<DateTime> payDates)
        {
            object[,] values = new object[16, 2];
            values[0, 0] = "Employee Number";
            values[0, 1] = employee.EmployeeNumber;
            values[1, 0] = "First Name";
            values[1, 1] = employee.FirstName;
            values[2, 0] = "Last Name";
            values[2, 1] = employee.LastName;
            values[3, 0] = "Hire Date";
            values[3, 1] = FormatDate(employee.HireDate);
            values[4, 0] = "Rehire Date";
            values[4, 1] = FormatDate(employee.RehireDate);
            values[5, 0] = "Termination Date";
            values[5, 1] = FormatDate(employee.TerminationDate);
            values[6, 0] = "Last Paid Date";
            values[6, 1] = FormatDate(employee.LastPaidDate);
            values[7, 0] = "Range";
            values[7, 1] = lastSixPayPeriods
                ? "Last 6 pay periods"
                : startDate.ToString("M/d/yyyy") + " - " + endDate.ToString("M/d/yyyy");
            values[8, 0] = "Pay Periods Included";
            values[8, 1] = payDates.Count;

            PayrollHistoryPeriod llc = Sum(periods.Where(period => period.Company == Company.VALLEY_BUS_LLC));
            PayrollHistoryPeriod coaches = Sum(periods.Where(period => period.Company == Company.VALLEY_BUS_COACHES));
            PayrollHistoryPeriod combined = Sum(periods);
            values[10, 0] = "Company";
            values[10, 1] = "Totals for selected range";
            values[11, 0] = "Valley Bus LLC Hours / Gross / Net";
            values[11, 1] = FormatTotals(llc);
            values[12, 0] = "Valley Bus Coaches Hours / Gross / Net";
            values[12, 1] = FormatTotals(coaches);
            values[13, 0] = "Combined Hours / Gross / Net";
            values[13, 1] = FormatTotals(combined);

            Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[16, 2]];
            range.Value2 = values;
            sheet.Columns[1].ColumnWidth = 38;
            sheet.Columns[2].ColumnWidth = 40;
        }

        private static void WriteHistory(Excel.Worksheet sheet,
            Dictionary<DateTime, Dictionary<Company, PayrollHistoryPeriod>> byDate, List<DateTime> payDates)
        {
            string[] headers =
            {
                "Pay Date",
                "Valley Bus LLC Hours", "Valley Bus LLC Gross", "Valley Bus LLC Net",
                "Valley Bus Coaches Hours", "Valley Bus Coaches Gross", "Valley Bus Coaches Net",
                "Combined Hours", "Combined Gross", "Combined Net"
            };
            object[,] values = new object[payDates.Count + 2, headers.Length];
            for (int i = 0; i < headers.Length; i++)
            {
                values[0, i] = headers[i];
            }

            PayrollHistoryPeriod llcTotal = new();
            PayrollHistoryPeriod coachesTotal = new();
            for (int row = 0; row < payDates.Count; row++)
            {
                DateTime payDate = payDates[row];
                byDate.TryGetValue(payDate, out Dictionary<Company, PayrollHistoryPeriod>? byCompany);
                byCompany ??= new();
                byCompany.TryGetValue(Company.VALLEY_BUS_LLC, out PayrollHistoryPeriod? llc);
                byCompany.TryGetValue(Company.VALLEY_BUS_COACHES, out PayrollHistoryPeriod? coaches);
                llc ??= new PayrollHistoryPeriod { PayDate = payDate, Company = Company.VALLEY_BUS_LLC };
                coaches ??= new PayrollHistoryPeriod { PayDate = payDate, Company = Company.VALLEY_BUS_COACHES };
                llcTotal.Add(llc);
                coachesTotal.Add(coaches);

                values[row + 1, 0] = payDate.ToString("M/d/yyyy");
                values[row + 1, 1] = Round(llc.TotalHours);
                values[row + 1, 2] = Round(llc.GrossPay);
                values[row + 1, 3] = Round(llc.NetPay);
                values[row + 1, 4] = Round(coaches.TotalHours);
                values[row + 1, 5] = Round(coaches.GrossPay);
                values[row + 1, 6] = Round(coaches.NetPay);
                values[row + 1, 7] = Round(llc.TotalHours + coaches.TotalHours);
                values[row + 1, 8] = Round(llc.GrossPay + coaches.GrossPay);
                values[row + 1, 9] = Round(llc.NetPay + coaches.NetPay);
            }

            int totalRow = payDates.Count + 1;
            values[totalRow, 0] = "Totals";
            values[totalRow, 1] = Round(llcTotal.TotalHours);
            values[totalRow, 2] = Round(llcTotal.GrossPay);
            values[totalRow, 3] = Round(llcTotal.NetPay);
            values[totalRow, 4] = Round(coachesTotal.TotalHours);
            values[totalRow, 5] = Round(coachesTotal.GrossPay);
            values[totalRow, 6] = Round(coachesTotal.NetPay);
            values[totalRow, 7] = Round(llcTotal.TotalHours + coachesTotal.TotalHours);
            values[totalRow, 8] = Round(llcTotal.GrossPay + coachesTotal.GrossPay);
            values[totalRow, 9] = Round(llcTotal.NetPay + coachesTotal.NetPay);

            Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[payDates.Count + 2, headers.Length]];
            range.Value2 = values;
            sheet.Range[sheet.Cells[2, 2], sheet.Cells[payDates.Count + 2, headers.Length]].NumberFormat = "0.00";
            range.Columns.AutoFit();
        }

        private static void WriteDetail(Excel.Worksheet sheet, List<PayrollHistoryPeriod> periods)
        {
            string[] headers =
            {
                "Pay Date", "Company", "Source", "Total Hours", "Regular Hours", "Overtime Hours",
                "Holiday Hours", "Vacation Hours", "Min Guarantee Hours", "Gross Pay", "Net Pay",
                "Regular Earnings", "Overtime Earnings", "Bonus", "Tips", "Holiday Pay", "Vacation Pay",
                "Back Pay", "Employee Taxes"
            };
            object[,] values = new object[periods.Count + 1, headers.Length];
            for (int i = 0; i < headers.Length; i++)
            {
                values[0, i] = headers[i];
            }
            for (int row = 0; row < periods.Count; row++)
            {
                PayrollHistoryPeriod period = periods[row];
                values[row + 1, 0] = period.PayDate.ToString("M/d/yyyy");
                values[row + 1, 1] = period.Company == Company.VALLEY_BUS_LLC ? "Valley Bus LLC" : "Valley Bus Coaches";
                values[row + 1, 2] = period.Source;
                values[row + 1, 3] = Round(period.TotalHours);
                values[row + 1, 4] = Round(period.RegularHours);
                values[row + 1, 5] = Round(period.OvertimeHours);
                values[row + 1, 6] = Round(period.HolidayHours);
                values[row + 1, 7] = Round(period.VacationHours);
                values[row + 1, 8] = Round(period.MinGuaranteeHours);
                values[row + 1, 9] = Round(period.GrossPay);
                values[row + 1, 10] = Round(period.NetPay);
                values[row + 1, 11] = Round(period.RegularEarnings);
                values[row + 1, 12] = Round(period.OvertimeEarnings);
                values[row + 1, 13] = Round(period.BonusEarnings);
                values[row + 1, 14] = Round(period.TipsEarnings);
                values[row + 1, 15] = Round(period.HolidayEarnings);
                values[row + 1, 16] = Round(period.VacationEarnings);
                values[row + 1, 17] = Round(period.BackPayEarnings);
                values[row + 1, 18] = Round(period.EmployeeTaxes);
            }

            Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[periods.Count + 1, headers.Length]];
            range.Value2 = values;
            if (periods.Count > 0)
            {
                sheet.Range[sheet.Cells[2, 4], sheet.Cells[periods.Count + 1, headers.Length]].NumberFormat = "0.00";
            }
            range.Columns.AutoFit();
        }

        private static void WriteYearRow(object[,] values, int row, int year, string company, YearlyEarnings yearly)
        {
            values[row, 0] = year;
            values[row, 1] = company;
            values[row, 2] = Round(yearly.BasePay);
            values[row, 3] = Round(yearly.Overtime);
            values[row, 4] = Round(yearly.Tips);
            values[row, 5] = Round(yearly.Bonus);
            values[row, 6] = Round(yearly.Gross);
        }

        private static YearlyEarnings SumYear(IEnumerable<PayrollHistoryPeriod> periods, int year, Company? company)
        {
            YearlyEarnings yearly = new();
            foreach (PayrollHistoryPeriod period in periods)
            {
                if (period.PayDate.Year != year)
                {
                    continue;
                }
                if (company.HasValue && period.Company != company.Value)
                {
                    continue;
                }
                yearly.BasePay += period.BasePayEarnings;
                yearly.Overtime += period.OvertimeEarnings;
                yearly.Tips += period.TipsEarnings;
                yearly.Bonus += period.BonusEarnings;
                yearly.Gross += period.GrossPay;
            }
            return yearly;
        }

        private static PayrollHistoryPeriod Sum(IEnumerable<PayrollHistoryPeriod> periods)
        {
            PayrollHistoryPeriod total = new();
            foreach (PayrollHistoryPeriod period in periods)
            {
                total.Add(period);
            }
            return total;
        }

        private static string FormatBasePay(PayrollHistoryEmployee employee)
        {
            if (employee.IsSalaried && employee.AnnualSalary > 0.01f)
            {
                return Round(employee.AnnualSalary).ToString("C", CultureInfo.CurrentCulture) + " annually";
            }
            if (employee.HighestHourlyRate > 0.01f)
            {
                return Round(employee.HighestHourlyRate).ToString("C", CultureInfo.CurrentCulture) + " / hour";
            }
            return "";
        }

        private static string FormatTotals(PayrollHistoryPeriod period) =>
            Round(period.TotalHours).ToString("0.00", CultureInfo.InvariantCulture)
            + " / " + Round(period.GrossPay).ToString("0.00", CultureInfo.InvariantCulture)
            + " / " + Round(period.NetPay).ToString("0.00", CultureInfo.InvariantCulture);

        private static string FormatDate(DateTime? date) =>
            date.HasValue ? date.Value.ToString("M/d/yyyy") : "";

        private static double Round(float value) => Math.Round(value, 2);

        private sealed class YearlyEarnings
        {
            public float BasePay;
            public float Overtime;
            public float Tips;
            public float Bonus;
            public float Gross;
        }
    }
}
