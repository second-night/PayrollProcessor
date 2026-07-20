using System.Diagnostics;
using System.Text;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    internal class VacationTracker
    {
        public const float MinimumCompensatedHoursForAccrual = 60f;
        private const float CoachHoursPerDay = 8f;

        private static readonly List<string> CsvHeaders = new()
        {
            "PositionID",
            "TimeOffPolicyName",
            "TransactionType",
            "ReasonCodes",
            "TransactionStartDate",
            "TransactionStartTime",
            "TransactionAmount",
            "TransactionUnit",
            "SendToPayroll"
        };

        public void ProcessAndWriteCsv(IEnumerable<Employee> employees)
        {
            List<Dictionary<string, string>> rows = new();
            string transactionDate = DateTime.Today.ToString("MM/dd/yyyy");

            foreach (Employee emp in employees.OrderBy(e => e.IdNumber))
            {
                if (emp.IsSalaried || emp.IsPartialEntry || EmployeeIdsToIgnore.Contains(emp.IdNumber))
                {
                    continue;
                }

                if (emp.Shifts.Count == 0 && !HasShiftTotals(emp) && emp.ManualEntries.Count == 0)
                {
                    continue;
                }

                float compensatedHours = GetCompensatedHoursForPayPeriod(emp);
                float vacationTaken = GetVacationHoursTaken(emp);
                float accrual = compensatedHours >= MinimumCompensatedHoursForAccrual
                    ? GetAccrualRateForYearsOfService(emp.YearsOfService)
                    : 0f;
                float transactionAmount = (float)Math.Round(accrual - vacationTaken, 3);

                if (Math.Abs(transactionAmount) < 0.001f)
                {
                    continue;
                }

                emp.NetVacationChangeForPayPeriod = transactionAmount;

                rows.Add(new Dictionary<string, string>
                {
                    ["PositionID"] = "MMF" + emp.IdNumber.ToString("D6"),
                    ["TimeOffPolicyName"] = "Valley Vacation",
                    ["TransactionType"] = "External Award",
                    ["ReasonCodes"] = "",
                    ["TransactionStartDate"] = transactionDate,
                    ["TransactionStartTime"] = "",
                    ["TransactionAmount"] = FormatTransactionAmount(transactionAmount),
                    ["TransactionUnit"] = "hours",
                    ["SendToPayroll"] = ""
                });

                LogVacationSummary(emp, compensatedHours, accrual, vacationTaken, transactionAmount);
            }

            string path = DesktopPath() + "AccrualsImport.csv";
            WriteCsv(path, rows);
            Log("Vacation time-off import written to " + path + " (" + rows.Count + " rows).");

            if (rows.Count > 0)
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
        }

        public static float GetAccrualRateForYearsOfService(int yearsOfService)
        {
            if (yearsOfService >= 10)
            {
                return 6.154f;
            }
            if (yearsOfService >= 5)
            {
                return 4.616f;
            }
            if (yearsOfService >= 2)
            {
                return 3.077f;
            }
            return 1.539f;
        }

        public static float GetCompensatedHoursForPayPeriod(Employee emp)
        {
            float total = 0f;

            for (int company = (int)Company.VALLEY_BUS_LLC; company <= (int)Company.VALLEY_BUS_COACHES; ++company)
            {
                for (int shiftType = 0; shiftType < 3; ++shiftType)
                {
                    if (emp.ShiftTotals[company, shiftType] == null)
                    {
                        continue;
                    }

                    foreach (Dictionary<int, List<Shift>> weekMap in emp.ShiftTotals[company, shiftType].Values)
                    {
                        foreach (List<Shift> shifts in weekMap.Values)
                        {
                            foreach (Shift shift in shifts)
                            {
                                if (shift.JobType == Jobs.DRIVER_COACH || !shift.IsValid(emp))
                                {
                                    continue;
                                }

                                total += shift.AllHours(false);
                            }
                        }
                    }
                }
            }

            total += GetCoachDrivingHours(emp);

            foreach (ManualEntry entry in emp.ManualEntries)
            {
                total += entry.AllHours();
            }

            return total;
        }

        private static float GetCoachDrivingHours(Employee emp)
        {
            int coachDays = 0;

            foreach (Shift shift in emp.Shifts)
            {
                if (shift.IsATotalsShift || shift.JobType != Jobs.DRIVER_COACH || !shift.IsValid(emp))
                {
                    continue;
                }

                if (shift.ShiftTime + shift.DollarAmount + shift.BonusDollars + shift.PerDiem < 0.01f)
                {
                    continue;
                }

                coachDays += shift.CoachTripDays;
            }

            return coachDays * CoachHoursPerDay;
        }

        private static float GetVacationHoursTaken(Employee emp)
        {
            float total = 0f;

            for (int company = (int)Company.VALLEY_BUS_LLC; company <= (int)Company.VALLEY_BUS_COACHES; ++company)
            {
                for (int shiftType = 0; shiftType < 3; ++shiftType)
                {
                    if (emp.ShiftTotals[company, shiftType] == null)
                    {
                        continue;
                    }

                    foreach (Dictionary<int, List<Shift>> weekMap in emp.ShiftTotals[company, shiftType].Values)
                    {
                        foreach (List<Shift> shifts in weekMap.Values)
                        {
                            foreach (Shift shift in shifts)
                            {
                                if (shift.JobType == Jobs.VACATION && shift.IsValid(emp))
                                {
                                    total += shift.ShiftTime;
                                }
                            }
                        }
                    }
                }
            }

            foreach (ManualEntry entry in emp.ManualEntries)
            {
                total += entry.VacationHours + entry.RoundUpVacationHours;
            }

            return total;
        }

        private static bool HasShiftTotals(Employee emp)
        {
            for (int company = (int)Company.VALLEY_BUS_LLC; company <= (int)Company.VALLEY_BUS_COACHES; ++company)
            {
                for (int shiftType = 0; shiftType < 3; ++shiftType)
                {
                    if (emp.ShiftTotals[company, shiftType] != null)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string FormatTransactionAmount(float amount)
        {
            return amount.ToString("0.###");
        }

        private static void LogVacationSummary(Employee emp, float compensatedHours, float accrual, float vacationTaken, float transactionAmount)
        {
            Log("Vacation: " + emp.Name + " (" + emp.IdNumber + ") — "
                + Math.Round(compensatedHours, 2) + " compensated hrs, "
                + Math.Round(accrual, 3) + " accrued, "
                + Math.Round(vacationTaken, 2) + " taken, "
                + "transaction " + FormatTransactionAmount(transactionAmount));
        }

        private static void WriteCsv(string path, List<Dictionary<string, string>> rows)
        {
            using StreamWriter writer = new(path, false, new UTF8Encoding(false));
            writer.WriteLine(string.Join(",", CsvHeaders.Select(EscapeCsv)));
            foreach (Dictionary<string, string> row in rows)
            {
                writer.WriteLine(string.Join(",", CsvHeaders.Select(header => EscapeCsv(row.GetValueOrDefault(header, "")))));
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
    }
}
