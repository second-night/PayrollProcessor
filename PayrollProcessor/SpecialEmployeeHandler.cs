using System.Text.Json;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class SpecialEmployeeHandler
    {
        public SpecialEmployees SpecialEmployees;

        private static SpecialEmployeeHandler? Instance;

        public static Dictionary<int, float> SpecialMgShiftTotals = new();

        public static Dictionary<int, float> SpecialMgNonShiftTotals = new();

        private static string ExceptionLog = "";

        private SpecialEmployeeHandler()
        {
            try
            {
                string path = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.Parent.Parent.FullName;
                string mainFile = path + "\\SpecialEmployees.json";
                string backUpFile = path + "\\PayrollProcessor\\PayrollProcessor\\SpecialEmployeesBackup.json";
                //purpose of two files - the main file is at the front and therefore easier to find and edit, but isn't in the git directory. The backup file is included in git.
                if (!File.Exists(mainFile))
                {
                    if (File.Exists(backUpFile))
                    {
                        mainFile = backUpFile;
                        backUpFile = path + "\\SpecialEmployees.json";
                    }
                    else
                    {
                        Log("Error loading special exceptions Json. If you have moved this program, please make sure that the folder structure has stayed intact beginning with the folder 'Payroll'. This should not be ignored.", true);
                    }
                }
                string file = File.ReadAllText(mainFile);
                SpecialEmployees = JsonSerializer.Deserialize<SpecialEmployees>(file);
                try
                {
                    File.Copy(mainFile, backUpFile, true);
                }
                catch (Exception)
                {
                    Log("Warning: Problem backing up SpecialEmployees.json", true);
                }
            }
            catch (Exception)
            {
                Log("Error loading special exceptions Json. Either the file format is incorrect or the file was not found. If you have moved this program, please make sure that the folder structure has stayed intact beginning with the folder 'Payroll'. This should not be ignored.", true);
            }
            if (SpecialEmployees.ShiftMgExceptions.Count == 0 && SpecialEmployees.PayRateSubstitutionExceptions.Count == 0)
            {
                Log("Error loading special exceptions Json. Please make sure the file's json format has not been comprimised. Employee exceptions will not be active unless this is fixed.", true);
            }
        }

        public static SpecialEmployeeHandler GetInstance()
        {
            if (null == Instance)
            {
                Instance = new SpecialEmployeeHandler();
            }
            return Instance;
        }

        public void CheckForMgExceptionForShift(Employee emp, Shift shift, out float maxMgTime)
        {
            maxMgTime = 0f;
            if (shift.JobType == Jobs.DRIVER_SCHOOL || shift.JobType == Jobs.AIDE_SCHOOL)
            {
                if (emp.IdNumber == 2206)
                {
                    maxMgTime = 2.5f;
                    return;
                }
                foreach (var entry in SpecialEmployees.ShiftMgExceptionsInDollars)
                {
                    if (entry != null && entry.IdNumber == emp.IdNumber)
                    {
                        float rate = shift.JobType == Jobs.AIDE_SCHOOL ? emp.PayRates.GetValueOrDefault(Jobs.AIDE_SCHOOL, FargoDefaultRates.GetValueOrDefault(Jobs.AIDE_SCHOOL)) : emp.GetDriverRateForSchoolRouteShift(shift);
                        maxMgTime = Math.Max(maxMgTime, entry.Dollars / rate);
                        break;
                    }
                }
                foreach (var entry in SpecialEmployees.ShiftMgExceptions)
                {
                    if (entry != null && entry.IdNumber == emp.IdNumber)
                    {
                        maxMgTime = Math.Max(maxMgTime, entry.Hours);
                        break;
                    }
                }
                foreach (var entry in SpecialEmployees.SpecificShiftMgExceptions)
                {
                    if (entry != null && entry.IdNumber == emp.IdNumber && shift.TimeContext() == (RouteTimeContext)entry.ShiftNumber)
                    {
                        maxMgTime = Math.Max(maxMgTime, entry.Hours);
                        break;
                    }
                }
            }
        }

        public void AddExceptionNotificationsToLog()
        {
            ExceptionLog += "The following special exceptions are currently in place:\n\n";
            ExceptionLog += "Employees who have a special mg for each shift:\n";
            SpecialEmployees.ShiftMgExceptions.FindAll(entry => SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber) > 0).ForEach(entry => LogEntry(entry.IdNumber, entry.Hours));
            ExceptionLog += "\n";
            SpecialEmployees.SpecificShiftMgExceptions.FindAll(entry => SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber) > 0).ForEach(entry => LogEntry(entry.IdNumber, "empname receiving a MG of " + entry.Hours + " hours per shift for shifts during the " + ((RouteTimeContext)entry.ShiftNumber).ToString() + ".", SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\nOther exceptions: \n";
            SpecialEmployees.WeeklyMgExceptions.FindAll(entry => SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber) > 0).ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a weekly MG of " + entry.Hours + " hours.", SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.WeeklyCompensationFloorExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname has a weekly compensation floor of " + entry.Hours + " hours at route pay. If normal weekly pay is lower, MG hours are added at route pay to reach that floor.", SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber), SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber) > 0));
            ExceptionLog += "\n";
            SpecialEmployees.DailyMgExceptions.FindAll(entry => SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber) > 0).ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a daily MG of " + entry.Hours + " hours.", SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.ShiftMgExceptionsInDollars.FindAll(entry => SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber) > 0).ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a MG of $" + entry.Dollars + " per shift.", SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.SmallMgExceptions.FindAll(entry => SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber) > 0).ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a specially reduced MG of " + entry.Hours + " hours, specifically while driving bus# " + entry.BusNumber + ".", SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.PayRateSubstitutionExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname receives their payrate for " + ((Jobs)entry.OverridingJobType).ToString() + " when they clock in as " + ((Jobs)entry.OverriddenJobType).ToString() + ".", 0f, false));
            ExceptionLog += "\n";
            SpecialEmployees.PayRateExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname receives a special payrate of " + (entry.Rate).ToString() + " when they clock in as " + ((Jobs)entry.JobType).ToString() + ".", 0f, false));
            ExceptionLog += "\n\n\n\n";
            Log(ExceptionLog);
        }

        public void ApplyWeeklyCompensationFloors(Employee emp, DateTime firstDayWeek2)
        {
            if (emp == null || SpecialEmployees.WeeklyCompensationFloorExceptions == null || SpecialEmployees.WeeklyCompensationFloorExceptions.Count == 0)
            {
                return;
            }

            foreach (var entry in SpecialEmployees.WeeklyCompensationFloorExceptions)
            {
                if (entry == null || entry.IdNumber != emp.IdNumber || entry.Hours < 0.01f)
                {
                    continue;
                }

                DateTime? effectiveDate = ParseEffectiveDate(entry.EffectiveDate);
                Jobs floorJob = entry.JobType != 0 ? (Jobs)entry.JobType : Jobs.DRIVER_SCHOOL;
                float routePay = emp.PayRates.GetValueOrDefault(floorJob, 0f);
                if (routePay < 0.01f)
                {
                    Log("Cannot apply weekly compensation floor for " + emp.Name + " because route pay was not found.", true);
                    continue;
                }

                for (int weekNumber = 1; weekNumber < 3; ++weekNumber)
                {
                    DateTime weekStart = weekNumber == 1 ? firstDayWeek2.AddDays(-7) : firstDayWeek2;
                    DateTime weekEnd = weekStart.AddDays(6);
                    if (effectiveDate.HasValue && weekEnd.Date < effectiveDate.Value.Date)
                    {
                        continue;
                    }

                    GetWeeklyPayables(emp, weekNumber, out float weeklyCompensation, out float workingHours);
                    if (workingHours < 0.01f)
                    {
                        continue;
                    }

                    float floorDollars = routePay * entry.Hours;
                    float shortfall = floorDollars - weeklyCompensation;
                    if (shortfall < 0.01f)
                    {
                        continue;
                    }

                    Shift? shift = emp.FindShiftForWeek(weekNumber, floorJob, Company.VALLEY_BUS_LLC, false)
                        ?? emp.FindShiftForWeek(weekNumber, emp.PrimaryJobType(), Company.VALLEY_BUS_LLC, false)
                        ?? emp.FindShiftForWeek(weekNumber, floorJob, Company.VALLEY_BUS_LLC, true);
                    if (shift == null || shift.JobType == Jobs.HOLIDAY || shift.JobType == Jobs.VACATION)
                    {
                        Log("Cannot apply weekly compensation floor for " + emp.Name + " because no suitable shift was found for week " + weekNumber + ".", true);
                        continue;
                    }

                    if (shift.PayRate == null || shift.PayRate < 0.01f)
                    {
                        shift.PayRate = routePay;
                    }

                    float mgPayRate = shift.PayRate.Value;
                    float mgHours = (float)Math.Round(shortfall / mgPayRate, 2);
                    if (mgHours < 0.01f)
                    {
                        continue;
                    }

                    shift.MinimumGuaranteeHours += mgHours;
                    SpecialMgNonShiftTotals[emp.IdNumber] = SpecialMgNonShiftTotals.GetValueOrDefault(emp.IdNumber, 0f) + mgHours;
                    DelayedLog("Giving " + mgHours + " weekly compensation floor hours ($"
                        + Math.Round(mgHours * mgPayRate, 2) + ") to " + emp.Name + " for week " + weekNumber
                        + " (floor is " + entry.Hours + " hours at $" + routePay + "/hr = $" + Math.Round(floorDollars, 2)
                        + "; normal pay was $" + Math.Round(weeklyCompensation, 2) + ").");
                }
            }
        }

        private static DateTime? ParseEffectiveDate(string? effectiveDate)
        {
            if (string.IsNullOrWhiteSpace(effectiveDate))
            {
                return null;
            }
            if (DateTime.TryParse(effectiveDate, out DateTime parsed))
            {
                return parsed.Date;
            }
            Log("Could not parse EffectiveDate '" + effectiveDate + "' for a weekly compensation floor exception.", true);
            return null;
        }

        private static void GetWeeklyPayables(Employee emp, int weekNumber, out float weeklyCompensation, out float workingHours)
        {
            weeklyCompensation = 0f;
            workingHours = 0f;
            for (int company = 0; company < 2; ++company)
            {
                for (int shiftType = 0; shiftType < 3; ++shiftType)
                {
                    if (emp.ShiftTotals[company, shiftType] == null)
                    {
                        continue;
                    }
                    foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                    {
                        if (!pair.TryGetValue(weekNumber, out List<Shift>? shifts) || shifts == null)
                        {
                            continue;
                        }
                        foreach (Shift shift in shifts)
                        {
                            if (!shift.IsValid(emp))
                            {
                                continue;
                            }
                            workingHours += shift.WorkingHours();
                            weeklyCompensation += (shift.PayRate ?? 0f) * shift.AllHours(false)
                                + shift.DollarAmount
                                + shift.BonusDollars;
                        }
                    }
                }
            }
        }

        public void CheckForTimeFrameException(Employee employee, Shift shift)
        {
            if (shift.JobType != Jobs.WASH_BAY && shift.JobType != Jobs.MECHANIC && shift.JobType != Jobs.ADMIN && shift.JobType != Jobs.BODY_SHOP)
            {
                return;
            }
            foreach (var entry in SpecialEmployees.LimitedTimeFrameExceptions)
            {
                if (StringSearch(entry.Notes, "void"))
                {
                    continue;
                }

                if (entry.IdNumber == employee.IdNumber)
                {
                    if (TimeSpan.TryParse(entry.EarliestClockIn, out TimeSpan earliestClockIn))
                    {
                        if (shift.ClockIn.CompareTo(earliestClockIn) < 0)
                        {
                            shift.ModifyClockIn(earliestClockIn);
                            Log("Modifying Clock in time for " + employee.Name + ".");
                        }
                    }
                    if (TimeSpan.TryParse(entry.LatestClockOut, out TimeSpan latestClockOut))
                    {
                        if (shift.ClockOut.CompareTo(latestClockOut) > 0)
                        {
                            shift.ModifyClockOut(latestClockOut);
                            Log("Modifying Clock out time for " + employee.Name + ".");
                        }
                    }
                }
            }
        }

        private void LogEntry(int employeeIdNumber, string message, float hoursGiven, bool bShouldDisplayTotals = true)
        {
            if (EmployeeDictionary.ContainsKey(employeeIdNumber))
            {
                ExceptionLog += message.Replace("empname", EmployeeDictionary[employeeIdNumber].Name) + (bShouldDisplayTotals ? ((EmployeeDictionary[employeeIdNumber].IsMale ? " He" : " She") + " received a total of " + Math.Round(hoursGiven, 2) + " hours for this exception.") : "") + "\n";
            }
            else
            {
                Log("Warning: There is an exception documented for employee " + employeeIdNumber + " but this employee was not found.");
            }
        }

        private void LogEntry(int employeeIdNumber, float guarantee)
        {
            if (EmployeeDictionary.ContainsKey(employeeIdNumber))
            {
                ExceptionLog += EmployeeDictionary[employeeIdNumber].Name + ": " + guarantee + " hours guaranteed, " + Math.Round(SpecialMgShiftTotals.GetValueOrDefault(employeeIdNumber, 0f), 2) + " hours earned from this guarantee.\n";
            }
            else
            {
                Log("Warning: There is an exception documented for employee " + employeeIdNumber + " but this employee was not found.");
            }
        }
    }

    public class SpecialEmployees
    {
        public string? JsonInstructions { get; set; } //for users to view inside the json, has no purpose in this code

        public List<SpecialHoursEntry> WeeklyMgExceptions { get; set; } = new();

        public List<WeeklyCompensationFloorEntry> WeeklyCompensationFloorExceptions { get; set; } = new();

        public List<SpecialHoursEntry> DailyMgExceptions { get; set; } = new();

        public List<SpecialDollarsEntry> ShiftMgExceptionsInDollars { get; set; } = new();

        public List<SpecialHoursEntry> ShiftMgExceptions { get; set; } = new();

        public List<SpecialBusEntry> SmallMgExceptions { get; set; } = new();

        public List<SpecialShiftEntry> SpecificShiftMgExceptions { get; set; } = new();

        public List<SpecialPayRateSubstitutionEntry> PayRateSubstitutionExceptions { get; set; } = new();

        public List<SpecialPayRateEntry> PayRateExceptions { get; set; } = new();

        public List<StartingRateEntry> StartingRateExceptions { get; set; } = new();

        public List<TimeFrameEntry> LimitedTimeFrameExceptions { get; set; } = new();

        public List<SpecialBonusDollarsEntry> BusStartingBonusDollars { get; set; } = new();
    }

    public class SpecialEntry
    {
        public string Name { get; set; }
        public int IdNumber { get; set; }
        public string Notes { get; set; }

        public SpecialEntry(string name, int idNumber, string notes)
        {
            Name = name;
            IdNumber = idNumber;
            Notes = notes;
        }

        public SpecialEntry() { }
    }

    public class SpecialHoursEntry : SpecialEntry
    {
        public float Hours { get; set; }
    }

    public class WeeklyCompensationFloorEntry : SpecialHoursEntry
    {
        public int JobType { get; set; }
        public string? EffectiveDate { get; set; }
    }

    public class SpecialDollarsEntry : SpecialEntry
    {
        public float Dollars { get; set; }
    }

    public class SpecialBonusDollarsEntry : SpecialDollarsEntry
    {
        public int JobType { get; set; }
        public bool ReceivesBusStartingMinimumGuarantee { get; set; }
    }

    public class SpecialBusEntry : SpecialHoursEntry
    {
        public int BusNumber { get; set; }
    }

    public class SpecialShiftEntry : SpecialHoursEntry
    {
        public int ShiftNumber { get; set; }
    }

    public class SpecialPayRateSubstitutionEntry : SpecialEntry
    {
        public int OverriddenJobType { get; set; }
        public int OverridingJobType { get; set; }
    }

    public class SpecialPayRateEntry : SpecialEntry //example: special cdl rate that isn't qualified for raises
    {
        public int JobType { get; set; }
        public float Rate { get; set; }
    }

    public class StartingRateEntry : SpecialEntry
    {
        public int JobType { get; set; }
        public float Rate { get; set; }
    }

    public class TimeFrameEntry : SpecialEntry
    {
        public string EarliestClockIn{ get; set; }
        public string LatestClockOut { get; set; }
    }
}
