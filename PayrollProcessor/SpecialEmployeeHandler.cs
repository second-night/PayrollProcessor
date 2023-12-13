using System.Text.Json;
using System.Windows.Forms;
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
            if (SpecialEmployees.ShiftMgExceptions.Count == 0 && SpecialEmployees.PayRateExceptions.Count == 0)
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
            SpecialEmployees.ShiftMgExceptions.ForEach(entry => LogEntry(entry.IdNumber, entry.Hours));
            ExceptionLog += "\n";
            SpecialEmployees.SpecificShiftMgExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname receiving a MG of " + entry.Hours + " hours per shift for shifts during the " + ((RouteTimeContext)entry.ShiftNumber).ToString() + ".", SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\nOther exceptions: \n";
            SpecialEmployees.WeeklyMgExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a weekly MG of " + entry.Hours + " hours.", SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.DailyMgExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a daily MG of " + entry.Hours + " hours.", SpecialMgNonShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.ShiftMgExceptionsInDollars.ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a MG of $" + entry.Dollars + " per shift.", SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.SmallMgExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname is receiving a specially reduced MG of " + entry.Hours + " hours, specifically while driving bus# " + entry.BusNumber + ".", SpecialMgShiftTotals.GetValueOrDefault(entry.IdNumber)));
            ExceptionLog += "\n";
            SpecialEmployees.PayRateExceptions.ForEach(entry => LogEntry(entry.IdNumber, "empname receives their payrate for " + ((Jobs)entry.OverridingJobType).ToString() + " when they clock in as " + ((Jobs)entry.OverriddenJobType).ToString() + ".", 0f, false));
            ExceptionLog += "\n\n\n\n";
            Log(ExceptionLog, true);
        }

        public void CheckForTimeFrameException(Employee employee, Shift shift)
        {
            foreach (var entry in SpecialEmployees.LimitedTimeFrameExceptions)
            {
                if (entry.IdNumber == employee.IdNumber)
                {
                    if (TimeSpan.TryParse(entry.EarliestClockIn, out TimeSpan earliestClockIn))
                    {
                        if (shift.ClockIn.CompareTo(earliestClockIn) < 0)
                        {
                            shift.ModifyClockIn(earliestClockIn);
                        }
                    }
                    if (TimeSpan.TryParse(entry.LatestClockOut, out TimeSpan latestClockOut))
                    {
                        if (shift.ClockOut.CompareTo(latestClockOut) > 0)
                        {
                            shift.ModifyClockOut(latestClockOut);
                        }
                    }
                }
            }
        }

        private void LogEntry(int employeeIdNumber, string message, float hoursGiven, bool bShouldDisplayTotals = true)
        {
            if (EmployeeDictionary.ContainsKey(employeeIdNumber))
            {
                ExceptionLog += message.Replace("empname", EmployeeDictionary[employeeIdNumber].Name) + (bShouldDisplayTotals ? ((EmployeeDictionary[employeeIdNumber].IsMale ? " He" : " She") + " recieved a total of " + Math.Round(hoursGiven, 2) + " hours for this exception.") : "") + "\n";
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
                ExceptionLog += EmployeeDictionary[employeeIdNumber].Name + ": " + guarantee + " hours guaranteed, " + Math.Round(SpecialMgShiftTotals.GetValueOrDefault(employeeIdNumber, 0f), 2) + " hours earned from guarantee.\n";
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

        public List<SpecialHoursEntry> DailyMgExceptions { get; set; } = new();

        public List<SpecialDollarsEntry> ShiftMgExceptionsInDollars { get; set; } = new();

        public List<SpecialHoursEntry> ShiftMgExceptions { get; set; } = new();

        public List<SpecialBusEntry> SmallMgExceptions { get; set; } = new();

        public List<SpecialShiftEntry> SpecificShiftMgExceptions { get; set; } = new();

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
    }

    public class SpecialHoursEntry : SpecialEntry
    {
        public float Hours { get; set; }
    }

    public class SpecialDollarsEntry : SpecialEntry
    {
        public float Dollars { get; set; }
    }

    public class SpecialBonusDollarsEntry : SpecialDollarsEntry
    {
        public int JobType { get; set; }
    }

    public class SpecialBusEntry : SpecialHoursEntry
    {
        public int BusNumber { get; set; }
    }

    public class SpecialShiftEntry : SpecialHoursEntry
    {
        public int ShiftNumber { get; set; }
    }

    public class SpecialPayRateEntry : SpecialEntry
    {
        public int OverriddenJobType { get; set; }
        public int OverridingJobType { get; set; }
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
