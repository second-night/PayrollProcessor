using System.Text.Json;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class SpecialEmployeeHandler
    {
        public SpecialEmployees SpecialEmployees;

        private static SpecialEmployeeHandler? Instance;

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
                    Log("Warning: Problem backing up SpecialEmployees.Json", true);
                }
            }
            catch (Exception)
            {
                Log("Error loading special exceptions Json. Either the file format is incorrect or the file was not found. If you have moved this program, please make sure that the folder structure has stayed intact beginning with the folder 'Payroll'. This should not be ignored.", true);
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

        public void CheckForMgExceptionForShift(Employee emp, Shift shift, ref float maxMgTime)
        {
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


    }

    public class SpecialEmployees
    {
        public string? JsonInstructions { get; set; }

        public List<SpecialHoursEntry> WeeklyMgExceptions { get; set; } = new();

        public List<SpecialHoursEntry> DailyMgExceptions { get; set; } = new();

        public List<SpecialDollarsEntry> ShiftMgExceptionsInDollars { get; set; } = new();

        public List<SpecialHoursEntry> ShiftMgExceptions { get; set; } = new();

        public List<SpecialBusEntry> SmallMgExceptions { get; set; } = new();

        public List<SpecialShiftEntry> SpecificShiftMgExceptions { get; set; } = new();

        public List<SpecialPayRateEntry> PayRateExceptions { get; set; } = new();
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
}
