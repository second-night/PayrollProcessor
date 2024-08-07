using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class RaiseTracker
    {
        public EmployeeRecords EmployeeRecords { get; set; }

        private static RaiseTracker? Instance;
        private List<Jobs> RelevantJobs = new List<Jobs> { Jobs.MECHANIC, Jobs.WASH_BAY, Jobs.BODY_SHOP, Jobs.ADMIN, Jobs.CLEANING, Jobs.SALARY };
        private RaiseTracker()
        {
            try
            {
                string path = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.Parent.Parent.FullName;
                string filePath = path + "\\PayrollProcessor\\PayrollProcessor\\EmployeeRecords.json";
                //purpose of two files - the main file is at the front and therefore easier to find and edit, but isn't in the git directory. The backup file is included in git.
                if (!File.Exists(filePath))
                {
                    Log("Error loading EmployeeRecords Json. If you have moved this program, please make sure that the folder structure has stayed intact beginning with the folder 'Payroll'. This should not be ignored.", true);
                }
                string file = File.ReadAllText(filePath);
                EmployeeRecords = JsonSerializer.Deserialize<EmployeeRecords>(file);
            }
            catch (Exception)
            {
                Log("Error loading EmployeeRecords Json. Either the file format is incorrect or the file was not found. If you have moved this program, please make sure that the folder structure has stayed intact beginning with the folder 'Payroll'. This should not be ignored.", true);
            }

        }

        public static RaiseTracker GetInstance()
        {
            if (null == Instance)
            {
                Instance = new();
            }
            return Instance;
        }

        public void CheckEmployeeForPayChange(Employee employee)
        {
            if (null != EmployeeRecords)
            {
                bool bShouldAddEmployee = false;
                EmployeeEntry employeeEntry = null;
                foreach (var payRate in employee.PayRates)
                {
                    if (RelevantJobs.Contains(payRate.Key))
                    {
                        if (EmployeeRecords.Contains(employee))
                        {
                            employeeEntry = EmployeeRecords.FindEmployeeEntry(employee);
                        }
                        else
                        {
                            bShouldAddEmployee = true;
                            employeeEntry = new(
                                employee.Name,
                                employee.IdNumber,
                                "",
                                new(),
                                new()
                                );
                        }
                        float oldRate = employeeEntry.CurrentPayRates.ContainsKey((int)payRate.Key) ? employeeEntry.CurrentPayRates[(int) payRate.Key] : 0f;
                        if (oldRate < payRate.Value)
                        {
                            employeeEntry.CurrentPayRates[(int)payRate.Key] = payRate.Value;
                            employeeEntry.Raises.Add(new(payRate.Key.ToString(), (int)payRate.Key, oldRate, payRate.Value, DateTime.Now));
                        }
                    }
                }
                if (bShouldAddEmployee)
                {
                    //EmployeeRecords.Add()
                }
            }
        }
    }

    public class EmployeeRecords
    {
        public List<EmployeeEntry> Entries { get; set; } = new();

        public bool Contains(Employee employee)
        {
            return Contains(employee.IdNumber);
        }

        public bool Contains(int id)
        {
            foreach (EmployeeEntry entry in Entries)
            {
                if (entry.IdNumber == id)
                {
                    return true;
                }
            }
            return false;
        }

        public EmployeeEntry FindEmployeeEntry(Employee employee)
        {
            foreach (EmployeeEntry entry in Entries)
            {
                if (entry.IdNumber == employee.IdNumber)
                {
                    return entry;
                }
            }
            return null;
        }

        public void AddEmployeeEntry(EmployeeEntry employeeEntry)
        {
            if (Contains(employeeEntry.IdNumber))
            {
                Log("Warning: trying to add an employeeEntry when the employeeEntry already is in the registry)");
                return;
            }
            Entries.Add(employeeEntry);
        }
    }

    public class EmployeeEntry : SpecialEntry
    {
        public List<Raise> Raises { get; set; }
        public Dictionary<int, float> CurrentPayRates { get; set; }

        public EmployeeEntry(string name, int idNumber, string notes, List<Raise> raises, Dictionary<int, float> currentPayRates) : base(name, idNumber, notes)
        {
            Raises = raises;
            CurrentPayRates = currentPayRates;
        }

        public EmployeeEntry()
        {
        }
    }

    public class Raise
    {
        public string? JobName { get; set; }
        public int JobIdNumber { get; set; }
        public float PreviousRate { get; set; }
        public float NewRate { get; set; }
        public DateTime? DateOfChange { get; set; }

        public Raise(string jobName, int jobIdNumber, float previousRate, float newRate, DateTime? dateOfChange)
        {
            JobName = jobName;
            JobIdNumber = jobIdNumber;
            PreviousRate = previousRate;
            NewRate = newRate;
            DateOfChange = dateOfChange;
        }
    }
}
