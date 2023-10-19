using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class Employee
    {

        public int IdNumber { get; protected set; }
        public string Name { get; protected set; }
        public Dictionary<Jobs, float> PayRates { get; private set; } = new();
        public Dictionary<Jobs, float> Raises { get; private set; } = new();
        public List<Shift> Shifts = new();
        public float[] OverTimeHours = new float[3];
        public bool IsSalaried;
        public bool IsGrandForksEmployee;
        public string SocialSecurityNumber;
        public string EmploymentCategory;
        public string PhoneNumber;
        public bool WasCreatedFromEmployeeExport;
        public bool ShouldBeConsideredForRaises; //means they have been confirmed to have hours in Timesheets.xlsx
        public bool HasADirectDepositAccount;
        public int YearsOfService;
        public Dictionary<string/*job code*/, Dictionary<int/*week num*/, Shift>>[,] ShiftTotals = new Dictionary<string/*job code*/, Dictionary<int/*week num*/, Shift>>[2/*company*/,3/*0-has hours,1-has dollars,2-has both*/];



        public Employee(int idNumber, string name)
        {
            this.IdNumber = idNumber;
            this.Name = name;
        }

        public void SetPayRate(Jobs job, float rate)
        {
            PayRates[job] = Math.Max(PayRates.GetValueOrDefault(job, 0f), rate);
        }

        public float GetDriverRateForSchoolRouteShift(Shift shift)
        {
            if (shift.JobType != Jobs.DRIVER_SCHOOL && shift.JobType != Jobs.NON_CDL_DRIVER)
            {
                Log("Trying to get driver rate for school route shift for shift.jobtype == " + shift.JobType, true);
            }
            float rate = 0f;
            if (shift.JobType == Jobs.NON_CDL_DRIVER)
            {
                if (PayRates.ContainsKey(Jobs.DRIVER_SCHOOL))
                {
                    DelayedLog("Problem in GetDriverRateForSchoolRouteShift()", true);
                }

                //aides don't get downgraded for driving a non-cdl route.
                rate = Math.Max(PayRates.GetValueOrDefault(Jobs.AIDE_SCHOOL, 0f), IsGrandForksEmployee || shift.IsAGrandForksShift ? GrandForksDefaultRates[shift.JobType] : FargoDefaultRates[shift.JobType]);
            }
            else
            {
                if ((IsGrandForksEmployee || shift.IsAGrandForksShift) && PayRates.ContainsKey(Jobs.DRIVER_SCHOOL) && PayRates[Jobs.DRIVER_SCHOOL] < GrandForksDefaultRates[Jobs.DRIVER_SCHOOL])
                {
                    rate = GrandForksDefaultRates[Jobs.DRIVER_SCHOOL];
                }
                rate = Math.Max(rate, PayRates.GetValueOrDefault(shift.JobType, 0f));
            }
            
            return Math.Max(Math.Max(PayRates.GetValueOrDefault(Jobs.MECHANIC, 0f), PayRates.GetValueOrDefault(Jobs.WASH_BAY)), rate);
        }

        public List<Shift> SchoolRouteShifts()
        {
            return Shifts.FindAll(shift => shift.IsASchoolRouteShift());
        }

        public List<Shift> NonSchoolRouteShiftsWithAPotentialMinimumGuarantee()
        {
            return Shifts.FindAll(shift => !shift.IsASchoolRouteShift() && shift.GetMinimumGuaranteeMax(this) > 0f);
        }

        public Jobs IsADriverOrAnAide()
        {
            foreach (var shift in Shifts) 
            {
                if (shift.JobType == Jobs.DRIVER_SCHOOL)
                {
                    return Jobs.DRIVER_SCHOOL;
                }
            }

            foreach (var shift in Shifts)
            {
                if (shift.JobType == Jobs.AIDE_SCHOOL)
                {
                    return Jobs.AIDE_SCHOOL;
                }
            }

            DelayedLog("Warning: Couldn't determine if " + Name + " is a driver or an aide.", true);
            return Jobs.DRIVER_SCHOOL;
        }

        //only use this for weekly MG excpetions - otherwise make sure it will work properly if used for another purpose.
        public Shift FindDriverOrAideShiftForWeek(int week, Jobs jobType)
        {
            for (int shiftType = 0; shiftType <= (int)Type.BOTH; ++ shiftType)
            {
                if (null != ShiftTotals[(int)Company.VALLEY_BUS_LLC, shiftType])
                {
                    foreach (var entry in ShiftTotals[(int)Company.VALLEY_BUS_LLC, shiftType].Values)
                    {
                        foreach (Shift shift in entry.Values)
                        {
                            if (shift.WeekNumber == week && shift.JobType == jobType)
                            {
                                return shift;
                            }
                        }
                    }
                }
            }

            {//c# scope bs
                Shift shift = new(Company.VALLEY_BUS_LLC);
                Shifts.Add(shift);
                if (!ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC].ContainsKey(Shift.GetLaborCode(jobType, false)))
                {
                    ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC].Add(Shift.GetLaborCode(jobType, false), new());
                }
                if (!ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC][Shift.GetLaborCode(jobType, false)].ContainsKey(week))
                {
                    ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC][Shift.GetLaborCode(jobType, false)].Add(week, shift);
                }
                else
                {
                    Log("Error: How was shift not found above?", true);
                }

                shift.WeekNumber = week;
                shift.JobType = jobType;

                if (!PayRates.ContainsKey(jobType))
                {
                    DelayedLog("Check " + Name + " to ensure they are correctly categorized as a driver or aide. Maybe they are a non-cdl driver?");
                }

                return shift;
            }
        }
    }

    enum Exceptions
    {
        BURINGRUD, 
    }

}
