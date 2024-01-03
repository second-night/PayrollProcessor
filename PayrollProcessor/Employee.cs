using System.Data;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class Employee
    {

        public int IdNumber { get; protected set; }
        public string Name { get; protected set; }
        public Dictionary<Jobs, float> PayRates { get; private set; } = new();
        public List<Shift> Shifts = new();
        public float[] OverTimeHours = new float[3];
        public bool IsSalaried;
        public bool IsGrandForksEmployee;
        public string SocialSecurityNumber;
        public string EmploymentCategory;
        public string PhoneNumber;
        public bool WasCreatedFromEmployeeExport;
        public bool NeedsUpdateInPayroll;
        public bool HadHoursInTimesheets; //means they have been confirmed to have hours in Timesheets.xlsx
        public bool HasADirectDepositAccount;
        public bool IsMale;
        public bool IsAMechanicApprentice;
        public int YearsOfService;
        public bool WasAlreadyInPayroll;
        public Dictionary<string/*job code*/, Dictionary<int/*week num*/, List<Shift>>>[,] ShiftTotals = new Dictionary<string/*job code*/, Dictionary<int/*week num*/, List<Shift>>>[2/*company*/,3/*0-has hours,1-has dollars,2-has both*/];



        public Employee(int idNumber, string name)
        {
            this.IdNumber = idNumber;
            this.Name = name;
        }

        public void SetPayRate(Jobs job, float rate)
        {
            PayRates[job] = Math.Max(PayRates.GetValueOrDefault(job, 0f), rate);
        }

        public float GetPayRateForShift(Shift shift)
        {
            foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.PayRateExceptions)
            {
                if (entry.IdNumber == IdNumber && (Jobs)entry.JobType == shift.JobType)
                {
                    return entry.Rate;
                }
            }
            foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.PayRateSubstitutionExceptions)
            {
                if (entry.IdNumber == IdNumber && (Jobs)entry.OverriddenJobType == shift.JobType)
                {
                    foreach (var entry2 in SpecialEmployeeHandler.GetInstance().SpecialEmployees.PayRateExceptions)
                    {
                        if (entry2.IdNumber == IdNumber && entry2.JobType == entry.OverridingJobType)
                        {
                            return entry2.Rate;
                        }
                    }
                    return PayRates[(Jobs)entry.OverridingJobType];
                }
            }

            if (shift.JobType == Jobs.DRIVER_SCHOOL || shift.JobType == Jobs.NON_CDL_DRIVER)
            {
                return GetDriverRateForSchoolRouteShift(shift);
            }

            float specialRate = shift.SpecialRate(this);

            if (PayRates.ContainsKey(shift.JobType))
            {
                return Math.Max(specialRate, PayRates[shift.JobType]);
            }

            if (specialRate > 0.001f)
            {
                return specialRate;
            }

            DelayedLog("Warninig: Cannot determine a payrate for Employee " + Name + " ( " + IdNumber + " ) for jobType: " + shift.JobType.ToString());
            return 0f;
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
                float paraRate = PayRates.GetValueOrDefault(Jobs.AIDE_SCHOOL, 0f);
                float nonCdlRate = IsGrandForksEmployee || shift.IsAGrandForksShift ? GrandForksDefaultRates[shift.JobType] : FargoDefaultRates[shift.JobType];
                rate = Math.Max(paraRate, nonCdlRate);
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
            return Shifts.FindAll(shift => !shift.IsASchoolRouteShift() && shift.GetMinimumGuaranteeMax(this, out _) > 0f);
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
            for (int shiftType = 0; shiftType <= (int)Type.DOLLAR_AMOUNT; ++ shiftType)
            {
                if (null != ShiftTotals[(int)Company.VALLEY_BUS_LLC, shiftType])
                {
                    foreach (var entry in ShiftTotals[(int)Company.VALLEY_BUS_LLC, shiftType].Values)
                    {
                        foreach (List<Shift> shiftList in entry.Values)
                        {
                            foreach (Shift shift in shiftList)
                            {
                                if (shift.WeekNumber == week && shift.JobType == jobType)
                                {
                                    return shift;
                                }
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
                    ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC][Shift.GetLaborCode(jobType, false)][week].Add(shift);
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
