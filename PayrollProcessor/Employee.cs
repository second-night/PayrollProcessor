using DocumentFormat.OpenXml.Spreadsheet;
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
        public float AnnualSalaryAmount;
        public bool IsGrandForksEmployee;
        public string SocialSecurityNumber;
        public DateTime HireDate = DateTime.Now;
        public string EmploymentCategory;
        public string PhoneNumber;
        public bool WasCreatedFromEmployeeExport;
        public bool NeedsUpdateInPayroll;
        public bool HadHoursInTimesheets; //means they have been confirmed to have hours in Timesheets.xlsx
        public bool HasAnActiveDirectDepositAccount;
        public bool HasAnyDirectDepositAccount;
        public bool IsMale;
        public bool IsAMechanicApprentice;
        public int YearsOfService;
        public bool WasAlreadyInPayroll;
        public bool IsPartialEntry = false;
        public DateTime BirthDate;
        public bool IsTerminated;
        public bool WasReportedForPartialEntry;
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
                    if (PayRates.ContainsKey((Jobs)entry.OverridingJobType))
                    {
                        return PayRates[(Jobs)entry.OverridingJobType];
                    }
                    else if ((Jobs)entry.OverridingJobType == Jobs.NON_CDL_DRIVER)
                    {
                        return NonCdlRate(shift.IsAGrandForksShift);
                    }
                    else
                    {
                        Log("Problem finding PayRate: " + ((Jobs)entry.OverridingJobType).ToString() + " for employee " + Name, true);
                    }
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

            if (shift.ShiftTime == 0 && shift.ShiftTime > 0.05f)
            {
                DelayedLog("Warninig: Cannot determine a payrate for Employee " + Name + " ( " + IdNumber + " ) for jobType: " + shift.JobType.ToString());
            }
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
                rate = NonCdlRate(shift.IsAGrandForksShift);
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

        private float NonCdlRate(bool bIsForGrandForks)
        {
            float baseRate = GetBasePayRateForEmployee(Jobs.NON_CDL_DRIVER, this, bIsForGrandForks);
            float newRate = TimeInServiceAdjustment(baseRate, this, Jobs.NON_CDL_DRIVER, true);
            float paraRate = GetBasePayRateForEmployee(Jobs.AIDE_SCHOOL, this, bIsForGrandForks);
            paraRate = TimeInServiceAdjustment(paraRate, this, Jobs.AIDE_SCHOOL, false);

            if (paraRate > newRate)
            {
                Log("Possible issue where paraRate > nonCdl rate.");
            }

            return Math.Max(newRate, paraRate);
        }

        public bool IsANonCdlDriver()
        {
            float baseRate = GetBasePayRateForEmployee(Jobs.DRIVER_SCHOOL, this, IsGrandForksEmployee);
            float newRate = TimeInServiceAdjustment(baseRate, this, Jobs.DRIVER_SCHOOL, true);
            return !PayRates.ContainsKey(Jobs.DRIVER_SCHOOL) && (!PayRates.ContainsKey(Jobs.MECHANIC) || PayRates[Jobs.MECHANIC] < newRate);
        }

        public List<Shift> SchoolRouteShifts()
        {
            return Shifts.FindAll(shift => shift.IsASchoolRouteShift());
        }

        public List<Shift> NonSchoolRouteShiftsWithAPotentialMinimumGuarantee()
        {
            return Shifts.FindAll(shift => !shift.IsASchoolRouteShift() && shift.GetMinimumGuaranteeMax(this, out _) > 0f);
        }

        public Jobs PrimaryJobType()
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

            foreach (var shift in Shifts)
            {
                if (shift.JobType == Jobs.WASH_BAY)
                {
                    return Jobs.WASH_BAY;
                }
            }

            foreach (var shift in Shifts)
            {
                if (shift.JobType == Jobs.ADMIN)
                {
                    return Jobs.ADMIN;
                }
            }

            foreach (var shift in Shifts)
            {
                if (shift.JobType == Jobs.MECHANIC)
                {
                    return Jobs.MECHANIC;
                }
            }

            float adminRate = PayRates.GetValueOrDefault(Jobs.ADMIN, 0f);
            float mechanicRate = PayRates.GetValueOrDefault(Jobs.MECHANIC, 0f);
            if (adminRate + mechanicRate > 0.001f)
            {
                return adminRate > mechanicRate ? Jobs.ADMIN : Jobs.MECHANIC;
            }
            

            DelayedLog("Warning: Couldn't determine primary job type for " + Name + ".", true);
            return Jobs.DRIVER_SCHOOL;
        }

        //only use this for weekly MG excpetions - otherwise make sure it will work properly if used for another purpose.
        public Shift? FindShiftForWeek(int week, Jobs jobType, bool bShouldCreateNewShiftIfShiftIsNotFound)
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

            if (!bShouldCreateNewShiftIfShiftIsNotFound)
            {
                return null;
            }

            //didn't find shift, make new shift
            {//c# scope bs
                Shift shift = new(Company.VALLEY_BUS_LLC);
                Shifts.Add(shift);
                if (null == ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC])
                {
                    ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC] = new();
                }
                if (!ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC].ContainsKey(Shift.GetLaborCode(jobType, false)))
                {
                    ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC].Add(Shift.GetLaborCode(jobType, false), new());
                }
                if (!ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC][Shift.GetLaborCode(jobType, false)].ContainsKey(week))
                {
                    ShiftTotals[(int)Type.HOURS, (int)Company.VALLEY_BUS_LLC][Shift.GetLaborCode(jobType, false)][week] = new()
                    {
                        shift
                    };
                }
                else
                {
                    Log("Error: How was shift not found above?", true);
                }

                shift.WeekNumber = week;
                shift.JobType = jobType;

                if (!PayRates.ContainsKey(jobType))
                {
                    DelayedLog("Couldn't find payrate for " + Name + " When creating a new shift for special MG.");
                }
                else
                {
                    shift.PayRate = PayRates[jobType];
                }

                return shift;
            }
        }
    }
}
