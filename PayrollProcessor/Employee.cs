using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using System.Data;
using System.Windows.Forms.VisualStyles;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class Employee
    {

        public int IdNumber { get; protected set; }
        public string Name { get; protected set; }
        public string FirstName = "";
        public string MiddleName = "";
        public string LastName = "";
        public Dictionary<Jobs, float> PayRates { get; private set; } = new();
        public List<Shift> Shifts = new();
        public float[,] OverTimeHours = new float[2,3]; //[company, week number]
        public bool IsSalaried;
        public float AnnualSalaryAmount;
        public Company PrimaryCompany;
        public HashSet<Company> ActiveCompanies = new();
        public bool IsAGrandForksEmployee;
        public string SocialSecurityNumber;
        public DateTime HireDate = DateTime.Now;
        public string EmploymentCategory;
        public string JobTitleCode = "";
        public string RecommendedJobTitleCode = "";
        public string PhoneNumber;
        public bool NeedsUpdateInPayroll;
        public bool HadHoursInTimesheets; //means they have been confirmed to have hours in Timesheets.xlsx
        public bool HadManualEntry; //means they have been confirmed to have values in manual_entries.xlsx
        public bool HasAnActiveDirectDepositAccount;
        public bool HasAnyDirectDepositAccount;
        public bool IsMale;
        public bool IsAMechanicApprentice;
        public int YearsOfService;
        public bool WasAlreadyInPayroll;
        public bool IsPartialEntry = false;
        public DateTime BirthDate;
        public bool IsTerminated;
        public DateTime TerminationDate;
        public bool WasReportedForPartialEntry;
        public DateTime DateOfDirectDepositUpdateInWorkBright;
        public bool needsDDImported;
        public float VacationHours;
        public List<ManualEntry> ManualEntries = new();
        public Dictionary<string/*job code*/, Dictionary<int/*week num*/, List<Shift>>>[,] ShiftTotals = new Dictionary<string/*job code*/, Dictionary<int/*week num*/, List<Shift>>>[2/*company*/, 3/*0-has hours,1-has dollars,2-has both*/];
        public float NetVacationChangeForPayPeriod;

        //for schedule matching
        //public Dictionary<int/*bus number*/, Dictionary<RouteTimeContext, List<DayOfWeek>>> RoutesByBusNumber = new();
        public Dictionary<DayOfWeek, Dictionary<RouteTimeContext, int>> BusShiftTotals = new();
        public Dictionary<DayOfWeek, Dictionary<RouteTimeContext, Dictionary<int/*BusNumber*/, int/*count*/>>> ShiftsByBusNumber = new();
        public Dictionary<DayOfWeek, Dictionary<RouteTimeContext, TimeSpan>> ScheduleExceptions = new();


        public void fixIdNumber(int newNumber)
        {
            IdNumber = newNumber;
        }

        public Employee(int idNumber, string name)
        {
            this.IdNumber = idNumber;
            this.Name = name;
        }

        public string MiddleInitial
        {
            get
            {
                foreach (char character in MiddleName ?? "")
                {
                    if (char.IsLetter(character))
                    {
                        return char.ToUpperInvariant(character).ToString();
                    }
                }
                return "";
            }
        }

        public void EnsureNameParts()
        {
            if (!string.IsNullOrWhiteSpace(FirstName) && !string.IsNullOrWhiteSpace(LastName))
            {
                return;
            }

            string[] parts = (Name ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
            {
                if (string.IsNullOrWhiteSpace(FirstName))
                {
                    FirstName = parts[0];
                }
            }
            else if (parts.Length == 2)
            {
                if (string.IsNullOrWhiteSpace(FirstName))
                {
                    FirstName = parts[0];
                }
                if (string.IsNullOrWhiteSpace(LastName))
                {
                    LastName = parts[1];
                }
            }
            else if (parts.Length >= 3)
            {
                if (string.IsNullOrWhiteSpace(FirstName))
                {
                    FirstName = parts[0];
                }
                if (string.IsNullOrWhiteSpace(MiddleName))
                {
                    MiddleName = parts[1];
                }
                if (string.IsNullOrWhiteSpace(LastName))
                {
                    LastName = string.Join(" ", parts.Skip(2));
                }
            }
        }

        public float OverTimeHoursForAllCompaniesForWeek(int weekNumber)
        {
            return OverTimeHours[0, weekNumber] + OverTimeHours[1, weekNumber];
        }

        public void SetPayRate(Jobs job, float rate)
        {
            PayRates[job] = Math.Max(PayRates.GetValueOrDefault(job, 0f), rate);
        }

        public bool IsActive()
        {
            bool bIsActive = !IsTerminated;
            if (bIsActive != ActiveCompanies.Count > 0)
            {
                Log("Warning: Employee " + Name + " ( " + IdNumber + " ) has a mismatch between IsTerminated and ActiveCompanies.", true);
            }
            return bIsActive;
        }


        private static Dictionary<Employee, List<Jobs>> PayrateMessages = new();
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
                        Shift tempShift = new(Company.VALLEY_BUS_LLC)
                        {
                            JobType = (Jobs)entry.OverridingJobType,
                            IsAGrandForksShift = shift.IsAGrandForksShift,
                            BusNumber = shift.BusNumber,
                            ShiftTime = shift.ShiftTime,
                            ClockIn = shift.ClockIn,
                            ClockOut = shift.ClockOut,
                            Date = shift.Date
                        };
                        float payRate = GetPayRateForShift(tempShift);

                        //weird exception due to sped rate pay bump
                        if (shift.ShiftTime > 2.5f && shift.IsASpedBusShift() && !shift.IsAGrandForksShift && (Jobs)entry.OverridingJobType == Jobs.DRIVER_SCHOOL)
                        {
                            float payBumpTime = shift.ShiftTime - (shift.ShiftTime > 6f ? 3f : 1.5f); //try to figure out if they drove 1 shift or 2
                            float weightedRate1 = (shift.ShiftTime - payBumpTime) * (payRate - FARGO_SPED_CDL_DRIVER_RATE_BUMP);
                            float weightedRate2 = payBumpTime * payRate;
                            payRate = (weightedRate1 + weightedRate2) / shift.ShiftTime;
                            Log("Applied partial sped rate pay bump for " + Name + " for " + shift.Date.ToShortDateString() + " for " + shift.ShiftTime + " hours. New payrate is " + payRate + ".");
                        }

                        if (payRate < PayRates[(Jobs)entry.OverridingJobType])
                        {
                            Log("Check payrate substitution for " + Name + " because payrate for " + ((Jobs)entry.OverridingJobType).ToString() + " is lower than expected.", true);
                        }
                        return Math.Max(PayRates[(Jobs)entry.OverridingJobType], payRate);
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

            if (shift.JobType == Jobs.DRIVER_CHARTER_PUBLIC)
            {
                float defaultRate = GetBasePayRateForEmployee(Jobs.DRIVER_CHARTER_PUBLIC, this, shift.IsAGrandForksShift);
                if (defaultRate < 19f)
                {
                    Log("Problem finding PayRate for employee " + Name + " for DRIVER_CHARTER_PUBLIC", true);
                }
                return TimeInServiceAdjustment(defaultRate, this, Jobs.DRIVER_CHARTER_PUBLIC, shift.IsAGrandForksShift);
            }
            if (!PayRates.ContainsKey(shift.JobType) &&
                shift.JobType != Jobs.NON_CDL_DRIVER && shift.JobType != Jobs.VACATION && shift.JobType != Jobs.HOLIDAY && shift.JobType != Jobs.WASH_BAY_OT && shift.JobType != Jobs.COACH_PUBLIC_DRIVING && shift.JobType != Jobs.DRIVER_COACH)
            {
                if (SocialSecurityNumber == "" || IsPartialEntry)
                {
                    return 0f;
                }
                if (!PayrateMessages.ContainsKey(this) || !PayrateMessages[this].Contains(shift.JobType))
                {
                    bool bGiveDefault = shift.JobType == Jobs.DRIVER_OUT_OF_TOWN_CHARTER || shift.JobType == Jobs.DRIVER_CHARTER_PUBLIC || shift.JobType == Jobs.AIDE_CHARTER || shift.JobType == Jobs.AIDE_SCHOOL || shift.JobType == Jobs.TRAINING || shift.JobType == Jobs.DRIVER_CHARTER;
                    if (!bGiveDefault)
                    {
                        string specialString = shift.JobType == Jobs.WASH_BAY && IsAGrandForksEmployee ? " (default rate for helping out in washbay in GF is $17.00/hour)." : "";
                        float newRate = PrintForm.InputNumber("Warninig: Employee " + Name + (IsAGrandForksEmployee ? " (GF) " : " (Fargo) ") + " doesn't have a payrate for " + shift.JobType.ToString() + ". Would you like to assign one now?" + specialString + "\nPut '1' for default rate", out string nonNumberInput);
                        if (newRate > 0)
                        {
                            if (newRate == 1) //default
                            {
                                bGiveDefault = true;
                            }
                            else
                            {
                                GiveRaiseToEmployee(this, shift.JobType, newRate);
                            }
                        }
                        else
                        {
                            if (null != nonNumberInput && nonNumberInput != "")
                            {
                                for (int i = 0; i <= (int)Jobs.NON_CDL_DRIVER; ++i)
                                {
                                    if ("" != ((Jobs)i).ToString() && StringSearch(((Jobs)i).ToString(), nonNumberInput))
                                    {
                                        shift.JobType = (Jobs)i;
                                        return GetPayRateForShift(shift);
                                    }
                                }
                            }
                            if (!PayrateMessages.ContainsKey(this))
                            {
                                PayrateMessages[this] = new();
                            }
                            PayrateMessages[this].Add(shift.JobType);
                            DelayedLog("Warninig: Employee " + Name + " ( " + IdNumber + " )" + (IsAGrandForksEmployee ? " (GF) " : " (Fargo) ") + "doesn't have a payrate for " + shift.JobType.ToString());
                        }
                    }

                    if (bGiveDefault)
                    {
                        float newRate = 0f;
                        if (shift.JobType == Jobs.CLEANING || shift.JobType == Jobs.WASH_BAY)
                        {
                            newRate = GetBasePayRateForEmployee(Jobs.AIDE_SCHOOL, this);
                        }
                        else
                        {
                            newRate = GetBasePayRateForEmployee(shift.JobType, this);
                        }
                        if (newRate < 1)
                        {
                            Log("Assigning default rate failed", true);

                        }
                        GiveRaiseToEmployee(this, shift.JobType, newRate);
                    }
                }
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
                if (shift.IsAGrandForksShift && !IsAGrandForksEmployee)
                {
                    if (PayRates.ContainsKey(shift.JobType) && FargoDefaultRates.ContainsKey(shift.JobType))
                    {
                        var preUpdate = PayRates.GetValueOrDefault(shift.JobType, 0f);
                        //float badRate = preUpdate + PayRates[shift.JobType] - FargoDefaultRates[shift.JobType];
                        rate = GrandForksDefaultRates[shift.JobType] + PayRates[shift.JobType] - FargoDefaultRates[shift.JobType];
                        if (rate > preUpdate)
                        {
                            //Log("Upgraded special GF payrate for " + shift.JobType.ToString() + " from " + preUpdate + " to " + rate);
                        }
                        else if (YearsOfService > 0)
                        {
                            Log("Expected there to be a rate increase for special GF rate for " + Name + " because years of service == " + YearsOfService + " but there was no increase.");
                        }
                    }
                }
                rate = Math.Max(rate, PayRates.GetValueOrDefault(shift.JobType, 0f));
                if (!shift.IsAGrandForksShift && shift.IsASpedBusShift())
                {
                    rate += FARGO_SPED_CDL_DRIVER_RATE_BUMP;
                }
            }

            var finalRate = Math.Max(Math.Max(PayRates.GetValueOrDefault(Jobs.MECHANIC, 0f), PayRates.GetValueOrDefault(Jobs.WASH_BAY)), rate);
            if (finalRate < FargoDefaultRates[Jobs.NON_CDL_DRIVER])
            {
                Log("Warning! Final rate for driver shift for " + Name + " == " +  rate);
            }
            return finalRate;
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
            float baseRate = GetBasePayRateForEmployee(Jobs.DRIVER_SCHOOL, this, IsAGrandForksEmployee);
            float newRate = TimeInServiceAdjustment(baseRate, this, Jobs.DRIVER_SCHOOL, true);
            return !PayRates.ContainsKey(Jobs.DRIVER_SCHOOL) && (!PayRates.ContainsKey(Jobs.MECHANIC) || PayRates[Jobs.MECHANIC] < newRate);
        }

        public float VacationHoursUsedThisPayPeriod()
        {
            float total = 0f;
            foreach (var shift in Shifts)
            {
                if (shift.JobType == Jobs.VACATION && !shift.IsATotalsShift)
                {
                    total += shift.ShiftTime;
                }
            }
            return total;
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
        public Shift? FindShiftForWeek(int week, Jobs jobType, Company company, bool bShouldCreateNewShiftIfShiftIsNotFound)
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
                                if (shift.WeekNumber == week && shift.JobType == jobType && shift.CompanyName == company)
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
