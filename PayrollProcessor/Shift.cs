using DocumentFormat.OpenXml.Spreadsheet;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    public class Shift
    {
        public const int WEST_FARGO_BUS_PLACE_HOLDER = int.MaxValue;
        private static readonly int[] BigBusNumbers = new int[] { WEST_FARGO_BUS_PLACE_HOLDER, 14, 26, 28, 29, 32, 33, 37, 38, 39, 40, 41, 43, 44, 45, 46, 48, 49, 52, 53, 55, 56, 57, 58, 59, 60, 61, 63, 64, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 100, 101, 105, 109, 111, 113, 301, 302, 306, 308, 309, 310, 311, 312, 313, 314, 315, 317, 321, 322, 329, 330, 333 };
        private static readonly int[] SpedBusNumbers = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 30, 31, 34, 35, 36, 42, 47, 50, 51, 54, 62, 65, 99, 102, 103, 104, 106, 107, 108, 110, 112, 114, 115, 116, 117, 118, 119, 120, 502, 503, 504, 505, 303, 304, 305, 307, 316, 318, 319, 320, 323, 324, 325, 326, 327, 328, 331, 332 };
        private const int TJ_MAX_BUS = 799;
        private const int TJ_MIN_BUS = 700;
        private const int BusStartingDailyBonus = 10;
        public static int ShiftCounter = 0;
        public static TimeSpan WORK_DAY_BEGIN = new TimeSpan(5, 30, 0);
        public static TimeSpan WORK_DAY_END = new TimeSpan(17, 0, 0);
        public static int[/*location*/,/*day*/] DailySchoolRouteCounter = new int[4/*location*/,32/*day*/];

        public float ShiftTime;
        public float? PayRate;
        //public float Overtime;
        public float MinimumGuaranteeHours;
        public float SummerGuaranteeHours;
        public Jobs JobType;
        public float DollarAmount;
        public float BonusDollars;
        public float PerDiem;
        public int BusNumber;
        public string? Notes;
        public DateTime Date;
        public TimeSpan ClockIn;
        public TimeSpan ClockOut;
        public int WeekNumber;
        public bool IsAGrandForksShift;
        public bool IsATotalsShift = false;
        public Company CompanyName;
        public int ShiftId;
        public Location ShiftLocation; //WARNING: Be wary of using location for any shift that isn't a driver shift. 
        public bool ExtrasWereWrittenToExport = false;
        public int JobInt;
        public int CoachTripDays = 1;

        public Shift()
        {
            ShiftId = ShiftCounter++;
        }

        public Shift(Company companyName) : this()
        {
            //Log("ShiftCounter == " + ShiftCounter);
            CompanyName = companyName;
        }
        public Shift(Company companyName, Jobs jobType) : this(companyName)
        {
            JobType = jobType;
        }

        public DateTime GetDate()
        {
            return Date;
        }

        public bool IsValid(Employee emp)
        {
            bool isValid = ShiftTime + DollarAmount + MinimumGuaranteeHours + PerDiem + SummerGuaranteeHours + BonusDollars > 0;
            if (!isValid)
            {
                if (ClockIn.CompareTo(ClockOut) == 0)
                {
                    //this shift was deemed to be in valid due to schedule conflict.
                    return false;
                }
                Log("Shift: " + this.ToString() + " is not valid. Please investigate.");
                return false;
            }

            if (emp.IsSalaried)
            {
                if (emp.IdNumber !=  1991 && (JobType == Jobs.DRIVER_SCHOOL || JobType == Jobs.AIDE_SCHOOL || JobType == Jobs.NON_CDL_DRIVER))
                {
                    return false;
                }

                if (!IsATotalsShift)
                {
                    isValid = ClockIn.CompareTo(WORK_DAY_BEGIN) < 0 || ClockOut.CompareTo(WORK_DAY_END) > 0 || Date.DayOfWeek == DayOfWeek.Saturday || Date.DayOfWeek == DayOfWeek.Sunday;
                }
            }
            return isValid;
        }

        public float WorkingHours()
        {
            if (JobType == Jobs.VACATION || JobType == Jobs.HOLIDAY)
            {
                return 0f;
            }
            return ShiftTime;
        }

        public float AllHours(bool bIncludeEstimatedHoursFromCoachShifts)
        {
            float time = ShiftTime + MinimumGuaranteeHours + SummerGuaranteeHours;
            if (time < 0.01f && bIncludeEstimatedHoursFromCoachShifts && DollarAmount > 0 && JobType == Jobs.DRIVER_COACH)
            {
                time += DollarAmount / COACH_HOURLY_RATE_ESTIMATE;
            }
            return time;
        }

        public float TotalCompensation(Employee employee)
        {
            if (!IsValid(employee))
            {
                return 0f;
            }
            
            if ((null == PayRate || PayRate < 1) && DollarAmount < 0.01f && PerDiem < 0.01f && BonusDollars < 0.01f)
            {
                Log("Issue with shift.TotalCompensation(). It is likely being called before the payrate has been calculated for the shift.", true);
            }

            float dollarAmountLocal = DollarAmount + BonusDollars;
            if (dollarAmountLocal < 0.01f && null != PayRate)
            {
                dollarAmountLocal += AllHours(false) * PayRate.Value;
            }
            return dollarAmountLocal;
        }

        public float GetMinimumGuaranteeMax(Employee employee, out MgSource sourceOfMg, List<Shift>? shiftsInRouteTimeContext = null)
        {
            sourceOfMg = MgSource.NONE;
            if (null != Notes && (StringSearch(Notes, "no min") || StringSearch(Notes, "nomin") || StringSearch(Notes, "no minimum") || StringSearch(Notes, "tnt") || StringSearch(Notes, "trolley") || StringSearch(Notes, "training")))
            {
                return 0f;
            }
            else
            {
                if (ClockIn.CompareTo(new TimeSpan(0, 1, 0)) < 0)
                {
                    return 0f; //shift carried over from day before. If there's MG, they would have gotten it for yesterday (technically they would get more than they should).
                }
                if (IsASchoolRouteShift())
                {

                    var shiftTime = shiftsInRouteTimeContext == null ? ShiftTime : shiftsInRouteTimeContext.Sum(shift => shift.ShiftTime);

                    if (null != shiftsInRouteTimeContext && shiftTime < 0.08)
                    {
                        DelayedLog("Giving no minimum guarantee for shift because hours are suspciciously low for " + employee.Name + " on " + Date);
                        return 0f;
                    }

                    if (!Shift.WereThereSchoolRoutesOnThisDay(ShiftLocation, Date.Day))
                    {
                        HashSet<int> temporaryExceptions = new HashSet<int>() { 2061, 2628 };
                        if (temporaryExceptions.Contains(employee.IdNumber))
                        {
                            Log("Giving MG to " + employee.Name + " (" + employee.IdNumber + ") even though it doesn't seem like there were routes that day because they are listed as a temporary exception.");
                        }
                        else
                        {
                            Log("No mg for " + employee.Name + " (" + employee.IdNumber + ") because it was determined that there were no school shifts happening in " + ShiftLocation + " on " + this.Date.DayOfWeek + " at " + this.Date.ToShortTimeString());
                            return 0f;
                        }
                    }

                    foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.SmallMgExceptions)
                    {
                        if (entry != null && entry.IdNumber == employee.IdNumber && entry.BusNumber == BusNumber)
                        {
                            sourceOfMg = MgSource.SPECIAL_EXCEPTION;
                            return entry.Hours;
                        }
                    }

                    if (null != shiftsInRouteTimeContext && shiftTime < 0.2)
                    {
                        DelayedLog("Giving no minimum guarantee for shift because hours are suspciciously low for " + employee.Name + " on " + Date);
                        return 0f;
                    }

                    //standard mg 
                    sourceOfMg = MgSource.BUS_TYPE;
                    float maxMg = 1.5f;
                    if (BigBusNumbers.Contains(BusNumber) || SpedBusNumbers.Contains(BusNumber))
                    {
                        if (BigBusNumbers.Contains(BusNumber) && SpedBusNumbers.Contains(BusNumber))
                        {
                            Log("Problem with bus numbers. Bus #" + BusNumber.ToString() + " is included in both small and big bus numbers.", true);
                        }
                        maxMg = BigBusNumbers.Contains(BusNumber) ? 2f : 1.5f;
                    }
                    else
                    {
                        //if we couldn't find the bus in big buses or in speds (or if there is no bus number), check if the person is a big bus driver, in which case the driver should probably be getting a 2 hour minimum.
                        if (JobType == Jobs.DRIVER_SCHOOL)
                        {
                            foreach (var shift in employee.Shifts)
                            {
                                if (shift.BusNumber != 0)
                                {
                                    if (BigBusNumbers.Contains(shift.BusNumber))
                                    {
                                        maxMg = 2f;
                                        break;
                                    }
                                }
                            }
                        }
                        if (BusNumber != 0 && BusNumber < 300)
                        {
                            Log("Bus not found, bus#" + BusNumber);
                        }
                    }

                    //non standard mg
                    float summerMg = /*ShiftLocation == Location.WEST_FARGO ? 3f :*/ 2f;
                    if (IsASummerRoute() && maxMg < summerMg)
                    {
                        sourceOfMg = MgSource.SUMMER_ROUTE;
                        maxMg = summerMg;
                    }
                    if ((IsAGrandForksShift || employee.IsAGrandForksEmployee) && maxMg < 2)
                    {
                        sourceOfMg = MgSource.GRAND_FORKS_ROUTE;
                        maxMg = 2f;
                    }

                    //special mg
                    SpecialEmployeeHandler.GetInstance().CheckForMgExceptionForShift(employee, this, out float specialMg);
                    if (specialMg > maxMg)
                    {
                        sourceOfMg = MgSource.SPECIAL_EXCEPTION;
                        maxMg = specialMg;
                    }

                    return maxMg;
                }
                else if (JobIsCharter(JobType))
                {
                    sourceOfMg = MgSource.STANDARD_CHARTER;
                    if (JobType == Jobs.DRIVER_OUT_OF_TOWN_CHARTER)
                    {
                        return OUT_OF_TOWN_MIN_GUARANTEE_DRIVER_IN_DOLLARS / CalculateCharterRate(employee);
                    }
                    if (StringSearch(Notes, "Hock"))
                    {
                        if (StringSearch(Notes, "Hockey"))
                        {
                            float payRate = employee.GetPayRateForShift(this);
                            if (StringSearch(Notes, "Band") || StringSearch(Notes, "120"))
                            {
                                return (float)Math.Round(GF_HOCKEY_BAND_PAY / payRate, 2);
                            }
                            else
                            {
                                return (float)Math.Round(GF_HOCKEY_PAY / payRate, 2);
                            }
                        }
                        else
                        {
                            Log("Found 'Hock', but not 'Hockey'. Typo?", true);
                        }
                    }

                    if ((null != Notes && StringSearch(Notes, "private")) || (BusNumber >= TJ_MIN_BUS && BusNumber <= TJ_MAX_BUS))
                    {
                        sourceOfMg = BusNumber >= TJ_MIN_BUS && BusNumber <= TJ_MAX_BUS ? MgSource.T_AND_J_CHARTER : MgSource.PRIVATE_CHARTER;
                        return T_AND_J_CHARTERS_MG_IN_DOLLARS / CalculateCharterRate(employee);
                    }
                    else if (JobType == Jobs.DRIVER_CHARTER_PRIVATE)
                    {
                        return PRIVATE_CHARTER_MIN_GUARANTEE_DRIVER_IN_DOLLARS / CalculateCharterRate(employee);
                    }
                    else if (Date.DayOfWeek == DayOfWeek.Saturday || Date.DayOfWeek == DayOfWeek.Sunday)
                    {
                        sourceOfMg = MgSource.WEEKEND_CHARTER;
                        float weekendMinimum = JobType == Jobs.AIDE_CHARTER ? OUT_OF_TOWN_OR_WEEKEND_MIN_GUARANTEE_AIDE_IN_DOLLARS : WEEKEND_MIN_GUARANTEE_DRIVER_IN_DOLLARS;
                        return weekendMinimum / CalculateCharterRate(employee);
                    }
                    else
                    {
                        return 3f;
                    }
                }
            }
            return 0f;
        }

        public bool IsASpedBusShift()
        {
            return SpedBusNumbers.Contains(this.BusNumber);
        }

        static bool TAndJMessageWasDisplayed = false;
        private float CalculateCharterRate(Employee employee)
        {
            if (JobType == Jobs.AIDE_CHARTER)
            {
                return employee.IsAGrandForksEmployee || IsAGrandForksShift ? GrandForksDefaultRates[Jobs.AIDE_CHARTER] : FargoDefaultRates[Jobs.AIDE_CHARTER];
            }

            if (BusNumber >= TJ_MIN_BUS && BusNumber <= TJ_MAX_BUS && !TAndJMessageWasDisplayed)
            {
                Log("Attention: There is a shift in a T&J bus. I thought all T&J was supposed to make $19.00/hr (aka Sarah would be putting their shifts on the coaches sheet and they wouldn't be clocking-in).", true);
                TAndJMessageWasDisplayed = true;
            }

            if ((null != Notes && StringSearch(Notes, "private")) || (BusNumber >= TJ_MIN_BUS && BusNumber <= TJ_MAX_BUS))
            {
                return Math.Max(employee.PayRates.GetValueOrDefault(JobType, 0f), T_AND_J_CHARTER_RATE);
            }

            if (JobType == Jobs.DRIVER_CHARTER_PRIVATE)
            {
                return Math.Max(employee.PayRates.GetValueOrDefault(JobType, 0f), PRIVATE_CHARTER_RATE);
            }

            if (JobType == Jobs.DRIVER_OUT_OF_TOWN_CHARTER)
            {
                return Math.Max(employee.PayRates.GetValueOrDefault(JobType, 0f), OUT_OF_TOWN_CHARTER_RATE);
            }

            return employee.PayRates.GetValueOrDefault(JobType, 0f);
        }

        public static string GetLaborCode(Jobs jobType, bool isOvertime)
        {
            if (isOvertime)
            {
                return "OT";
            }
            //"Wash BayOT"
            switch (jobType)
            {
                case Jobs.DRIVER_CHARTER_PRIVATE:
                case Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS:
                case Jobs.DRIVER_OUT_OF_TOWN_CHARTER:
                    return "DrvrSchool";
                case Jobs.DRIVER_SCHOOL:
                case Jobs.NON_CDL_DRIVER:
                    return "DrvrDlySch";
                case Jobs.MECHANIC:
                    return "Mechanic";
                case Jobs.WASH_BAY:
                    return "Wash Bay";
                case Jobs.WASH_BAY_OT:
                    return "Wash BayOT";
                case Jobs.TRAINING:
                    return "Training";
                case Jobs.BODY_SHOP:
                    return "Body Shop";
                case Jobs.ADMIN:
                    return "Admin";
                case Jobs.CLEANING:
                    return "Cleaning";
                case Jobs.HOLIDAY:
                    return "MechHolida";
                case Jobs.VACATION:
                    return "MechVaca";
                case Jobs.AIDE_CHARTER:
                    return "AidDlyChrt";
                case Jobs.AIDE_SCHOOL:
                    return "AidDlyScho";
                case Jobs.DRIVER_COACH:
                    return "Driver Coa";
            }
            DelayedLog("Failed to find labor code for " + jobType.ToString(), true);
            return "";
        }

        public static string GetDepartmentCode(Jobs jobType)
        {
            switch (jobType)
            {
                case Jobs.DRIVER_SCHOOL:
                    return "000001";
                case Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS:
                case Jobs.DRIVER_CHARTER_PRIVATE:
                case Jobs.DRIVER_OUT_OF_TOWN_CHARTER:
                    return "000002";
                case Jobs.MECHANIC:
                    return "000007";
                case Jobs.WASH_BAY:
                    return "000009";
                case Jobs.WASH_BAY_OT:
                    return "000010";
                case Jobs.TRAINING:
                    return "000011";
                case Jobs.BODY_SHOP:
                    return "000012";
                case Jobs.ADMIN:
                    return "000013";
                case Jobs.CLEANING:
                    return "000014";
                case Jobs.HOLIDAY:
                    return "000015";
                case Jobs.VACATION:
                    return "000016";
                case Jobs.AIDE_CHARTER:
                    return "000024";
                case Jobs.AIDE_SCHOOL:
                    return "000025";
                case Jobs.DRIVER_COACH:
                    return "000026";
                case Jobs.NON_CDL_DRIVER:
                    return "000028";
            }
            DelayedLog("Failed to find labor code for " + jobType.ToString(), true);
            return "";
        }

        public void CheckWashbayOT(Employee employee)
        {
            if (JobType == Jobs.WASH_BAY_OT)
            {
                if (!this.Date.DayOfWeek.Equals(DayOfWeek.Saturday) && !this.Date.DayOfWeek.Equals(DayOfWeek.Sunday))
                {
                    DelayedLog("Changing wash bay OT shift to washbay shift for " + employee.Name + " on " + this.Date.DayOfWeek.ToString());
                    JobType = Jobs.WASH_BAY;
                }
            }
        }

        public string GetLaborCode(bool isOvertime)
        {
            return GetLaborCode(JobType, isOvertime);
        }

        public bool IsASummerRoute()
        {
            return IsSummerDate(this.Date, this.ShiftLocation);
        }

        public RouteTimeContext TimeContext()
        {
            if (Date.TimeOfDay.CompareTo(new TimeSpan(0, 3, 0)) < 0)
            {
                if (JobType != Jobs.DRIVER_OUT_OF_TOWN_CHARTER)
                {
                    Log("Warning: Trying to get TimeContext for shift with TimeOfDay: " + Date.TimeOfDay.ToString(), true);
                }
            }
            if (Date.TimeOfDay.CompareTo(new TimeSpan(9, 10, 0)) <= 0)
            {
                return PayrollProcessor.RouteTimeContext.MORNING;
            }
            return Date.TimeOfDay.CompareTo(new TimeSpan(12, 30, 0)) <= 0 ? PayrollProcessor.RouteTimeContext.NOON : PayrollProcessor.RouteTimeContext.AFTERNOON;
        }

        public bool IsASchoolRouteShift()
        {
            return JobType == Jobs.DRIVER_SCHOOL || JobType == Jobs.AIDE_SCHOOL || JobType == Jobs.NON_CDL_DRIVER;
        }

        private static Dictionary<Employee, List<Jobs>> PayrateMessages = new();
        public float SpecialRate(Employee emp)
        {
            float specialRate = 0f;

            switch (JobType)
            {
                case Jobs.DRIVER_SCHOOL:
                case Jobs.NON_CDL_DRIVER:
                    return emp.GetDriverRateForSchoolRouteShift(this);
                case Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS:
                case Jobs.DRIVER_CHARTER_PRIVATE:
                case Jobs.AIDE_CHARTER:
                case Jobs.DRIVER_OUT_OF_TOWN_CHARTER:
                    return Math.Max(specialRate, CalculateCharterRate(emp));
                case Jobs.WASH_BAY_OT:
                    if (emp.PayRates.ContainsKey(Jobs.WASH_BAY))
                    {
                        return emp.PayRates[Jobs.WASH_BAY] * 1.5f;
                    }
                    Log("ERROR: Employee using washbay OT but they don't have a washbay rate. Using starting washbay rate.", true);
                    float STARTING_WASH_BAY_RATE = emp.IsAGrandForksEmployee ? GrandForksDefaultRates[Jobs.AIDE_SCHOOL] : FargoDefaultRates[Jobs.AIDE_SCHOOL];
                    return STARTING_WASH_BAY_RATE * 1.5f;
                case Jobs.HOLIDAY:
                case Jobs.VACATION:
                    return emp.PayRates.Values.Max();
                default:
                    if (IsAGrandForksShift && GrandForksDefaultRates.ContainsKey(JobType) && GrandForksDefaultRates[JobType] != FargoDefaultRates.GetValueOrDefault(JobType, 0f))
                    {
                        specialRate = GrandForksDefaultRates[JobType];
                        if (!emp.IsAGrandForksEmployee && emp.PayRates.ContainsKey(JobType) && FargoDefaultRates.ContainsKey(JobType))
                        {
                            var preUpdate = specialRate;
                            specialRate += emp.PayRates[JobType] - FargoDefaultRates[JobType];
                            if (specialRate > preUpdate)
                            {
                                Log("Upgraded special GF payrate for " + JobType.ToString() + " from " + preUpdate + " to " + specialRate);
                            }
                            else if (emp.YearsOfService > 0)
                            {
                                Log("Expected there to be a rate increase for special GF rate for " + emp.Name + " because years of service == " + emp.YearsOfService + " but there was no increase.");
                            }
                        }
                    }
                    else if (!IsAGrandForksShift && FargoDefaultRates.ContainsKey(JobType))
                    {
                        specialRate = FargoDefaultRates[JobType];
                    }
                    break;
            }

            return specialRate; //could be less than their default rate here, and that's fine.
        }

        public void ModifyClockIn(TimeSpan newClockIn)
        {
            ClockIn = newClockIn;
            ShiftTime = (float)ClockOut.Subtract(ClockIn).TotalHours;
        }

        public void ModifyClockOut(TimeSpan newClockOut)
        {
            ClockOut = newClockOut;
            ShiftTime = (float)ClockOut.Subtract(ClockIn).TotalHours;
        }

        public void AddAll(Shift shift)
        {
            ShiftTime += shift.ShiftTime;
            MinimumGuaranteeHours += shift.MinimumGuaranteeHours;
            SummerGuaranteeHours += shift.SummerGuaranteeHours;
            DollarAmount += shift.DollarAmount;
            BonusDollars += shift.BonusDollars;
            PerDiem += shift.PerDiem;
            JobType = shift.JobType;
            WeekNumber = shift.WeekNumber;
            CompanyName = shift.CompanyName;
        }

        public Type Type()
        {
            return ShiftTime > 0 ? PayrollProcessor.Type.HOURS : PayrollProcessor.Type.DOLLAR_AMOUNT;
        }

        public static bool WereThereSchoolRoutesOnThisDay(Location location, int dayNumber)
        {
            return DailySchoolRouteCounter[(int)location, (int)dayNumber] > 4
                || (location == Location.GRAND_FORKS && DailySchoolRouteCounter[(int)location, (int)dayNumber] > 2)
                ;
        }
    }

    public enum MgSource
    {
        BUS_TYPE, GRAND_FORKS_ROUTE, SUMMER_ROUTE, STANDARD_CHARTER, PRIVATE_CHARTER, T_AND_J_CHARTER, WEEKEND_CHARTER, SPECIAL_EXCEPTION, NONE
    }

    public enum RouteTimeContext
    {
        MORNING, NOON, AFTERNOON
    }

    public enum Company
    {
        VALLEY_BUS_LLC, VALLEY_BUS_COACHES
    }
    public enum Location
    {
        FARGO, WEST_FARGO, GRAND_FORKS
    }

    public enum Type
    {
        HOURS, DOLLAR_AMOUNT
    }
}