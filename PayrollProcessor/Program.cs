using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    //  taskkill /f /im excel.exe

    public static class Program
    {
        public static Dictionary<int, Employee> EmployeeDictionary = new();
        public const float GF_HOCKEY_PAY = 100f;
        public const float GF_HOCKEY_BAND_PAY = 120f;
        public const float T_AND_J_CHARTERS_MG_IN_DOLLARS = 120f;
        public const float OUT_OF_TOWN_MIN_GUARANTEE_DRIVER_IN_DOLLARS = 120f;
        public const float PRIVATE_CHARTER_MIN_GUARANTEE_DRIVER_IN_DOLLARS = 120f;
        public const float WEEKEND_MIN_GUARANTEE_DRIVER_IN_DOLLARS = 70f;
        public const float OUT_OF_TOWN_OR_WEEKEND_MIN_GUARANTEE_AIDE_IN_DOLLARS = 50f;
        public const float DRIVER_CHARTER_RATE = 18f; 
        public const float OUT_OF_TOWN_CHARTER_RATE = 18.5f;
        public const float PRIVATE_CHARTER_RATE = 19f;
        public const float T_AND_J_CHARTER_RATE = 19.5f; //this shouldn't be used I think, Sarah provides the pay for these drivers.
        public const float TRAINING_RATE= 13f;
        public const float COACH_HOURLY_RATE_ESTIMATE = 20f;
        public const float TEN_YEAR_RATE_BUMP = 0.5f;
        public const float FARGO_SPED_CDL_DRIVER_RATE_BUMP = 0.5f;
        public static string LogString = "";
        public static HashSet<int> BusStartingDays = new();
        private static ExcelWorker ExcelWorker;
        private static bool DoMedhusDeferredPayment;
        private static bool DoJeffShawVacation = true;
        public static List<int> EmployeeIdsToIgnore = new() { 503/*John Mc*/, DoMedhusDeferredPayment ? 1657 : 0/*Bob Medhus*/};

        //fields for logging
        private static Dictionary<MgSource, float> MgSourceTotals = new();
        public static Dictionary<String, bool> DelayedLogMessages = new();
        public static HashSet<Employee> NonCdlDrivers = new();
        private static Dictionary<int, Dictionary<Jobs, float>> ApprenticeMechanicHours = new();

        public static Dictionary<Jobs, float> FargoDefaultRates = new()
        {
            {Jobs.DRIVER_SCHOOL, 22.3f },
            {Jobs.DRIVER_CHARTER, DRIVER_CHARTER_RATE },
            {Jobs.DRIVER_CHARTER_PUBLIC, PRIVATE_CHARTER_RATE },
            {Jobs.COACH_PUBLIC_DRIVING, OUT_OF_TOWN_CHARTER_RATE },
            {Jobs.AIDE_SCHOOL, 18.5f },
            {Jobs.AIDE_CHARTER, 16.5f },
            {Jobs.NON_CDL_DRIVER, 19f },
            {Jobs.TRAINING, TRAINING_RATE }
        };
        public static Dictionary<Jobs, float> GrandForksDefaultRates = new()
        {
            {Jobs.DRIVER_SCHOOL, 23.7f },
            {Jobs.DRIVER_CHARTER, DRIVER_CHARTER_RATE },
            {Jobs.DRIVER_CHARTER_PUBLIC, PRIVATE_CHARTER_RATE },
            {Jobs.AIDE_SCHOOL, 19f },
            {Jobs.AIDE_CHARTER, 18f },
            {Jobs.NON_CDL_DRIVER, 19.7f },
            {Jobs.TRAINING, TRAINING_RATE }
        };

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Console.SetOut(new ToDebugWriter());
            ApplicationConfiguration.Initialize();
            ExcelWorker = new();
            CheckForVacationCutOff(ExcelWorker.FirstDayWeek2);
            ExcelWorker.ReadIsolvedEmployees();
            ExcelWorker.PreCheckTimeSheets();
            ExcelWorker.ReadEmployeeExport();
            DoEmployeeRaises();
            ExcelWorker.ReadTimeSheets();
            ExcelWorker.ReadCoachesPayroll();
            CalculateMinimumGuarantees();
            TotalUpShiftsForEmployees();
            ExcelWorker.WriteEmployeeImports();
            ExcelWorker.WritePayrollImports();
            ExcelWorker.WriteBirthDates();
            ExcelWorker.WriteOverTimeReport();
            FinalLogging();
            //Log("Processed is finished. Have a nice day!", true);
        }

        public class ToDebugWriter : System.IO.StringWriter
        {
            public override void WriteLine(string? value)
            {
                Debug.WriteLine(value);
                base.WriteLine(value);
            }
        }

        public static void Log(string text, bool bShouldDisplayForm = false)
        {
            //System.Diagnostics.Debug.WriteLine(text);
            //Console.Write(text + "\t");
            new ToDebugWriter().WriteLine(text);
            LogString += text + "\n";
            if (bShouldDisplayForm)
            {
                System.Windows.Forms.Application.Run(new PrintForm(text));
            }
        }

        public static void DelayedLog(string text, bool bShouldDisplayForm = false)
        {
            DelayedLogMessages[text] = bShouldDisplayForm;
        }

        public static bool IsSummerDate(DateTime dateTime, Location shiftLocation)
        {
            DateTime SummerStart = new DateTime(DateTime.Now.Year, 6, 1);
            if (shiftLocation == Location.FARGO)
            {
                SummerStart = new DateTime(DateTime.Now.Year, 6, 6);
            }
            return dateTime.CompareTo(SummerStart) > 0 && dateTime.CompareTo(new DateTime(DateTime.Now.Year, 8, 20)) < 0;
        }

        public static bool StringSearch(string? str, string subStr)
        {
            if (str == null)
            {
                return false;
            }
            if (str.Length < subStr.Length)
            {
                return false;
            }
            if (str == subStr)
            {
                return true;
            }

            return str.IndexOf(subStr, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void CalculateMinimumGuarantees()
        {
            int iterationCounter = 0;
            foreach (Employee emp in EmployeeDictionary.Values)
            {
                if (emp.Shifts.Count > 0)
                {
                    emp.Shifts = emp.Shifts.OrderBy(shift => shift.Date).ToList();
                    
                    CalculateMgForSchoolRouteShifts(emp, emp.SchoolRouteShifts());
                    CalculateMgForNonSchoolRouteShifts(emp, emp.NonSchoolRouteShiftsWithAPotentialMinimumGuarantee());
                    if (emp.IdNumber == 1354)
                    {
                        HusseinShallalSpecial(emp);
                    }
                }
            }
        }

        public static void HusseinShallalSpecial(Employee emp)
        {
            List<Shift> newShifts = new();
            foreach (var shift in emp.Shifts)
            {
                if (shift.ShiftTime > 4f && shift.JobType == Jobs.DRIVER_SCHOOL)
                {
                    if (Shift.WereThereSchoolRoutesOnThisDay(Location.FARGO, shift.Date.Day))
                    {
                        Shift newShift = new()
                        {
                            JobType = Jobs.WASH_BAY,
                            ShiftTime = shift.ShiftTime - 4f,
                            CompanyName = Company.VALLEY_BUS_LLC,
                            WeekNumber = shift.WeekNumber
                        };

                        shift.ShiftTime = 4f;

                        newShifts.Add(newShift);
                    }
                    else
                    {
                        shift.JobType = shift.Date.DayOfWeek == DayOfWeek.Saturday || shift.Date.DayOfWeek == DayOfWeek.Sunday ? Jobs.WASH_BAY_OT : Jobs.WASH_BAY;
                    }
                }
            }
            emp.Shifts.AddRange(newShifts);
        }

        private static void CalculateMgForSchoolRouteShifts(Employee emp, List<Shift> shifts)
        {
            Dictionary<int, Dictionary<RouteTimeContext, List<Shift>>> categorizedShifts = new();
            foreach (var shift in shifts)
            {
                //Log("line 129 iterationCounter == " + iterationCounter++);
                if (!categorizedShifts.ContainsKey(shift.Date.Day))
                {
                    categorizedShifts.Add(shift.Date.Day, new());
                }

                if (!categorizedShifts[shift.Date.Day].ContainsKey(shift.TimeContext()))
                {
                    categorizedShifts[shift.Date.Day].Add(shift.TimeContext(), new());
                }

                categorizedShifts[shift.Date.Day][shift.TimeContext()].Add(shift);
            }

            foreach (var pair in categorizedShifts)
            {
                //Log("line 145 iterationCounter == " + iterationCounter++);
                foreach (var pair2 in pair.Value)
                { //pair2<route time, List<shift>>

                    MgSource sourceOfMg = MgSource.NONE;
                    //float maxMinGuarantee = pair2.Value.Max(shift => shift.GetMinimumGuaranteeMax(emp, out sourceOfMg));
                    Shift? shiftToUseForMg = null;

                    float maxMinGuarantee = 0f;
                    foreach (var shift in pair2.Value)
                    {
                        float minGuarantee = pair2.Value.Max(shift => shift.GetMinimumGuaranteeMax(emp, out sourceOfMg));
                        if (minGuarantee > maxMinGuarantee)
                        {
                            shiftToUseForMg = shift;
                            maxMinGuarantee = minGuarantee;
                        }

                        if (shift.JobType == Jobs.MECHANIC && emp.IdNumber == 1248)
                        {
                            //ali omar exception
                            shift.MinimumGuaranteeHours = Math.Max(0, 1 - shift.ShiftTime);
                        }
                    }

                    if (shiftToUseForMg == null)
                    {
                        continue;
                    }
                    float mg = maxMinGuarantee;

                    foreach (var shift in pair2.Value)
                    {
                        if (shift.JobInt == shiftToUseForMg.JobInt)
                        {
                            mg -= shift.ShiftTime;
                        }
                    }

                    if (mg > 0)
                    {
                        if (sourceOfMg == MgSource.SUMMER_ROUTE)
                        {
                            shiftToUseForMg.SummerGuaranteeHours = (float)Math.Round(mg, 2);
                        }
                        else
                        {
                            shiftToUseForMg.MinimumGuaranteeHours = (float)Math.Round(mg, 2);
                        }
                        DocumentMinimumGuaranteeAmountForShift(emp.IdNumber, sourceOfMg, (float)Math.Round(mg, 2));
                    }
                }
            }
        }

        private static void CalculateMgForNonSchoolRouteShifts(Employee emp, List<Shift> shifts)
        {
            Dictionary<int, List<Shift>> categorizedShifts = new();
            foreach (var shift in shifts)
            {
                if (!categorizedShifts.ContainsKey(shift.Date.Day))
                {
                    categorizedShifts.Add(shift.Date.Day, new());
                }

                categorizedShifts[shift.Date.Day].Add(shift);
            }
            foreach (var pair in categorizedShifts)
            {
                float maxMinGuarantee = 0f;
                if (pair.Value.Count > 0)
                {
                    if (pair.Value[0].Date.DayOfWeek == DayOfWeek.Saturday || pair.Value[0].Date.DayOfWeek == DayOfWeek.Sunday)
                    {
                        MgSource sourceOfMg = MgSource.NONE;
                        maxMinGuarantee = pair.Value.Max(shift => shift.GetMinimumGuaranteeMax(emp, out sourceOfMg));
                        foreach (var shift in pair.Value)
                        {
                            if (maxMinGuarantee > shift.ShiftTime)
                            {
                                float mg = maxMinGuarantee;
                                foreach (var shift2 in pair.Value)
                                {
                                    mg -= shift2.ShiftTime;
                                }

                                if (mg > 0)
                                {
                                    shift.MinimumGuaranteeHours = (float)Math.Round(mg, 2);
                                    DocumentMinimumGuaranteeAmountForShift(emp.IdNumber, sourceOfMg, (float)Math.Round(mg, 2));
                                }
                            }
                            break;
                        }
                    }
                    else
                    {
                        foreach (var shift in pair.Value)
                        {
                            float mg = (float)Math.Round(Math.Max(0f, shift.GetMinimumGuaranteeMax(emp, out MgSource sourceOfMg) - shift.ShiftTime), 2);
                            if (mg > 0)
                            {
                                shift.MinimumGuaranteeHours = mg;
                                DocumentMinimumGuaranteeAmountForShift(emp.IdNumber, sourceOfMg, mg);
                            }
                        }
                    }
                }
            }
        }

        private static void DocumentMinimumGuaranteeAmountForShift(int employeeId, MgSource sourceOfMg, float mg)
        {
            if (MgSourceTotals.ContainsKey(sourceOfMg))
            {
                MgSourceTotals[sourceOfMg] += mg;
            }
            else
            {
                MgSourceTotals.Add(sourceOfMg, mg);
            }
            if (sourceOfMg == MgSource.SPECIAL_EXCEPTION)
            {
                if (SpecialEmployeeHandler.SpecialMgShiftTotals.ContainsKey(employeeId))
                {
                    SpecialEmployeeHandler.SpecialMgShiftTotals[employeeId] += mg;
                }
                else
                {
                    SpecialEmployeeHandler.SpecialMgShiftTotals.Add(employeeId, mg);
                }
            }
        }

        private static void TotalUpShiftsForEmployees()
        {
            foreach (var emp in EmployeeDictionary.Values)
            {
                if (emp != null && emp.Shifts.Count > 0)
                {
                    float[,] dailyRunningTotal = new float[2, 32]; //first index:1-working hours,2-all hours second index: dayNumber
                    Shift[] shiftForDay = new Shift[32];
                    bool[] bDriverOrAideShiftWasFoundForDay = new bool[32];
                    foreach (var shift in emp.Shifts)
                    {
                        if (shift.IsValid(emp))
                        {
                            SpecialEmployeeHandler.GetInstance().CheckForTimeFrameException(emp, shift);
                            DoBusStartingBonusAndMg(shift, emp);
                            CheckForMechanicHoursForShift(shift, emp);

                            float? payRate = shift.PayRate;
                            if (shift.Type() == Type.HOURS && null == payRate)
                            {
                                payRate = emp.GetPayRateForShift(shift);
                                shift.PayRate = payRate;
                            }
                            else if (shift.ShiftTime > 0.01f && shift.JobType != Jobs.DRIVER_COACH)
                            {
                                Log("Check here 131312321", true);
                            }

                            FindOrMakeMatchingShiftTotalShift(shift, emp, out Shift shiftTotalShift);
                            shiftTotalShift.PayRate = payRate;

                            //sum all hours
                            shiftTotalShift.AddAll(shift);

                            dailyRunningTotal[0, shift.Date.Day] += shift.WorkingHours();
                            dailyRunningTotal[1, shift.Date.Day] += shift.AllHours(true);
                            if (shift.JobType == Jobs.DRIVER_SCHOOL || shift.JobType == Jobs.AIDE_SCHOOL)
                            {
                                bDriverOrAideShiftWasFoundForDay[shift.Date.Day] = true;
                                shiftForDay[shift.Date.Day] = shiftTotalShift;
                            }
                        }
                    }

                    float[,,] medhusCounter = new float[3, 6, 2]; // week, hours/dollar/bonus/per diem/ot hours/ot dollars, company number
                    float jeffShawHours = 0f;
                    //total up weeks
                    float[,,] weeklyRunnningTotal = new float[2, 2, 3]; //first index:company; second index: 0-working hours,1-all hours ;third index: weekNumber
                    for (int company = 0; company < 2; ++company)
                    {
                        for (int shiftType = 0; shiftType < 3; ++shiftType)
                        {
                            if (null != emp.ShiftTotals[company, shiftType])
                            {
                                foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                                {
                                    foreach (var shifts in pair.Values)
                                    {
                                        foreach (var shift in shifts)
                                        {
                                            if (shift.IsValid(emp))
                                            {
                                                weeklyRunnningTotal[company, 0, shift.WeekNumber] += shift.WorkingHours();
                                                weeklyRunnningTotal[company, 1, shift.WeekNumber] += shift.AllHours(true);


                                                //bob medhus
                                                if (DoMedhusDeferredPayment && emp.IdNumber == 1657)
                                                {
                                                    medhusCounter[shift.WeekNumber, 0, company] += shift.WorkingHours();
                                                    float dollarAmount = shift.DollarAmount;
                                                    if (dollarAmount < 0.0001f)
                                                    {
                                                        float payRate = emp.GetPayRateForShift(shift);
                                                        if (payRate < 0.1f)
                                                        {
                                                            Log("Problem getting payrate for totaling up bob medhus's hours", true);
                                                        }
                                                        else
                                                        {
                                                            dollarAmount = shift.WorkingHours() * emp.PayRates[shift.JobType];
                                                        }
                                                    }
                                                    medhusCounter[shift.WeekNumber, 1, company] += dollarAmount;
                                                    medhusCounter[shift.WeekNumber, 2, company] += shift.BonusDollars;
                                                    medhusCounter[shift.WeekNumber, 3, company] += shift.PerDiem;
                                                }

                                                //jeff shaw
                                                if (DoJeffShawVacation && emp.IdNumber == 876)
                                                {
                                                    jeffShawHours += shift.AllHours(true);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //daily min
                    foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.DailyMgExceptions)
                    {
                        if (entry.IdNumber == emp.IdNumber)
                        {
                            List<float> dailyMgList = new();
                            for (int dayNumber = 0; dayNumber < 32; ++dayNumber)
                            {
                                if (bDriverOrAideShiftWasFoundForDay[dayNumber])
                                {
                                    if (entry.Hours > dailyRunningTotal[1, dayNumber])
                                    {
                                        float dailyMg = entry.Hours - dailyRunningTotal[1, dayNumber];
                                        shiftForDay[dayNumber].MinimumGuaranteeHours += (float)Math.Round(dailyMg, 2);
                                        dailyMgList.Add((float)Math.Round(dailyMg, 2));
                                    }
                                }
                            }
                            if (dailyMgList.Count > 0)
                            {
                                DelayedLog("Giving a total of " + dailyMgList.Sum() + " daily MG hours to " + emp.Name + " for a total of " + dailyMgList.Count + " days.");
                                SpecialEmployeeHandler.SpecialMgNonShiftTotals[emp.IdNumber] = SpecialEmployeeHandler.SpecialMgNonShiftTotals.GetValueOrDefault(emp.IdNumber, 0f) + dailyMgList.Sum();
                            }
                        }
                    }

                    for (int weekNumber = 1; weekNumber < 3; ++weekNumber)
                    {
                        //weekly min 
                        foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.WeeklyMgExceptions)
                        {
                            if (entry.IdNumber == emp.IdNumber)
                            {
                                for (int company = 0; company < 2; ++company)
                                {
                                    if (entry.Hours > weeklyRunnningTotal[company, 1, weekNumber])
                                    {
                                        float weeklyMg = entry.Hours - weeklyRunnningTotal[company, 1, weekNumber];
                                        var shift = emp.FindShiftForWeek(weekNumber, emp.PrimaryJobType(), (Company)company, false);
                                        if (shift != null)
                                        {
                                            if (shift.JobType != Jobs.HOLIDAY && shift.JobType != Jobs.VACATION)
                                            {
                                                shift.MinimumGuaranteeHours += (float)Math.Round(weeklyMg, 2);
                                                DelayedLog("Giving " + weeklyMg + " weekly MG hours to " + emp.Name);
                                                SpecialEmployeeHandler.SpecialMgNonShiftTotals[emp.IdNumber] = SpecialEmployeeHandler.SpecialMgNonShiftTotals.GetValueOrDefault(emp.IdNumber, 0f) + weeklyMg;
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }

                        //ot
                        for (int company = 0; company < 2; ++company)
                        {
                            if (weeklyRunnningTotal[company, 0, weekNumber] > 40f)
                            {
                                emp.OverTimeHours[company, weekNumber] = weeklyRunnningTotal[company, 0, weekNumber] - 40f;
                                //bob medhus
                                if (DoMedhusDeferredPayment && emp.IdNumber == 1657 && company == (int)Company.VALLEY_BUS_LLC)
                                {
                                    medhusCounter[weekNumber, 4, 0] = emp.OverTimeHours[company, weekNumber];
                                    medhusCounter[weekNumber, 5, 0] = (medhusCounter[weekNumber, 1, 0] / medhusCounter[weekNumber, 0, 0]) * medhusCounter[weekNumber, 4, 0] * 0.5f;
                                }
                            }
                        }
                    }

                    //bob medhus
                    if (DoMedhusDeferredPayment && emp.IdNumber == 1657)
                    {
                        string message = "\nFor Bob Medhus, payroll run on " + DateTime.Now.ToShortDateString() + ":";
                        for (int i = 0; i < 2; i++)
                        {
                            message += "\nFor " + (i == 0 ? "Valley Bus, LLC:" : "Valley Bus Coaches:");
                            message += "\nHours week 1:\n" + Math.Round(medhusCounter[1, 0, i], 2).ToString();
                            message += "\nHours week 2:\n" + Math.Round(medhusCounter[2, 0, i], 2).ToString();
                            message += "\nDollars week 1:\n" + Math.Round(medhusCounter[1, 1, i], 2).ToString();
                            message += "\nDollars week 2:\n" + Math.Round(medhusCounter[2, 1, i], 2).ToString();
                            message += "\nBonus Dollars week 1:\n" + Math.Round(medhusCounter[1, 2, i], 2).ToString();
                            message += "\nBonus Dollars week 2:\n" + Math.Round(medhusCounter[2, 2, i], 2).ToString();
                            message += "\nPer Diem Dollars week 1:\n" + Math.Round(medhusCounter[1, 3, i], 2).ToString();
                            message += "\nPer Diem Dollars week 2:\n" + Math.Round(medhusCounter[2, 3, i], 2).ToString();
                            message += "\nOvertime Hours week 1:\n" + Math.Round(medhusCounter[1, 4, i], 2).ToString();
                            message += "\nOvertime Hours week 2:\n" + Math.Round(medhusCounter[2, 4, i], 2).ToString();
                            message += "\nOvertime Dollars week 1:\n" + Math.Round(medhusCounter[1, 5, i], 2).ToString();
                            message += "\nOvertime Dollars week 2:\n" + Math.Round(medhusCounter[2, 5, i], 2).ToString();
                        }
                        Log(message);
                    }

                    //jeff shaw
                    if (DoJeffShawVacation && emp.IdNumber == 876 && jeffShawHours < 80)
                    {
                        if (emp.VacationHours > 70)
                        {
                            float hours = Math.Min(80 - jeffShawHours, emp.VacationHours - 70);
                            Shift newShift = new()
                            {
                                JobType = Jobs.VACATION,
                                ShiftTime = hours,
                                CompanyName = Company.VALLEY_BUS_LLC,
                                WeekNumber = 1
                            };
                            FindOrMakeMatchingShiftTotalShift(newShift, emp, out Shift shiftTotalShift);
                            shiftTotalShift.ShiftTime = hours;
                            shiftTotalShift.PayRate = emp.GetPayRateForShift(newShift);
                        }
                    }
                    if (IsLastPayPeriodOfTheYear(ExcelWorker.FirstDayWeek2))
                    {
                        float vacationHours = emp.VacationHours;
                        float hoursAlreadyRequestedForThisPayPeriod = emp.VacationHoursUsedThisPayPeriod();
                        vacationHours -= hoursAlreadyRequestedForThisPayPeriod;
                        if (vacationHours > 76)
                        {
                            float extraVacationHours = (float)Math.Round(vacationHours - 75, 2);

                            if (PrintForm.InputBool("Should we use up " + extraVacationHours + " vacation hours for " + emp.Name + " (in addition to the " + hoursAlreadyRequestedForThisPayPeriod + " hours requested)?"))
                            {
                                Log("Using " + extraVacationHours + " vacation hours for " + emp.Name + " (in addition to the " + hoursAlreadyRequestedForThisPayPeriod + " hours requested).");
                                Shift newShift = new()
                                {
                                    JobType = Jobs.VACATION,
                                    ShiftTime = extraVacationHours,
                                    CompanyName = Company.VALLEY_BUS_LLC,
                                    WeekNumber = 1
                                };
                                FindOrMakeMatchingShiftTotalShift(newShift, emp, out Shift shiftTotalShift);
                                shiftTotalShift.ShiftTime = extraVacationHours;
                                shiftTotalShift.PayRate = emp.GetPayRateForShift(newShift);
                            }
                        }
                    }
                }
            }

            LogPayRateAverages();
        }

        static void LogPayRateAverages()
        {
            bool bShouldTreatWestFargoAsFargo = true;

            for (int location = (int)Location.FARGO; location <= (int)Location.GRAND_FORKS; ++location)
            {
                var payRates = new Dictionary<Jobs, List<float>>();
                var hours = new Dictionary<Jobs, List<float>>();
                var shifts = new Dictionary<Jobs, List<Shift>>();
                var employeesPerJob = new Dictionary<Jobs, HashSet<Employee>>();
                var payRatesByEmployee = new Dictionary<Jobs, List<float>>();
                var averagePayRateByEmployee = new Dictionary<Jobs, float>();

                if (bShouldTreatWestFargoAsFargo && (Location)location == Location.WEST_FARGO)
                {
                    continue;
                }

                foreach (var emp in EmployeeDictionary.Values)
                {
                    if (emp != null && emp.Shifts.Count > 0)
                    {
                        foreach (var shift in emp.Shifts)
                        {
                            if (shift.PayRate != null)
                            {
                                if ((int)shift.ShiftLocation != location)
                                {
                                    if ((Location)location != Location.FARGO || shift.ShiftLocation != Location.WEST_FARGO)
                                    {
                                        continue;
                                    }
                                    if (!bShouldTreatWestFargoAsFargo)
                                    {
                                        continue;
                                    }
                                }
                                // Initialize collections if needed
                                if (!payRates.ContainsKey(shift.JobType))
                                {
                                    payRates[shift.JobType] = new List<float>();
                                    hours[shift.JobType] = new List<float>();
                                    shifts[shift.JobType] = new();
                                    payRatesByEmployee[shift.JobType] = new List<float>();
                                    employeesPerJob[shift.JobType] = new HashSet<Employee>();
                                }

                                // Add pay rate and hours
                                payRates[shift.JobType].Add((float)shift.PayRate);
                                hours[shift.JobType].Add((float)shift.AllHours(false));
                                shifts[shift.JobType].Add(shift);

                                // Track employee-specific pay rates
                                if (!employeesPerJob[shift.JobType].Contains(emp))
                                {
                                    employeesPerJob[shift.JobType].Add(emp);
                                    payRatesByEmployee[shift.JobType].Add((float)shift.PayRate);
                                }
                            }
                        }
                    }
                }

                // Calculate and display average pay rates
                foreach (var job in payRates.Keys)
                {
                    float averagePayRate = CalculateAveragePayRate(payRates[job], hours[job], job, (Location)location, shifts[job]);
                }

                // Calculate average pay rate by employee
                foreach (var jobType in payRatesByEmployee.Keys)
                {
                    var totalPayRate = payRatesByEmployee[jobType].Sum(); // Sum of unique pay rates
                    var employeeCount = employeesPerJob[jobType].Count;   // Count of unique employees
                    if (employeeCount > 0)
                    {
                        float averagePayRate = totalPayRate / employeeCount;
                        averagePayRateByEmployee[jobType] = averagePayRate;
                        if (jobType == Jobs.NON_CDL_DRIVER || jobType == Jobs.DRIVER_SCHOOL || jobType == Jobs.AIDE_SCHOOL)
                        {
                            Log($"{jobType}: " + $"{(Location)location}: Average Pay Rate by employee = {averagePayRate:F2} and employee count == " + employeeCount.ToString());
                        }
                    }
                }
            }
        }

        static float CalculateAveragePayRate(List<float> rates, List<float> hours, Jobs job, Location location, List<Shift> shifts)
        {
            // Ensure data consistency
            if (rates.Count != hours.Count || rates.Count == 0)
                throw new InvalidOperationException("Rates and hours must have the same non-zero count.");

            // Calculate weighted average
            float totalPay = 0;
            float totalHours = hours.Sum();
            float averagedPayRate = 0f;
            for (int i = 0; i < rates.Count; i++)
            {
                totalPay += rates[i] * hours[i];
                averagedPayRate += (hours[i] / totalHours) * rates[i];

            }

            var totalCumulativePayRate = rates.Sum();
            if (job == Jobs.NON_CDL_DRIVER || job == Jobs.DRIVER_SCHOOL || job == Jobs.AIDE_SCHOOL)
            {
                Log("(" + job.ToString() + ") averagedPayRate == " + $"{averagedPayRate:F2} and totalHours == " + totalHours + "(" + hours.Sum() + ") and totalPay == " + totalPay);
            }
            return totalPay / totalHours;
        }

        private static void FindOrMakeMatchingShiftTotalShift(Shift shiftToMatch, Employee emp, out Shift shiftTotalShift)
        {
            if (shiftToMatch.WeekNumber == 0)
            {
                shiftToMatch.WeekNumber = 1;
            }
            if (null == emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()])
            {
                emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()] = new();
            }

            if (!emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()].ContainsKey(shiftToMatch.GetLaborCode(false)))
            {
                emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()].Add(shiftToMatch.GetLaborCode(false), new());
            }

            if (!emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()][shiftToMatch.GetLaborCode(false)].ContainsKey(shiftToMatch.WeekNumber))
            {
                emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()][shiftToMatch.GetLaborCode(false)].Add(shiftToMatch.WeekNumber, new());
            }

            foreach (Shift possibleLikeShift in emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()][shiftToMatch.GetLaborCode(false)][shiftToMatch.WeekNumber])
            {
                if (null != shiftToMatch.PayRate)
                {
                    if (null != possibleLikeShift.PayRate && Math.Abs((float)(possibleLikeShift.PayRate - shiftToMatch.PayRate)) < 0.01f)
                    {
                        shiftTotalShift = possibleLikeShift;
                        return;
                    }
                }
                else if (null == possibleLikeShift.PayRate)
                {
                    shiftTotalShift = possibleLikeShift;
                    return;
                }
            }

            shiftTotalShift = new Shift(Company.VALLEY_BUS_LLC, shiftToMatch.JobType)
            {
                IsATotalsShift = true
            };
            emp.ShiftTotals[(int)shiftToMatch.CompanyName, (int)shiftToMatch.Type()][shiftToMatch.GetLaborCode(false)][shiftToMatch.WeekNumber].Add(shiftTotalShift);
        }

        private static void DoBusStartingBonusAndMg(Shift shift, Employee emp)
        {
            if (BusStartingDays.Contains(shift.Date.Day))
            {
                if (shift.ClockIn.CompareTo(new TimeSpan(6, 10, 0)) < 0 || StringSearch(shift.Notes, "starting"))
                {
                    bool foundBusStartingException = false;
                    foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.BusStartingBonusDollars)
                    {
                        if (entry.IdNumber == emp.IdNumber && (Jobs)entry.JobType == shift.JobType)
                        {
                            foundBusStartingException = true;
                            shift.BonusDollars += entry.Dollars;
                            if (entry.ReceivesBusStartingMinimumGuarantee && shift.ShiftTime < 2)
                            {
                                shift.MinimumGuaranteeHours += 2f - shift.ShiftTime;
                            }
                        }
                    }
                    if (!foundBusStartingException)
                    {
                        Log("Warning: No bus starting exception found for employee " + emp.Name + " for shift on " + shift.Date.ToShortDateString() + " with job type " + shift.JobType.ToString() + ".\n(Shift notes: " + shift.Notes + ")");
                    }
                }
            }
        }

        private static void CheckForMechanicHoursForShift(Shift shift, Employee emp)
        {
            if (emp.IsAMechanicApprentice)
            {
                if (shift.JobType == Jobs.DRIVER_SCHOOL || shift.JobType == Jobs.MECHANIC)
                {
                    if (!ApprenticeMechanicHours.ContainsKey(emp.IdNumber))
                    {
                        ApprenticeMechanicHours.Add(emp.IdNumber, new());
                    }
                    ApprenticeMechanicHours[emp.IdNumber][shift.JobType] = ApprenticeMechanicHours[emp.IdNumber].GetValueOrDefault(shift.JobType, 0f) + shift.ShiftTime;
                }
            }
        }

        public static void DoEmployeeRaises()
        {
            foreach (var employee in EmployeeDictionary.Values)
            {
                if (!employee.HadHoursInTimesheets)
                {
                    continue;
                }
                for (int jobOrdinal = 0;  jobOrdinal <= (int)Jobs.AIDE_SCHOOL; ++jobOrdinal)
                {
                    Jobs jobType = (Jobs)jobOrdinal;
                    if (employee.PayRates.ContainsKey(jobType) && employee.PayRates[jobType] > 0)
                    {
                        float baseRate = GetBasePayRateForEmployee(jobType, employee);
                        float newRate = TimeInServiceAdjustment(baseRate, employee, jobType, false);
                        if (employee.PayRates[jobType] < newRate)
                        {
                            GiveRaiseToEmployee(employee, jobType, newRate);
                        }
                    }
                }
            }
        }

        public static float TimeInServiceAdjustment(float baseRate, Employee employee, Jobs jobType, bool bShouldIncludeNonCdl)
        {
            float newRate = baseRate;
            if (newRate > 0)
            {
                if (jobType == Jobs.DRIVER_SCHOOL || jobType == Jobs.AIDE_SCHOOL || (bShouldIncludeNonCdl && jobType == Jobs.NON_CDL_DRIVER))
                {
                    for (int years = 6; years > 0; --years)
                    {
                        if (employee.YearsOfService >= years)
                        {
                            newRate += 0.25f * years;
                            break;
                        }
                    }
                }
                if (employee.YearsOfService > 9 && (jobType == Jobs.DRIVER_SCHOOL || jobType == Jobs.NON_CDL_DRIVER || jobType == Jobs.DRIVER_CHARTER || jobType == Jobs.AIDE_SCHOOL || jobType == Jobs.AIDE_CHARTER))
                {
                    newRate += TEN_YEAR_RATE_BUMP;
                }
            }
            return newRate;
        }

        public static void GiveRaiseToEmployee(Employee employee, Jobs job, float rate)
        {
            if (Jobs.DRIVER_CHARTER_PUBLIC == job)
            {
                return;
            }
            employee.NeedsUpdateInPayroll = true;
            employee.PayRates[job] = rate;
            if (!ExcelWorker.ImportEmployees.ContainsKey(employee.IdNumber))
            {
                ExcelWorker.ImportEmployees.Add(employee.IdNumber, new()
                {
                    ImportFields = new()
                    {
                        { "EmployeeNumber", employee.IdNumber.ToString() },
                        { "EmploymentCategory", employee.EmploymentCategory },
                        { "SSN", employee.SocialSecurityNumber }
                    }
                });
            }
            switch (job)
            {
                case Jobs.DRIVER_SCHOOL:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_DrvrDlySchool"] = rate.ToString();
                    break;
                case Jobs.AIDE_SCHOOL:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_AidDlySchool"] = rate.ToString();
                    break;
                case Jobs.DRIVER_CHARTER:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_DrvrSchoolChrtr"] = rate.ToString();
                    break;
                case Jobs.AIDE_CHARTER:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_AidDlyChrter"] = rate.ToString();
                    break;
                case Jobs.TRAINING:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_Training"] = rate.ToString();
                    break;
                case Jobs.ADMIN:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_Admin"] = rate.ToString();
                    break;
                case Jobs.WASH_BAY:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_Wash Bay"] = rate.ToString();
                    break;
                case Jobs.BODY_SHOP:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_Body Shop"] = rate.ToString();
                    break;
                case Jobs.MECHANIC:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_Mechanic"] = rate.ToString();
                    break;
                case Jobs.CLEANING:
                    ExcelWorker.ImportEmployees[employee.IdNumber].ImportFields["Rate_Cleaning"] = rate.ToString();
                    break;
                default:
                    Log("Warning: Trying to import raise for " + job.ToString() + " but can't determine import header.");
                    break;
            }
        }

        public static float GetBasePayRateForEmployee(Jobs jobType, Employee employee, bool bIsForGrandForks = false)
        {
            float modifier = 0f;
            foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.StartingRateExceptions)
            {
                if (entry.IdNumber == employee.IdNumber && (Jobs)entry.JobType == jobType)
                {
                    modifier = entry.Rate;
                    break;
                }
            }

            if (jobType == Jobs.DRIVER_SCHOOL && employee.IsAGrandForksEmployee && employee.HireDate.CompareTo(new DateTime(2024, 05, 01)) < 0)
            {
                modifier = Math.Max(modifier, 1f);
            }

            bIsForGrandForks |= employee.IsAGrandForksEmployee;

            if (bIsForGrandForks && GrandForksDefaultRates.ContainsKey(jobType))
            {
                return GrandForksDefaultRates[jobType] + modifier;
            }
            else if (!bIsForGrandForks && FargoDefaultRates.ContainsKey(jobType))
            {
                return FargoDefaultRates[jobType] + modifier;
            }
            return 0;
        }

        public static void CheckForVacationCutOff(DateTime firstDayWeekTwo)
        {
           if (IsLastPayPeriodOfTheYear(firstDayWeekTwo))
            {
                Log("This is the last pay period for the year. Please check accrual cut-offs", true);
            }
        }

        public static bool IsLastPayPeriodOfTheYear(DateTime firstDayWeekTwo)
        {
            DateTime payDate = firstDayWeekTwo.AddDays(12);

            if (payDate.Month == 1 && payDate.Day == 1)
            {
                payDate = payDate.AddDays(-1);
            }

            DateTime nextPayDate = payDate.AddDays(7);

            return nextPayDate.Year > payDate.Year;
        }

        public static string DesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
        }

        public static string MakeLog()
        {
            string[] paths = new string[3];
            paths[0] = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.Parent.Parent.FullName;
            paths[0] += "\\Logs\\Log" + DateTime.Today.ToShortDateString().Replace("/", "-") + "_forPayDate_" + ExcelWorker.FirstDayWeek2.AddDays(12).ToShortDateString().Replace("/", "-") + ".txt";
            paths[1] = "C:\\Users\\User\\valleybusllc.com\\PayrollExceptionMonitoring - PayrollMonitoring\\" + "log_forPayDate_" + ExcelWorker.FirstDayWeek2.AddDays(12).ToShortDateString().Replace("/", "-") + ".txt";

            paths[2] = DesktopPath() + "PayrollLog.txt";

            for (int i = 0; i < paths.Length; ++i)
            {
                if (File.Exists(paths[i]))
                {
                    File.Delete(paths[i]);
                }

                // Create a new file
                using (FileStream fs = File.Create(paths[i]))
                {
                    // Add some text to file
                    Byte[] log = new UTF8Encoding(true).GetBytes(LogString);
                    fs.Write(log, 0, log.Length);
                }
            }
            return paths[1];
        }

        public static void Exit()
        {
            Environment.Exit(0);
        }

        private static void FinalLogging()
        {
            foreach (var pair in DelayedLogMessages)
            {
                Log(pair.Key, pair.Value);
            }

            Log("Minimum Guarantee totals:");
            foreach (var entry in MgSourceTotals)
            {
                Log(entry.Key.ToString() + ": " + Math.Round(entry.Value, 2));
            }

            //apprentice mechanics
            List<int> apprenticeMechanicOrder = new()
            {
                1947,1963,1419,1946,1876,2100,1976,2282
            };
            foreach (var empEntry in EmployeeDictionary)
            {
                if (empEntry.Value.IsAMechanicApprentice && !apprenticeMechanicOrder.Contains(empEntry.Key))
                {
                    apprenticeMechanicOrder.Add(empEntry.Key);
                }
            }
            Log("\nApprentice Mechanic Order:");
            foreach (var mc in apprenticeMechanicOrder)
            {
                Log(mc.ToString());
            }
            Log("Mechanic hours:");
            foreach (var mc in apprenticeMechanicOrder)
            {
                if (ApprenticeMechanicHours.ContainsKey(mc))
                {
                    Log(Math.Round(ApprenticeMechanicHours[mc].GetValueOrDefault(Jobs.MECHANIC, 0f), 2).ToString());
                }
                else
                {
                    Log("0");
                }
            }
            Log("Driver hours:");
            foreach (var mc in apprenticeMechanicOrder)
            {
                if (ApprenticeMechanicHours.ContainsKey(mc))
                {
                    Log(Math.Round(ApprenticeMechanicHours[mc].GetValueOrDefault(Jobs.DRIVER_SCHOOL, 0f), 2).ToString());
                }
                else
                {
                    Log("0");
                }
            }

            string nonCdlDrivers = "\nNon CDL Drivers: \n\nFargo:\n";
            foreach (var employee in NonCdlDrivers)
            {
                if (!employee.IsAGrandForksEmployee)
                {
                    nonCdlDrivers += employee.Name + "\n";
                }
            }
            nonCdlDrivers += "\nGrand Forks:\n";
            foreach (var employee in NonCdlDrivers)
            {
                if (employee.IsAGrandForksEmployee)
                {
                    nonCdlDrivers += employee.Name + "\n";
                }
            }
            nonCdlDrivers += "\n";
            Log(nonCdlDrivers);

            SpecialEmployeeHandler.GetInstance().AddExceptionNotificationsToLog();

            DolStatisticsTracker dolStatisticsTracker = new();
            foreach (var emp in EmployeeDictionary.Values)
            {
                dolStatisticsTracker.RegisterEmployeeAfterShiftTotals(emp);
            }
            dolStatisticsTracker.AddDolStatisticsToLog();


            string logPath = MakeLog();
            var process = new Process();
            process.StartInfo = new ProcessStartInfo()
            {
                UseShellExecute = true,
                FileName = logPath
            };

            process.Start();
        }
    }

    public enum Jobs
    {
        DRIVER_SCHOOL = 1, DRIVER_CHARTER = 2, DRIVER_CHARTER_PUBLIC = 3, MECHANIC = 7, WASH_BAY = 9, WASH_BAY_OT = 10, TRAINING = 11, BODY_SHOP = 12, ADMIN = 13, CLEANING = 14, HOLIDAY = 15, 
        VACATION = 16, COACH_PUBLIC_DRIVING = 19/*out of town yellows*/, AIDE_CHARTER = 24, 
        AIDE_SCHOOL = 25, DRIVER_COACH = 26, OUT_OF_TOWN_CHARTER = 27, NON_CDL_DRIVER = 28, 
        
        SALARY = 99
    }
    //  taskkill /f /im excel.exe
}