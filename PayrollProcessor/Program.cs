using Microsoft.VisualBasic;
using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace PayrollProcessor
{
    //taskkill /f /im excel.exe

    //TODO: display list of people getting special exceptions, display list of non-cdl drivers.
    public static class Program
    {
        public static Dictionary<int, Employee> EmployeeDictionary = new();
        public const float OUT_OF_TOWN_CHARTER_RATE = 18f;
        public const float T_AND_J_CHARTER_RATE = 19f;
        public const float OUT_OF_TOWN_CHARTERS_MG_IN_DOLLARS = 120f;
        public const float OUT_OF_TOWN_OR_WEEKEND_MIN_GUARANTEE_DRIVER_IN_DOLLARS = 50f;
        public const float TJ_OR_WEEKEND_MIN_GUARANTEE_AIDE_IN_DOLLARS = 40f;
        public const float DRIVER_CHARTER_RATE = 17.5f;
        public const float TRAINING_RATE= 12.5f;
        public const float STARTING_WASH_BAY_RATE = 16f;
        public const float COACH_HOURLY_RATE_ESTIMATE = 19f;
        public const float GF_HOCKEY_PAY = 100f;
        public const float GF_HOCKEY_BAND_PAY = 120f;
        public static string LogString = "";
        public static HashSet<int> BusStartingDays = new();
        private static ExcelWorker ExcelWorker;
        private static Dictionary<MgSource, float> MgSourceTotals = new();
        public static Dictionary<String, bool> DelayedLogMessages = new();
        public static HashSet<Employee> NonCdlDrivers = new();
        private static Dictionary<int, Dictionary<Jobs, float>> ApprenticeMechanicHours = new();
        public static Dictionary<Jobs, float> FargoDefaultRates = new()
        {
            {Jobs.DRIVER_SCHOOL, 20f },
            {Jobs.DRIVER_CHARTER, DRIVER_CHARTER_RATE },
            {Jobs.COACH_PUBLIC_DRIVING, OUT_OF_TOWN_CHARTER_RATE },
            {Jobs.AIDE_SCHOOL, 17.5f },
            {Jobs.AIDE_CHARTER, 16f },
            {Jobs.NON_CDL_DRIVER, 17.5f },
            {Jobs.TRAINING, TRAINING_RATE }
        };
        public static Dictionary<Jobs, float> GrandForksDefaultRates = new()
        {
            {Jobs.DRIVER_SCHOOL, 23f },
            {Jobs.DRIVER_CHARTER, DRIVER_CHARTER_RATE },
            {Jobs.AIDE_SCHOOL, 18.5f },
            {Jobs.AIDE_CHARTER, 17.5f },
            {Jobs.NON_CDL_DRIVER, 19f },
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
                {
                    MgSource sourceOfMg = MgSource.NONE;
                    float maxMinGuarantee = pair2.Value.Max(shift => shift.GetMinimumGuaranteeMax(emp, out sourceOfMg));
                    //if (emp.IdNumber == 1893)
                    //{
                    //    Log("maxMinGuarantee == " + maxMinGuarantee);
                    //}
                    foreach (var shift in  pair2.Value)
                    {
                        //if (emp.IdNumber == 1893)
                        //{
                        //    Log("shift.ShiftTime == " + shift.ShiftTime);
                        //}
                        if (maxMinGuarantee > shift.ShiftTime)
                        {
                            float mg = maxMinGuarantee;
                            foreach (var shift2 in pair2.Value)
                            {
                                mg -= shift2.ShiftTime;
                            }

                            if (mg > 0)
                            {
                                if (sourceOfMg == MgSource.SUMMER_ROUTE)
                                {
                                    shift.SummerGuaranteeHours = (float)Math.Round(mg, 2);
                                }
                                else
                                {
                                    shift.MinimumGuaranteeHours = (float)Math.Round(mg, 2);
                                }
                                DocumentMinimumGuaranteeAmountForShift(emp.IdNumber, sourceOfMg, (float)Math.Round(mg, 2));
                                //if (emp.IdNumber == 1893)
                                //{
                                //    Log("mg == " + mg);
                                //}
                            }
                        }
                        break;
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
                if (emp != null)
                {
                    if (emp.Shifts.Count > 0)
                    {
                        float[,] dailyRunningTotal = new float[2, 32]; //first index:1-working hours,2-all hours second index: dayNumber
                        Shift[] shiftForDay = new Shift[32];
                        bool[] bDriverOrAideShiftWasFoundForDay = new bool[32];
                        foreach (var shift in emp.Shifts)
                        {
                            if (shift.IsValid(emp))
                            {
                                SpecialEmployeeHandler.GetInstance().CheckForTimeFrameException(emp, shift);
                                if (shift.Type() == Type.HOURS && shift.HasSpecialPayRate(emp))
                                {
                                    shift.DollarAmount = (float)Math.Round(shift.ShiftTime * shift.SpecialRate(emp), 2);
                                    if (shift.MinimumGuaranteeHours > 0)
                                    {
                                        shift.MgDollars = (float)Math.Round(shift.MinimumGuaranteeHours * shift.SpecialRate(emp), 2);
                                    }
                                }
                                if (null == emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()])
                                {
                                    emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()] = new();
                                }

                                if (!emp.ShiftTotals[(int)shift.CompanyName, (int) shift.Type()].ContainsKey(shift.GetLaborCode(false)))
                                {
                                    emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()].Add(shift.GetLaborCode(false), new());
                                }

                                if (!emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()][shift.GetLaborCode(false)].ContainsKey(shift.WeekNumber))
                                {
                                    emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()][shift.GetLaborCode(false)].Add(shift.WeekNumber, new Shift(Company.VALLEY_BUS_LLC, shift.JobType));
                                }

                                if (BusStartingDays.Contains(shift.Date.Day))
                                {
                                    if ((shift.JobType == Jobs.MECHANIC && shift.ClockIn.CompareTo(new TimeSpan(6, 10, 0)) < 0) || StringSearch(shift.Notes, "bonus"))
                                    {
                                        foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.BusStartingBonusDollars)
                                        {
                                            if (entry.IdNumber == emp.IdNumber && (Jobs)entry.JobType == shift.JobType)
                                            {
                                                shift.BonusDollars += entry.Dollars;
                                            }
                                        }
                                    }
                                }

                                //sum all hours
                                emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()][shift.GetLaborCode(false)][shift.WeekNumber].IsATotalsShift = true;
                                emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()][shift.GetLaborCode(false)][shift.WeekNumber].AddAll(shift);

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

                                dailyRunningTotal[0, shift.Date.Day] += shift.WorkingHours();
                                dailyRunningTotal[1, shift.Date.Day] += shift.AllHours(true);
                                if (shift.JobType == Jobs.DRIVER_SCHOOL || shift.JobType == Jobs.AIDE_SCHOOL)
                                {
                                    bDriverOrAideShiftWasFoundForDay[shift.Date.Day] = true;
                                    shiftForDay[shift.Date.Day] = emp.ShiftTotals[(int)shift.CompanyName, (int)shift.Type()][shift.GetLaborCode(false)][shift.WeekNumber];
                                }
                            }
                        }

                        float[,,] medhusCounter = new float[3, 6, 2]; // week, hours/dollar/bonus/per diem/ot hours/ot dollars, company number
                        //total up weeks
                        float[,] weeklyRunnningTotal = new float[2, 3]; //first index:0-working hours,1-all hours second index: weekNumber
                        for (int company = 0; company < 2; ++company)
                        {
                            for (int shiftType = 0; shiftType < 3; ++shiftType)
                            {
                                if (null != emp.ShiftTotals[company, shiftType])
                                {
                                    foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                                    {
                                        foreach (var shift in pair.Values)
                                        {
                                            if (shift.IsValid(emp))
                                            {
                                                weeklyRunnningTotal[0, shift.WeekNumber] += shift.WorkingHours();
                                                weeklyRunnningTotal[1, shift.WeekNumber] += shift.AllHours(true);
                                                //bob medhus
                                                if (emp.IdNumber == 1657)
                                                {
                                                    medhusCounter[shift.WeekNumber, 0, company] += shift.WorkingHours();
                                                    float dollarAmount = shift.DollarAmount;
                                                    if (dollarAmount < 0.0001f)
                                                    {
                                                        if (!emp.PayRates.ContainsKey(shift.JobType))
                                                        {
                                                            Log("Problem getting payrate for totaling up bob medhus's hours", true);
                                                        }
                                                        else
                                                        {
                                                            dollarAmount = shift.WorkingHours() * emp.PayRates[shift.JobType];
                                                        }
                                                    }
                                                    medhusCounter[shift.WeekNumber, 1, company] += dollarAmount;
                                                    medhusCounter[shift.WeekNumber, 2, company] += shift.BonusDollars + shift.MgDollars;
                                                    medhusCounter[shift.WeekNumber, 3, company] += shift.PerDiem;
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
                                            if (shiftForDay[dayNumber].MgDollars > 0)
                                            {
                                                if (shiftForDay[dayNumber].SpecialRate(emp) < 0.01f)
                                                {
                                                    Log("ERROR:5454352", true);
                                                }
                                                shiftForDay[dayNumber].MgDollars += (float)Math.Round(dailyMg * shiftForDay[dayNumber].SpecialRate(emp), 2);
                                            }
                                        }
                                        break;
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
                                    if (entry.Hours > weeklyRunnningTotal[1, weekNumber])
                                    {
                                        float weeklyMg = entry.Hours - weeklyRunnningTotal[1, weekNumber];
                                        emp.FindDriverOrAideShiftForWeek(weekNumber, emp.IsADriverOrAnAide()).MinimumGuaranteeHours += (float)Math.Round(weeklyMg, 2);

                                        if (emp.FindDriverOrAideShiftForWeek(weekNumber, emp.IsADriverOrAnAide()).MgDollars > 0)
                                        {
                                            if (emp.FindDriverOrAideShiftForWeek(weekNumber, emp.IsADriverOrAnAide()).SpecialRate(emp) < 0.01f)
                                            {
                                                Log("ERROR:54543656552", true);
                                            }
                                            emp.FindDriverOrAideShiftForWeek(weekNumber, emp.IsADriverOrAnAide()).MgDollars += (float)Math.Round(weeklyMg * emp.FindDriverOrAideShiftForWeek(weekNumber, emp.IsADriverOrAnAide()).SpecialRate(emp), 2);
                                        }
                                        DelayedLog("Giving " + weeklyMg + " weekly MG hours to " + emp.Name);
                                        SpecialEmployeeHandler.SpecialMgNonShiftTotals[emp.IdNumber] = SpecialEmployeeHandler.SpecialMgNonShiftTotals.GetValueOrDefault(emp.IdNumber, 0f) + weeklyMg;
                                    }
                                    break;
                                }
                            }

                            //ot
                            if (weeklyRunnningTotal[0, weekNumber] > 40f)
                            {
                                emp.OverTimeHours[weekNumber] = weeklyRunnningTotal[0, weekNumber] - 40f;
                                //bob medhus
                                if (emp.IdNumber == 1657)
                                {
                                    medhusCounter[weekNumber, 4, 0] = emp.OverTimeHours[weekNumber];
                                    medhusCounter[weekNumber, 5, 0] = (medhusCounter[weekNumber, 1, 0] / medhusCounter[weekNumber, 0, 0]) * medhusCounter[weekNumber, 4, 0] * 0.5f;
                                }
                            }
                        }

                        //bob medhus
                        if (emp.IdNumber == 1657)
                        {
                            string message = "\nFor Bob Medhus, payroll run on " + DateTime.Now.ToShortDateString() + ":";
                            for (int i = 0; i < 2; i++)
                            {
                                message += "\nFor " + (i == 0 ? "Valley Bus, LLC:" : "Valley Bus Coaches:");
                                message += "\nHours week 1: " + Math.Round(medhusCounter[1, 0, i], 2).ToString();
                                message += "\nHours week 2: " + Math.Round(medhusCounter[2, 0, i], 2).ToString();
                                message += "\nDollars week 1: " + Math.Round(medhusCounter[1, 1, i], 2).ToString();
                                message += "\nDollars week 2: " + Math.Round(medhusCounter[2, 1, i], 2).ToString();
                                message += "\nBonus Dollars week 1: " + Math.Round(medhusCounter[1, 2, i], 2).ToString();
                                message += "\nBonus Dollars week 2: " + Math.Round(medhusCounter[2, 2, i], 2).ToString();
                                message += "\nPer Diem Dollars week 1: " + Math.Round(medhusCounter[1, 3, i], 2).ToString();
                                message += "\nPer Diem Dollars week 2: " + Math.Round(medhusCounter[2, 3, i], 2).ToString();
                                message += "\nOvertime Hours week 1: " + Math.Round(medhusCounter[1, 4, i], 2).ToString();
                                message += "\nOvertime Hours week 2: " + Math.Round(medhusCounter[2, 4, i], 2).ToString();
                                message += "\nOvertime Dollars week 1: " + Math.Round(medhusCounter[1, 5, i], 2).ToString();
                                message += "\nOvertime Dollars week 2: " + Math.Round(medhusCounter[2, 5, i], 2).ToString();
                            }
                            Log(message);
                        }
                    }
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
                bool bDoingAnnualRaises = false; //TODO: GUI
                float driverRaise = 0f;
                float aideRaise = 0f;
                for (int jobOrdinal = 0;  jobOrdinal <= (int)Jobs.AIDE_SCHOOL; ++jobOrdinal)
                {
                    Jobs job = (Jobs)jobOrdinal;
                    if (employee.PayRates.ContainsKey(job) && employee.PayRates[job] > 0)
                    {
                        float rate = GetBasePayRateForEmployee(job, employee);
                        float currentRate = employee.PayRates[job];
                        if (rate > 0)
                        {
                            if (job == Jobs.DRIVER_SCHOOL || job == Jobs.AIDE_SCHOOL)
                            {
                                if (bDoingAnnualRaises)
                                {
                                    currentRate += job == Jobs.DRIVER_SCHOOL ? driverRaise : aideRaise;
                                }
                                for (int years = 6; years > 0; --years)
                                {
                                    if (employee.YearsOfService >= years)
                                    {
                                        rate += 0.25f * years;
                                        break;
                                    }
                                }
                            }
                            float newRate = Math.Max(rate, currentRate);
                            if (job == Jobs.DRIVER_SCHOOL && employee.PayRates.GetValueOrDefault(Jobs.MECHANIC, 0f) > newRate)
                            {
                                if (employee.IdNumber != 105)
                                { //Michael Mollenhoff exception
                                    newRate = employee.PayRates[Jobs.MECHANIC];
                                    Log("Giving driver rate upgrade for mechanic; " + employee.Name + ". Upgrading from " + employee.PayRates[Jobs.DRIVER_SCHOOL] + " to " + newRate, employee.EmploymentCategory != "ACAFT");
                                }
                            }
                            if (employee.PayRates[job] < newRate)
                            {
                                GiveRaiseToEmployee(employee, job, newRate);
                            }
                        }
                    }
                }
            }
        }

        public static void GiveRaiseToEmployee(Employee employee, Jobs job, float rate)
        {
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

        public static float GetBasePayRateForEmployee(Jobs jobType, Employee employee)
        {
            foreach (var entry in SpecialEmployeeHandler.GetInstance().SpecialEmployees.StartingRateExceptions)
            {
                if (entry.IdNumber == employee.IdNumber && (Jobs)entry.JobType == jobType)
                {
                    return entry.Rate;
                }
            }
            if (employee.IsGrandForksEmployee && GrandForksDefaultRates.ContainsKey(jobType))
            {
                return GrandForksDefaultRates[jobType];
            }
            else if (!employee.IsGrandForksEmployee && FargoDefaultRates.ContainsKey(jobType))
            {
                return FargoDefaultRates[jobType];
            }
            return 0;
        }

        public static string DesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
        }

        public static string MakeLog()
        {
            string[] paths = new string[2];
            paths[0] = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.Parent.Parent.FullName;
            paths[0] += "\\Logs\\Log" + DateTime.Today.ToShortDateString().Replace("/", "-") + "_forPayday_" + ExcelWorker.FirstDayWeek2.AddDays(12).ToShortDateString().Replace("/", "-") + ".txt";

            paths[1] = DesktopPath() + "PayrollLog.txt";

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
                1947,1963,1419,1946,1876,2100
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
                if (!employee.IsGrandForksEmployee)
                {
                    nonCdlDrivers += employee.Name + "\n";
                }
            }
            nonCdlDrivers += "\nGrand Forks:\n";
            foreach (var employee in NonCdlDrivers)
            {
                if (employee.IsGrandForksEmployee)
                {
                    nonCdlDrivers += employee.Name + "\n";
                }
            }
            nonCdlDrivers += "\n";
            Log(nonCdlDrivers, true);

            SpecialEmployeeHandler.GetInstance().AddExceptionNotificationsToLog();


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
        DRIVER_SCHOOL = 1, DRIVER_CHARTER = 2, MECHANIC = 7, WASH_BAY = 9, WASH_BAY_OT = 10, TRAINING = 11, BODY_SHOP = 12, ADMIN = 13, CLEANING = 14, HOLIDAY = 15, VACATION = 16, COACH_PUBLIC_DRIVING = 19/*out of town yellows*/, AIDE_CHARTER = 24, AIDE_SCHOOL = 25, DRIVER_COACH, OUT_OF_TOWN_CHARTER, NON_CDL_DRIVER
    }
}