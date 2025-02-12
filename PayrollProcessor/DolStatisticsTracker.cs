using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PayrollProcessor
{
    internal class DolStatisticsTracker
    {
        public int EmployeeCount;
        public int FemaleEmployeeCount;
        public float TotalCompensation;
        public float TotalHoursCompensated;
        HashSet<int> EmployeeIds;

        private void RegisterSalariedEmployee(Employee emp)
        {
            float perPaySalary = emp.AnnualSalaryAmount / 26;
            float hours = emp.AnnualSalaryAmount < 30000f ? 40f : 80f; //close enough
            TotalCompensation += perPaySalary;
            TotalHoursCompensated += hours;
        }

        public void RegisterEmployeeAfterShiftTotals(Employee emp)
        {
            if (EmployeeIds.Contains(emp.IdNumber))
            {
                return;
            }
            EmployeeIds.Add(emp.IdNumber);

            if (emp.IsTerminated && emp.Shifts.Count == 0)
            {
                return;
            }

            EmployeeCount++;
            if (!emp.IsMale)
            {
                FemaleEmployeeCount++;
            }

            if (emp.AnnualSalaryAmount > 0 && !emp.IsTerminated)
            {  
                RegisterSalariedEmployee(emp);
            }

            for (int company = (int)Company.VALLEY_BUS_LLC; company <= (int)Company.VALLEY_BUS_COACHES; ++company)
            {
                for (int shiftType = 0; shiftType < 3; ++shiftType)
                {
                    if (null != emp.ShiftTotals[company, shiftType])
                    {
                        foreach (var pair in emp.ShiftTotals[company, shiftType].Values)
                        {
                            foreach (var shifts in pair.Values)
                            {
                                foreach (Shift shift in shifts)
                                {
                                    TotalCompensation += shift.TotalCompensation(emp);
                                    TotalHoursCompensated += shift.AllHours(true);
                                }
                            }
                        }
                    }
                }
            }
        }

        public void AddDolStatisticsToLog()
        {
            string str = "Department of Labor Statistics:";
            str += "\nEmployee count: " + EmployeeCount.ToString();
            str += "\nFemale employee count: " + FemaleEmployeeCount.ToString();
            str += "\nTotal compensation: " + Math.Round(TotalCompensation, 2).ToString();
            str += "\nTotal hours: " + Math.Round(TotalHoursCompensated, 0);
            str += "\n(Average pay rate: " + Math.Round(TotalCompensation / TotalHoursCompensated, 2).ToString() + ")";
        }
    }
}
