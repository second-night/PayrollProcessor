using Xunit;

namespace PayrollProcessor.Tests
{
    public class RehireSelectionTests
    {
        private static readonly DateTime PayPeriodStart = new(2026, 10, 4);
        private static readonly DateTime TerminatedLastYear = new(2025, 6, 1);

        [Fact]
        public void TerminatedCompanyWithHours_IsRehired()
        {
            Employee employee = TerminatedAt(Company.VALLEY_BUS_LLC, TerminatedLastYear);
            AddHours(employee, Company.VALLEY_BUS_LLC);

            List<Company> companies = EmployeePayrollHistory.SelectCompaniesToRehire(employee, PayPeriodStart);

            Assert.Equal(new[] { Company.VALLEY_BUS_LLC }, companies);
        }

        [Fact]
        public void ActiveCompanyWithHours_IsNotRehired()
        {
            Employee employee = TerminatedAt(Company.VALLEY_BUS_LLC, TerminatedLastYear);
            employee.ActiveCompanies.Add(Company.VALLEY_BUS_COACHES);
            AddHours(employee, Company.VALLEY_BUS_COACHES);

            List<Company> companies = EmployeePayrollHistory.SelectCompaniesToRehire(employee, PayPeriodStart);

            Assert.Empty(companies);
        }

        [Fact]
        public void TerminatedNonPrimaryCompany_IsRehiredWithoutTheActivePrimary()
        {
            Employee employee = new(100, "Test Person")
            {
                PrimaryCompany = Company.VALLEY_BUS_LLC
            };
            employee.ActiveCompanies.Add(Company.VALLEY_BUS_LLC);
            employee.TerminationDateByCompany[Company.VALLEY_BUS_COACHES] = TerminatedLastYear;
            AddHours(employee, Company.VALLEY_BUS_COACHES);

            List<Company> companies = EmployeePayrollHistory.SelectCompaniesToRehire(employee, PayPeriodStart);

            Assert.Equal(new[] { Company.VALLEY_BUS_COACHES }, companies);
        }

        [Fact]
        public void RehiringNonPrimary_AlsoRehiresTerminatedPrimary()
        {
            Employee employee = TerminatedAt(Company.VALLEY_BUS_LLC, TerminatedLastYear);
            employee.TerminationDateByCompany[Company.VALLEY_BUS_COACHES] = TerminatedLastYear;
            AddHours(employee, Company.VALLEY_BUS_COACHES);

            List<Company> companies = EmployeePayrollHistory.SelectCompaniesToRehire(employee, PayPeriodStart);

            Assert.Equal(new[] { Company.VALLEY_BUS_LLC, Company.VALLEY_BUS_COACHES }, companies);
        }

        [Theory]
        [InlineData(2026, 10, 4)]
        [InlineData(2026, 10, 8)]
        public void TerminatedOnOrAfterPayPeriodStart_IsNotRehired(int year, int month, int day)
        {
            Employee employee = TerminatedAt(Company.VALLEY_BUS_LLC, new DateTime(year, month, day));
            AddHours(employee, Company.VALLEY_BUS_LLC);

            List<Company> companies = EmployeePayrollHistory.SelectCompaniesToRehire(employee, PayPeriodStart);

            Assert.Empty(companies);
        }

        [Fact]
        public void ManualEntryOnTerminatedCompany_IsRehired()
        {
            Employee employee = TerminatedAt(Company.VALLEY_BUS_COACHES, TerminatedLastYear);
            employee.ManualEntries.Add(new ManualEntry
            {
                Company = Company.VALLEY_BUS_COACHES,
                RegularHours = 4f
            });

            List<Company> companies = EmployeePayrollHistory.SelectCompaniesToRehire(employee, PayPeriodStart);

            Assert.Equal(new[] { Company.VALLEY_BUS_COACHES }, companies);
        }

        private static Employee TerminatedAt(Company company, DateTime terminationDate)
        {
            Employee employee = new(100, "Test Person")
            {
                PrimaryCompany = company,
                IsTerminated = true,
                TerminationDate = terminationDate
            };
            employee.TerminationDateByCompany[company] = terminationDate;
            return employee;
        }

        private static void AddHours(Employee employee, Company company)
        {
            employee.Shifts.Add(new Shift(company) { ShiftTime = 8f });
        }
    }
}
