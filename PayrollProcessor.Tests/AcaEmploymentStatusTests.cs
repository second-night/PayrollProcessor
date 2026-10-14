using Xunit;

namespace PayrollProcessor.Tests
{
    public class AcaEmploymentStatusTests
    {
        private static readonly DateTime AsOf = new(2026, 11, 6);
        private static readonly DateTime[] PayDates =
        {
            new(2026, 8, 14),
            new(2026, 8, 28),
            new(2026, 9, 11),
            new(2026, 9, 25),
            new(2026, 10, 9),
            new(2026, 10, 23),
            new(2026, 11, 6)
        };

        [Theory]
        [InlineData("Full-Time", "ACAFT")]
        [InlineData("Full Time", "ACAFT")]
        [InlineData("FT", "ACAFT")]
        [InlineData("ACAFT", "ACAFT")]
        [InlineData("Part-Time", "PT")]
        [InlineData("Part Time", "PT")]
        [InlineData("pt", "PT")]
        public void MapEmploymentCategory_RecognizesHyphenatedNames(string source, string expected)
        {
            Assert.Equal(expected, WfnEmployeesReader.MapEmploymentCategory(source));
        }

        [Fact]
        public void SteadyThirtyHourWeek_QualifiesAfterTheWindowEnds()
        {
            PayrollHistoryEmployee employee = EmployeeWithPaychecks(60f, new DateTime(2024, 1, 1));
            List<AcaEligibilityPeriod> qualifying = AcaEligibility.QualifyingWindowsForStatus(employee, AsOf, PayDates);
            AcaEligibilityPeriod august = Assert.Single(qualifying);

            Assert.Equal(new DateTime(2026, 8, 1), august.Start);
            Assert.True(august.IsComplete);
            Assert.True(august.Qualifies);
            Assert.Equal(394.29, (double)august.TotalHours, 2);
            Assert.True(AsOf > august.End);
        }

        [Fact]
        public void SpecialPayrollAfterTheWindow_DoesNotCountTowardTheCompletedWindow()
        {
            PayrollHistoryEmployee employee = EmployeeWithPaychecks(60f, new DateTime(2024, 1, 1));
            AddHours(employee, new DateTime(2026, 11, 3), 100f);

            AcaEligibilityPeriod august = AcaEligibility.BuildPeriods(employee, AsOf, PayDates)
                .Single(period => period.Start == new DateTime(2026, 8, 1));
            AcaEligibilityPeriod september = AcaEligibility.BuildPeriods(employee, AsOf, PayDates)
                .Single(period => period.Start == new DateTime(2026, 9, 1));

            Assert.Equal(394.29, (double)august.TotalHours, 2);
            Assert.Equal(361.43, (double)september.TotalHours, 2);
        }

        [Fact]
        public void LowHours_RecommendsPartTimeFromTheLatestCompletedWindow()
        {
            PayrollHistoryEmployee employee = EmployeeWithPaychecks(20f, new DateTime(2026, 8, 1));

            AcaEligibilityPeriod? below = AcaEligibility.CompletedWindowBelowFullTime(employee, AsOf, PayDates);

            Assert.NotNull(below);
            Assert.Equal(new DateTime(2026, 8, 1), below.Start);
            Assert.Equal(new DateTime(2026, 10, 31), below.End);
            Assert.False(below.Qualifies);
        }

        [Fact]
        public void HiredDuringTheWindow_IsNotMovedToPartTime()
        {
            PayrollHistoryEmployee employee = EmployeeWithPaychecks(20f, new DateTime(2026, 8, 15));

            Assert.Null(AcaEligibility.CompletedWindowBelowFullTime(employee, AsOf, PayDates));
        }

        [Fact]
        public void MissingRegularPaycheck_IsNotMovedToPartTime()
        {
            PayrollHistoryEmployee employee = EmployeeWithPaychecks(20f, new DateTime(2026, 8, 1));
            employee.Periods.Remove((new DateTime(2026, 10, 9), Company.VALLEY_BUS_LLC));

            Assert.Null(AcaEligibility.CompletedWindowBelowFullTime(employee, AsOf, PayDates));
        }

        [Fact]
        public void OpenWindowAlreadyAtFullTime_KeepsFullTimeStatus()
        {
            DateTime asOf = new(2026, 11, 20);
            DateTime[] anchors = PayDates.Append(asOf).ToArray();
            PayrollHistoryEmployee employee = EmployeeWithPaychecks(20f, new DateTime(2024, 1, 1), anchors);
            AddHours(employee, asOf, 400f);

            Assert.Null(AcaEligibility.CompletedWindowBelowFullTime(employee, asOf, anchors));
            Assert.Contains(
                AcaEligibility.QualifyingWindowsForStatus(employee, asOf, anchors),
                period => period.Start == new DateTime(2026, 9, 1) && period.Qualifies && !period.IsComplete);
        }

        [Fact]
        public void HistoryLoadStart_IncludesThePaycheckThatTouchesTheCompletedWindow()
        {
            Assert.Equal(new DateTime(2025, 11, 17), AcaEligibility.HistoryLoadStart(new DateTime(2026, 3, 6)));
        }

        private static PayrollHistoryEmployee EmployeeWithPaychecks(float hours, DateTime hireDate, IEnumerable<DateTime>? payDates = null)
        {
            PayrollHistoryEmployee employee = new()
            {
                EmployeeNumber = 1,
                HireDate = hireDate
            };
            foreach (DateTime payDate in payDates ?? PayDates)
            {
                AddHours(employee, payDate, hours);
            }

            return employee;
        }

        private static void AddHours(PayrollHistoryEmployee employee, DateTime payDate, float hours)
        {
            employee.Periods[(payDate.Date, Company.VALLEY_BUS_LLC)] = new PayrollHistoryPeriod
            {
                PayDate = payDate.Date,
                Company = Company.VALLEY_BUS_LLC,
                Source = "Test",
                TotalHours = hours
            };
        }
    }
}
