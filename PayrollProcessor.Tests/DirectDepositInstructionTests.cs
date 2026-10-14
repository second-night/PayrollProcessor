using Xunit;

namespace PayrollProcessor.Tests
{
    public class DirectDepositInstructionTests
    {
        [Fact]
        public void EntireNetPay_IsValidAndDepositsTheFullNet()
        {
            Dictionary<string, object> account = Account("Entire Net Pay", amount: "50", percent: "10");

            Assert.Empty(ExcelWorker.GetDirectDepositInstructionErrors(new[] { account }));

            ExcelWorker.ApplyDirectDepositInstruction(account);
            Dictionary<string, string> fields = BankFields(account);

            Assert.Equal("Y", fields["Bank Full Deposit Flag"]);
            Assert.Equal("", fields["Bank Deposit Deduction Amount"]);
            Assert.Equal("", fields["Bank Deposit Percent Net"]);
            Assert.Equal("", account["Amount"]);
            Assert.Equal("", account["Percent"]);
        }

        [Fact]
        public void PartialNet_RequiresADollarAmount()
        {
            Dictionary<string, object> entire = Account("Entire Net Pay");
            Dictionary<string, object> partial = Account("Specific Dollar Amount", amount: "100");

            Assert.Empty(ExcelWorker.GetDirectDepositInstructionErrors(new[] { partial, entire }));

            Dictionary<string, string> fields = BankFields(partial);
            Assert.Equal("N", fields["Bank Full Deposit Flag"]);
            Assert.Equal("100", fields["Bank Deposit Deduction Amount"]);
            Assert.Equal("", fields["Bank Deposit Percent Net"]);
        }

        [Theory]
        [InlineData("")]
        [InlineData("All")]
        [InlineData("0")]
        public void PartialNet_WithoutANumber_LogsAnErrorAndIsNotTreatedAsFullNet(string amount)
        {
            Dictionary<string, object> account = Account("Specific Dollar Amount", amount: amount, percent: "100");

            string error = Assert.Single(ExcelWorker.GetDirectDepositInstructionErrors(new[] { account }),
                message => message.Contains("partial net"));
            Assert.Contains("not a number", error);
            Assert.Contains("full net", ExcelWorker.GetDirectDepositInstructionErrors(new[] { account }).Last());

            Dictionary<string, string> fields = BankFields(account);
            Assert.Equal("N", fields["Bank Full Deposit Flag"]);
            Assert.Equal("", fields["Bank Deposit Deduction Amount"]);
            Assert.Equal("", fields["Bank Deposit Percent Net"]);
        }

        [Fact]
        public void PartialNetPercentage_RequiresAPercentage()
        {
            Dictionary<string, object> entire = Account("Entire Net Pay");
            Dictionary<string, object> partial = Account("Net Amount", percent: "30%");

            Assert.Empty(ExcelWorker.GetDirectDepositInstructionErrors(new[] { partial, entire }));

            Dictionary<string, string> fields = BankFields(partial);
            Assert.Equal("N", fields["Bank Full Deposit Flag"]);
            Assert.Equal("", fields["Bank Deposit Deduction Amount"]);
            Assert.Equal("30", fields["Bank Deposit Percent Net"]);
        }

        [Theory]
        [InlineData("")]
        [InlineData("All")]
        [InlineData("800")]
        [InlineData("0")]
        public void PartialNetPercentage_WithoutAPercentage_LogsAnError(string percent)
        {
            Dictionary<string, object> entire = Account("Entire Net Pay");
            Dictionary<string, object> partial = Account("Net Amount", percent: percent);

            string error = Assert.Single(ExcelWorker.GetDirectDepositInstructionErrors(new[] { partial, entire }));
            Assert.Contains("partial net percentage", error);
            Assert.Contains("not a percentage", error);
        }

        [Fact]
        public void AccountsThatOnlySplitTheNet_RequireOneFullNetAccount()
        {
            Dictionary<string, object> first = Account("Net Amount", percent: "30%");
            Dictionary<string, object> second = Account("Net Amount", percent: "70%");

            string error = Assert.Single(ExcelWorker.GetDirectDepositInstructionErrors(new[] { first, second }));
            Assert.Contains("Entire Net Pay", error);
        }

        [Fact]
        public void PartialNetWording_IsAcceptedWhenTheAmountIsPresent()
        {
            Dictionary<string, object> entire = Account("Full Net");
            Dictionary<string, object> dollars = Account("Partial Net", amount: "25");
            Dictionary<string, object> percent = Account("Partial Net Percentage", percent: "40");

            Assert.Empty(ExcelWorker.GetDirectDepositInstructionErrors(new[] { dollars, percent, entire }));
            Assert.Equal("25", BankFields(dollars)["Bank Deposit Deduction Amount"]);
            Assert.Equal("40", BankFields(percent)["Bank Deposit Percent Net"]);
        }

        [Fact]
        public void MissingOrUnknownInstructions_AreErrors()
        {
            Dictionary<string, object> missing = Account("");
            Dictionary<string, object> unknown = Account("Remainder");

            List<string> errors = ExcelWorker.GetDirectDepositInstructionErrors(new[] { missing, unknown });

            Assert.Contains(errors, message => message.Contains("Account 1") && message.Contains("missing deposit instructions"));
            Assert.Contains(errors, message => message.Contains("Account 2") && message.Contains("unrecognized"));
            Assert.Contains(errors, message => message.Contains("Entire Net Pay"));
        }

        [Fact]
        public void EmptyAccountSlot_IsIgnored()
        {
            Dictionary<string, object> empty = new();
            Dictionary<string, object> entire = Account("Entire Net Pay");

            Assert.Empty(ExcelWorker.GetDirectDepositInstructionErrors(new[] { empty, entire }));
        }

        private static Dictionary<string, object> Account(string instruction, string amount = "", string percent = "")
        {
            return new Dictionary<string, object>
            {
                ["RoutingNumber"] = "021000021",
                ["AccountNumber"] = "123456789",
                ["DepositInstructions"] = instruction,
                ["Amount"] = amount,
                ["Percent"] = percent,
                ["AccountType"] = "C"
            };
        }

        private static Dictionary<string, string> BankFields(Dictionary<string, object> account)
        {
            Dictionary<string, string> fields = new();
            Assert.True(ExcelWorker.TryAddAdpBankDepositFields(fields, account, 1, new Employee(1, "Test Employee")));
            return fields;
        }
    }
}
