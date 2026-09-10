using System.ComponentModel;
using System.Globalization;

namespace PayrollProcessor
{
    internal sealed class PayrollHistoryParserForm : Form
    {
        private readonly PayrollHistoryCatalog catalog;
        private readonly Label statusLabel = new();
        private readonly TextBox searchBox = new();
        private readonly Button searchButton = new();
        private readonly ListView resultsList = new();
        private readonly Label selectedLabel = new();
        private readonly RadioButton lastSixRadio = new();
        private readonly RadioButton dateRangeRadio = new();
        private readonly DateTimePicker startPicker = new();
        private readonly DateTimePicker endPicker = new();
        private readonly Button submitButton = new();
        private readonly Button housingButton = new();
        private PayrollHistoryEmployee? selectedEmployee;
        private bool busy;

        public PayrollHistoryParserForm(PayrollHistoryCatalog catalog)
        {
            this.catalog = catalog;
            Text = "Payroll History Parser";
            Width = 760;
            Height = 620;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(680, 520);
            Font = new Font("Segoe UI", 9F);

            Label searchLabel = new()
            {
                Text = "Search by employee number, first name, or last name:",
                AutoSize = true,
                Location = new Point(12, 12)
            };
            searchBox.Location = new Point(12, 35);
            searchBox.Width = 430;
            searchBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            searchBox.KeyDown += SearchBoxKeyDown;
            searchBox.TextChanged += (_, _) => ApplySearch();

            searchButton.Text = "Search";
            searchButton.Location = new Point(450, 33);
            searchButton.Width = 90;
            searchButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            searchButton.Click += (_, _) => ApplySearch();

            resultsList.Location = new Point(12, 70);
            resultsList.Size = new Size(716, 250);
            resultsList.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            resultsList.View = View.Details;
            resultsList.FullRowSelect = true;
            resultsList.HideSelection = false;
            resultsList.MultiSelect = false;
            resultsList.Columns.Add("Employee #", 100);
            resultsList.Columns.Add("Last Name", 200);
            resultsList.Columns.Add("First Name", 180);
            resultsList.Columns.Add("Last Paid", 120);
            resultsList.SelectedIndexChanged += ResultsSelected;
            resultsList.DoubleClick += (_, _) => SubmitIfReady();

            selectedLabel.AutoSize = true;
            selectedLabel.Location = new Point(12, 330);
            selectedLabel.Text = "Selected employee: none";

            lastSixRadio.Text = "Last 6 pay periods (12 weeks)";
            lastSixRadio.AutoSize = true;
            lastSixRadio.Location = new Point(12, 365);
            lastSixRadio.Checked = true;
            lastSixRadio.CheckedChanged += RangeModeChanged;

            dateRangeRadio.Text = "Date range";
            dateRangeRadio.AutoSize = true;
            dateRangeRadio.Location = new Point(12, 395);

            Label fromLabel = new()
            {
                Text = "From",
                AutoSize = true,
                Location = new Point(140, 397)
            };
            startPicker.Location = new Point(180, 393);
            startPicker.Width = 160;

            Label toLabel = new()
            {
                Text = "To",
                AutoSize = true,
                Location = new Point(360, 397)
            };
            endPicker.Location = new Point(385, 393);
            endPicker.Width = 160;
            ApplyLastSixWindowDefaults();

            submitButton.Text = "Submit";
            submitButton.Location = new Point(12, 440);
            submitButton.Size = new Size(140, 32);
            submitButton.Click += (_, _) => SubmitIfReady();

            housingButton.Text = "Housing Request";
            housingButton.Location = new Point(160, 440);
            housingButton.Size = new Size(160, 32);
            housingButton.Click += (_, _) => WriteHousingRequest();

            statusLabel.AutoSize = true;
            statusLabel.Location = new Point(12, 490);
            statusLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            statusLabel.Text = "Loading employees from WfnEmployees.xlsx...";

            Controls.AddRange(new Control[]
            {
                searchLabel, searchBox, searchButton, resultsList, selectedLabel,
                lastSixRadio, dateRangeRadio, fromLabel, startPicker, toLabel, endPicker,
                submitButton, housingButton, statusLabel
            });

            RangeModeChanged(this, EventArgs.Empty);
            Shown += OnShown;
            FormClosed += (_, _) => PayrollHistoryReportWriter.CleanupOutputFolder();
            PayrollHistoryReportWriter.CleanupOutputFolder();
        }

        private void OnShown(object? sender, EventArgs e)
        {
            UseWaitCursor = true;
            searchBox.Enabled = false;
            searchButton.Enabled = false;
            submitButton.Enabled = false;
            housingButton.Enabled = false;
            BackgroundWorker worker = new();
            worker.DoWork += (_, _) => catalog.LoadEmployees();
            worker.RunWorkerCompleted += (_, args) =>
            {
                UseWaitCursor = false;
                searchBox.Enabled = true;
                searchButton.Enabled = true;
                submitButton.Enabled = true;
                housingButton.Enabled = true;
                if (args.Error != null)
                {
                    statusLabel.Text = "Load failed.";
                    MessageBox.Show(this, args.Error.Message, "Payroll History Parser",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                statusLabel.Text = "Loaded " + catalog.Employees.Count
                    + " employees from WfnEmployees.xlsx. Payroll files load when you run a report.";
                if (catalog.LoadWarnings.Count > 0)
                {
                    MessageBox.Show(this, string.Join(Environment.NewLine, catalog.LoadWarnings),
                        "Payroll History Parser", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                ApplySearch();
                searchBox.Focus();
            };
            worker.RunWorkerAsync();
        }

        private void SearchBoxKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ApplySearch();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void ApplySearch()
        {
            resultsList.BeginUpdate();
            resultsList.Items.Clear();
            int shown = 0;
            foreach (PayrollHistoryEmployee employee in catalog.Search(searchBox.Text))
            {
                ListViewItem item = new(employee.EmployeeNumber.ToString(CultureInfo.InvariantCulture));
                item.SubItems.Add(employee.LastName);
                item.SubItems.Add(employee.FirstName);
                item.SubItems.Add(employee.LastPaidDate?.ToString("M/d/yyyy") ?? "");
                item.Tag = employee;
                resultsList.Items.Add(item);
                shown++;
                if (shown >= 500)
                {
                    break;
                }
            }
            resultsList.EndUpdate();
            if (resultsList.Items.Count == 1)
            {
                resultsList.Items[0].Selected = true;
            }
        }

        private void ResultsSelected(object? sender, EventArgs e)
        {
            if (resultsList.SelectedItems.Count == 0)
            {
                selectedEmployee = null;
                selectedLabel.Text = "Selected employee: none";
                return;
            }

            selectedEmployee = resultsList.SelectedItems[0].Tag as PayrollHistoryEmployee;
            if (selectedEmployee == null)
            {
                selectedLabel.Text = "Selected employee: none";
                return;
            }

            selectedLabel.Text = "Selected employee: #" + selectedEmployee.EmployeeNumber + "  "
                + selectedEmployee.DisplayName;
        }

        private void ApplyLastSixWindowDefaults()
        {
            (DateTime Start, DateTime End)? window = PayPeriodSchedule.LastRegularPayPeriodWindow();
            if (window.HasValue)
            {
                startPicker.Value = window.Value.Start;
                endPicker.Value = window.Value.End;
                return;
            }

            startPicker.Value = DateTime.Today.AddMonths(-3);
            endPicker.Value = DateTime.Today;
        }

        private void RangeModeChanged(object? sender, EventArgs e)
        {
            bool useRange = dateRangeRadio.Checked;
            startPicker.Enabled = useRange;
            endPicker.Enabled = useRange;
        }

        private bool WarnIfMostRecentPayrollMissing()
        {
            string? error = catalog.MostRecentPayrollMissingError;
            if (error == null)
            {
                return false;
            }

            Program.Log(error, true);
            statusLabel.Text = error;
            return true;
        }

        private void SubmitIfReady()
        {
            LoadPayrollThen(housing: false, () =>
            {
                List<PayrollHistoryPeriod> periods = catalog.GetPeriods(selectedEmployee!, lastSixRadio.Checked,
                    startPicker.Value.Date, endPicker.Value.Date);
                if (periods.Count == 0)
                {
                    MessageBox.Show(this, "No pay history was found for that employee in the selected range.",
                        "Payroll History Parser", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string path = PayrollHistoryReportWriter.Write(selectedEmployee!, periods, lastSixRadio.Checked,
                    startPicker.Value.Date, endPicker.Value.Date);
                statusLabel.Text = "Loaded " + catalog.AdpFileCount + " ADP and " + catalog.IsolvedFileCount
                    + " iSolved file(s). Wrote " + path;
            });
        }

        private void WriteHousingRequest()
        {
            LoadPayrollThen(housing: true, () =>
            {
                string path = PayrollHistoryReportWriter.WriteHousingRequest(selectedEmployee!,
                    catalog.GetPeriods(selectedEmployee!, true, DateTime.MinValue, DateTime.MaxValue));
                statusLabel.Text = "Loaded " + catalog.AdpFileCount + " ADP and " + catalog.IsolvedFileCount
                    + " iSolved file(s). Wrote " + path;
            });
        }

        private void LoadPayrollThen(bool housing, Action onReady)
        {
            if (selectedEmployee == null)
            {
                MessageBox.Show(this, "Select an employee from the search results first.",
                    "Payroll History Parser", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (busy)
            {
                return;
            }

            PayrollHistoryEmployee employee = selectedEmployee;
            bool lastSix = housing || lastSixRadio.Checked;
            DateTime start = startPicker.Value.Date;
            DateTime end = endPicker.Value.Date;
            int previousWarningCount = catalog.LoadWarnings.Count;
            SetBusy(true);
            statusLabel.Text = "Loading payroll files...";
            BackgroundWorker worker = new();
            worker.DoWork += (_, _) => catalog.EnsurePayrollLoaded(employee, new PayrollLoadNeed
            {
                LastSixPayPeriods = lastSix,
                HousingYears = housing,
                RangeStart = start,
                RangeEnd = end
            });
            worker.RunWorkerCompleted += (_, args) =>
            {
                SetBusy(false);
                if (args.Error != null)
                {
                    statusLabel.Text = "Payroll load failed.";
                    MessageBox.Show(this, args.Error.Message, "Payroll History Parser",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                RestoreSelection(employee);
                List<string> newWarnings = catalog.LoadWarnings.Skip(previousWarningCount).ToList();
                if (newWarnings.Count > 0)
                {
                    MessageBox.Show(this, string.Join(Environment.NewLine, newWarnings),
                        "Payroll History Parser", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                if (lastSix && WarnIfMostRecentPayrollMissing())
                {
                    return;
                }

                try
                {
                    onReady();
                }
                catch (Exception exception)
                {
                    MessageBox.Show(this, exception.Message, "Payroll History Parser",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            worker.RunWorkerAsync();
        }

        private void SetBusy(bool value)
        {
            busy = value;
            UseWaitCursor = value;
            searchBox.Enabled = !value;
            searchButton.Enabled = !value;
            resultsList.Enabled = !value;
            submitButton.Enabled = !value;
            housingButton.Enabled = !value;
            lastSixRadio.Enabled = !value;
            dateRangeRadio.Enabled = !value;
            startPicker.Enabled = !value && dateRangeRadio.Checked;
            endPicker.Enabled = !value && dateRangeRadio.Checked;
        }

        private void RestoreSelection(PayrollHistoryEmployee employee)
        {
            ApplySearch();
            foreach (ListViewItem item in resultsList.Items)
            {
                if (item.Tag is PayrollHistoryEmployee listed && listed.EmployeeNumber == employee.EmployeeNumber)
                {
                    item.Selected = true;
                    item.EnsureVisible();
                    break;
                }
            }
            selectedEmployee = employee;
            selectedLabel.Text = "Selected employee: #" + employee.EmployeeNumber + "  " + employee.DisplayName;
        }
    }
}
