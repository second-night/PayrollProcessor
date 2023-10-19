namespace PayrollProcessor
{
    public partial class PrintForm : Form
    {
        protected DateTime LastDayWeekTwo;
        private Button button2;
        private TextBox textBox1;
        protected bool YesNoButton;
        private Button exirtButton;
        protected string InputText;
        public PrintForm(string message)
        {
            InitializeComponent();
            label1.Text = message;
            this.Controls.Add(label1);

        }

        private void InitializeComponent()
        {
            label1 = new Label();
            button1 = new Button();
            dateTimePicker1 = new DateTimePicker();
            button2 = new Button();
            textBox1 = new TextBox();
            exirtButton = new Button();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 45);
            label1.Name = "label1";
            label1.Size = new Size(38, 15);
            label1.TabIndex = 0;
            label1.Text = "label1";
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom;
            button1.Location = new Point(24, 217);
            button1.Name = "button1";
            button1.Size = new Size(94, 29);
            button1.TabIndex = 1;
            button1.Text = "No";
            button1.UseVisualStyleBackColor = true;
            button1.Visible = false;
            button1.Click += ButtonPressedFalse;
            button1.KeyDown += ButtonPressedFalse;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(43, 106);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(200, 23);
            dateTimePicker1.TabIndex = 2;
            dateTimePicker1.Visible = false;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Bottom;
            button2.Location = new Point(149, 217);
            button2.Name = "button2";
            button2.Size = new Size(94, 29);
            button2.TabIndex = 2;
            button2.Text = "Okay";
            button2.UseVisualStyleBackColor = true;
            button2.Click += ButtonPressed;
            button2.KeyDown += ButtonPressed;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(90, 172);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(100, 23);
            textBox1.TabIndex = 3;
            textBox1.Visible = false;
            // 
            // exirtButton
            // 
            exirtButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            exirtButton.Location = new Point(178, 12);
            exirtButton.Name = "exirtButton";
            exirtButton.Size = new Size(94, 29);
            exirtButton.TabIndex = 4;
            exirtButton.Text = "Exit Program";
            exirtButton.UseVisualStyleBackColor = true;
            exirtButton.Click += ExitButtonPressed;
            exirtButton.KeyDown += ExitButtonPressed;
            // 
            // PrintForm
            // 
            AutoSize = true;
            ClientSize = new Size(284, 261);
            Controls.Add(exirtButton);
            Controls.Add(textBox1);
            Controls.Add(button2);
            Controls.Add(dateTimePicker1);
            Controls.Add(button1);
            Controls.Add(label1);
            Name = "PrintForm";
            ResumeLayout(false);
            PerformLayout();
        }

        private void ExitButtonPressed(object? sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void ButtonPressed(object? sender, EventArgs e)
        {
            if (dateTimePicker1 != null)
            {
                LastDayWeekTwo = dateTimePicker1.Value;
            }
            YesNoButton = true;
            InputText = textBox1.Text;
            this.Close();
        }

        private void ButtonPressedFalse(object? sender, EventArgs e)
        {
            YesNoButton = false;
            this.Close();
        }

        private Label label1;
        private DateTimePicker dateTimePicker1;
        private Button button1;

        public static bool InputDateTime(string message, out DateTime dateTime)
        {
            PrintForm form = new PrintForm(message);
            form.dateTimePicker1.Visible = true;
            form.button2.Text = "Yes";
            form.button1.Visible = true;
            Application.Run(form);
            dateTime = form.LastDayWeekTwo;
            return form.YesNoButton;
        }

        public static bool InputBool(string message)
        {
            PrintForm form = new PrintForm(message);
            form.button2.Text = "Yes";
            form.button1.Visible = true;
            Application.Run(form);
            return form.YesNoButton;
        }

        public static float InputNumber(string message)
        {
            PrintForm form = new PrintForm(message);
            form.button2.Text = "Yes";
            form.button1.Visible = true;
            form.textBox1.Visible = true;
            Application.Run(form);
            if (float.TryParse(form.InputText, out float number))
            {
                return number;
            }
            return 0f;
        }
    }
}