using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPClockVSTO
{
    /// <summary>
    /// PPClock Timer Form - Windows Forms UI for VSTO Add-in
    /// Provides professional countdown interface with PowerPoint integration
    /// </summary>
    public partial class PPClockForm : Form
    {
        #region Fields
        private ThisAddIn addInInstance;
        private Label timeDisplayLabel;
        private ProgressBar progressBar;
        private Button pauseButton;
        private Button stopButton;
        private Button nextSlideButton;
        private Button previousSlideButton;
        private Button startSlideshowButton;
        private Button insertSlideButton;
        private Label slideInfoLabel;
        private NumericUpDown minutesInput;
        private NumericUpDown secondsInput;
        private ComboBox fontSizeCombo;
        private Button startButton;
        private Panel setupPanel;
        private Panel timerPanel;
        private Label statusLabel;
        #endregion

        #region Constructor
        public PPClockForm(ThisAddIn addIn)
        {
            addInInstance = addIn;
            InitializeComponent();
            SetupForm();
        }
        #endregion

        #region Form Setup
        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "PPClock - Professional Timer";
            this.Size = new Size(400, 600);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;
            this.Icon = SystemIcons.Information;
            
            // Setup Panel
            setupPanel = new Panel();
            setupPanel.Dock = DockStyle.Fill;
            setupPanel.BackColor = Color.FromArgb(248, 249, 250);
            
            // Timer Panel  
            timerPanel = new Panel();
            timerPanel.Dock = DockStyle.Fill;
            timerPanel.BackColor = Color.FromArgb(248, 249, 250);
            timerPanel.Visible = false;
            
            this.Controls.Add(setupPanel);
            this.Controls.Add(timerPanel);
            
            CreateSetupControls();
            CreateTimerControls();
            
            this.ResumeLayout(false);
        }

        private void CreateSetupControls()
        {
            int yPos = 20;
            
            // Title
            Label titleLabel = new Label();
            titleLabel.Text = "PPClock";
            titleLabel.Font = new Font("Segoe UI", 24, FontStyle.Light);
            titleLabel.ForeColor = Color.FromArgb(44, 62, 80);
            titleLabel.TextAlign = ContentAlignment.MiddleCenter;
            titleLabel.Size = new Size(360, 40);
            titleLabel.Location = new Point(20, yPos);
            setupPanel.Controls.Add(titleLabel);
            yPos += 50;
            
            // Subtitle
            Label subtitleLabel = new Label();
            subtitleLabel.Text = "VSTO .NET PowerPoint Add-in";
            subtitleLabel.Font = new Font("Segoe UI", 10);
            subtitleLabel.ForeColor = Color.FromArgb(108, 117, 125);
            subtitleLabel.TextAlign = ContentAlignment.MiddleCenter;
            subtitleLabel.Size = new Size(360, 20);
            subtitleLabel.Location = new Point(20, yPos);
            setupPanel.Controls.Add(subtitleLabel);
            yPos += 40;
            
            // Status
            statusLabel = new Label();
            statusLabel.Text = "✅ Connected to PowerPoint\nReady for professional timing!";
            statusLabel.Font = new Font("Segoe UI", 9);
            statusLabel.ForeColor = Color.FromArgb(0, 102, 204);
            statusLabel.TextAlign = ContentAlignment.MiddleCenter;
            statusLabel.Size = new Size(360, 40);
            statusLabel.Location = new Point(20, yPos);
            statusLabel.BorderStyle = BorderStyle.FixedSingle;
            statusLabel.BackColor = Color.FromArgb(231, 243, 255);
            setupPanel.Controls.Add(statusLabel);
            yPos += 60;
            
            // Minutes input
            Label minutesLabel = new Label();
            minutesLabel.Text = "Minutes:";
            minutesLabel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            minutesLabel.Size = new Size(100, 20);
            minutesLabel.Location = new Point(50, yPos);
            setupPanel.Controls.Add(minutesLabel);
            
            minutesInput = new NumericUpDown();
            minutesInput.Minimum = 0;
            minutesInput.Maximum = 999;
            minutesInput.Value = 5;
            minutesInput.Size = new Size(80, 25);
            minutesInput.Location = new Point(160, yPos);
            minutesInput.Font = new Font("Segoe UI", 12);
            minutesInput.TextAlign = HorizontalAlignment.Center;
            setupPanel.Controls.Add(minutesInput);
            yPos += 40;
            
            // Seconds input
            Label secondsLabel = new Label();
            secondsLabel.Text = "Seconds:";
            secondsLabel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            secondsLabel.Size = new Size(100, 20);
            secondsLabel.Location = new Point(50, yPos);
            setupPanel.Controls.Add(secondsLabel);
            
            secondsInput = new NumericUpDown();
            secondsInput.Minimum = 0;
            secondsInput.Maximum = 59;
            secondsInput.Value = 0;
            secondsInput.Size = new Size(80, 25);
            secondsInput.Location = new Point(160, yPos);
            secondsInput.Font = new Font("Segoe UI", 12);
            secondsInput.TextAlign = HorizontalAlignment.Center;
            setupPanel.Controls.Add(secondsInput);
            yPos += 40;
            
            // Font size selection
            Label fontLabel = new Label();
            fontLabel.Text = "Display Size:";
            fontLabel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            fontLabel.Size = new Size(100, 20);
            fontLabel.Location = new Point(50, yPos);
            setupPanel.Controls.Add(fontLabel);
            
            fontSizeCombo = new ComboBox();
            fontSizeCombo.Items.AddRange(new string[] { "Small", "Medium", "Large", "Extra Large" });
            fontSizeCombo.SelectedIndex = 2; // Large
            fontSizeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            fontSizeCombo.Size = new Size(120, 25);
            fontSizeCombo.Location = new Point(160, yPos);
            fontSizeCombo.Font = new Font("Segoe UI", 10);
            setupPanel.Controls.Add(fontSizeCombo);
            yPos += 50;
            
            // Start button
            startButton = new Button();
            startButton.Text = "Start Countdown";
            startButton.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            startButton.BackColor = Color.FromArgb(0, 120, 212);
            startButton.ForeColor = Color.White;
            startButton.FlatStyle = FlatStyle.Flat;
            startButton.Size = new Size(200, 40);
            startButton.Location = new Point(100, yPos);
            startButton.Click += StartButton_Click;
            setupPanel.Controls.Add(startButton);
            yPos += 60;
            
            // PowerPoint controls group
            GroupBox ppGroup = new GroupBox();
            ppGroup.Text = "PowerPoint Integration";
            ppGroup.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            ppGroup.Size = new Size(320, 120);
            ppGroup.Location = new Point(40, yPos);
            
            insertSlideButton = new Button();
            insertSlideButton.Text = "Insert Timer Slide";
            insertSlideButton.Size = new Size(130, 30);
            insertSlideButton.Location = new Point(20, 25);
            insertSlideButton.Click += InsertSlideButton_Click;
            ppGroup.Controls.Add(insertSlideButton);
            
            Button getInfoButton = new Button();
            getInfoButton.Text = "Get Slide Info";
            getInfoButton.Size = new Size(130, 30);
            getInfoButton.Location = new Point(160, 25);
            getInfoButton.Click += GetInfoButton_Click;
            ppGroup.Controls.Add(getInfoButton);
            
            slideInfoLabel = new Label();
            slideInfoLabel.Text = "Click 'Get Slide Info' for presentation details";
            slideInfoLabel.Font = new Font("Segoe UI", 9);
            slideInfoLabel.ForeColor = Color.FromArgb(108, 117, 125);
            slideInfoLabel.Size = new Size(280, 40);
            slideInfoLabel.Location = new Point(20, 65);
            slideInfoLabel.TextAlign = ContentAlignment.MiddleCenter;
            ppGroup.Controls.Add(slideInfoLabel);
            
            setupPanel.Controls.Add(ppGroup);
        }

        private void CreateTimerControls()
        {
            int yPos = 20;
            
            // Title
            Label titleLabel = new Label();
            titleLabel.Text = "PPClock";
            titleLabel.Font = new Font("Segoe UI", 20, FontStyle.Light);
            titleLabel.ForeColor = Color.FromArgb(44, 62, 80);
            titleLabel.TextAlign = ContentAlignment.MiddleCenter;
            titleLabel.Size = new Size(360, 35);
            titleLabel.Location = new Point(20, yPos);
            timerPanel.Controls.Add(titleLabel);
            yPos += 45;
            
            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Size = new Size(320, 10);
            progressBar.Location = new Point(40, yPos);
            progressBar.Style = ProgressBarStyle.Continuous;
            timerPanel.Controls.Add(progressBar);
            yPos += 30;
            
            // Time display
            timeDisplayLabel = new Label();
            timeDisplayLabel.Text = "05:00";
            timeDisplayLabel.Font = new Font("Segoe UI", 48, FontStyle.Bold);
            timeDisplayLabel.ForeColor = Color.FromArgb(44, 62, 80);
            timeDisplayLabel.TextAlign = ContentAlignment.MiddleCenter;
            timeDisplayLabel.Size = new Size(360, 80);
            timeDisplayLabel.Location = new Point(20, yPos);
            timeDisplayLabel.BackColor = Color.FromArgb(248, 249, 250);
            timeDisplayLabel.BorderStyle = BorderStyle.FixedSingle;
            timerPanel.Controls.Add(timeDisplayLabel);
            yPos += 100;
            
            // Control buttons
            pauseButton = new Button();
            pauseButton.Text = "Pause";
            pauseButton.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            pauseButton.BackColor = Color.FromArgb(255, 140, 0);
            pauseButton.ForeColor = Color.White;
            pauseButton.FlatStyle = FlatStyle.Flat;
            pauseButton.Size = new Size(100, 35);
            pauseButton.Location = new Point(70, yPos);
            pauseButton.Click += PauseButton_Click;
            timerPanel.Controls.Add(pauseButton);
            
            stopButton = new Button();
            stopButton.Text = "Stop";
            stopButton.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            stopButton.BackColor = Color.FromArgb(220, 53, 69);
            stopButton.ForeColor = Color.White;
            stopButton.FlatStyle = FlatStyle.Flat;
            stopButton.Size = new Size(100, 35);
            stopButton.Location = new Point(230, yPos);
            stopButton.Click += StopButton_Click;
            timerPanel.Controls.Add(stopButton);
            yPos += 55;
            
            // PowerPoint controls
            GroupBox ppControlGroup = new GroupBox();
            ppControlGroup.Text = "Presentation Controls";
            ppControlGroup.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            ppControlGroup.Size = new Size(320, 100);
            ppControlGroup.Location = new Point(40, yPos);
            
            nextSlideButton = new Button();
            nextSlideButton.Text = "Next Slide";
            nextSlideButton.Size = new Size(90, 30);
            nextSlideButton.Location = new Point(20, 25);
            nextSlideButton.Click += NextSlideButton_Click;
            ppControlGroup.Controls.Add(nextSlideButton);
            
            previousSlideButton = new Button();
            previousSlideButton.Text = "Previous";
            previousSlideButton.Size = new Size(90, 30);
            previousSlideButton.Location = new Point(120, 25);
            ppControlGroup.Controls.Add(previousSlideButton);
            previousSlideButton.Click += PreviousSlideButton_Click;
            
            startSlideshowButton = new Button();
            startSlideshowButton.Text = "Start Show";
            startSlideshowButton.Size = new Size(90, 30);
            startSlideshowButton.Location = new Point(220, 25);
            startSlideshowButton.Click += StartSlideshowButton_Click;
            ppControlGroup.Controls.Add(startSlideshowButton);
            
            Label statusTimerLabel = new Label();
            statusTimerLabel.Text = "Timer running - Focus on your presentation!";
            statusTimerLabel.Font = new Font("Segoe UI", 9);
            statusTimerLabel.ForeColor = Color.FromArgb(0, 102, 204);
            statusTimerLabel.Size = new Size(280, 20);
            statusTimerLabel.Location = new Point(20, 65);
            statusTimerLabel.TextAlign = ContentAlignment.MiddleCenter;
            ppControlGroup.Controls.Add(statusTimerLabel);
            
            timerPanel.Controls.Add(ppControlGroup);
        }

        private void SetupForm()
        {
            // Additional form setup if needed
        }
        #endregion

        #region Event Handlers
        private void StartButton_Click(object sender, EventArgs e)
        {
            int minutes = (int)minutesInput.Value;
            int seconds = (int)secondsInput.Value;
            
            if (minutes == 0 && seconds == 0)
            {
                MessageBox.Show("Please enter a valid time greater than 0.", "PPClock", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            setupPanel.Visible = false;
            timerPanel.Visible = true;
            
            addInInstance.StartCountdown(minutes, seconds, fontSizeCombo.SelectedItem.ToString());
        }

        private void PauseButton_Click(object sender, EventArgs e)
        {
            addInInstance.PauseResumeTimer();
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            addInInstance.StopTimer();
        }

        private void NextSlideButton_Click(object sender, EventArgs e)
        {
            addInInstance.NextSlide();
        }

        private void PreviousSlideButton_Click(object sender, EventArgs e)
        {
            addInInstance.PreviousSlide();
        }

        private void StartSlideshowButton_Click(object sender, EventArgs e)
        {
            addInInstance.StartSlideshow();
        }

        private void InsertSlideButton_Click(object sender, EventArgs e)
        {
            addInInstance.InsertTimerSlide();
        }

        private void GetInfoButton_Click(object sender, EventArgs e)
        {
            slideInfoLabel.Text = addInInstance.GetSlideInfo();
        }
        #endregion

        #region Public Methods
        public void UpdateDisplay(string timeText, int progress)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string, int>(UpdateDisplay), timeText, progress);
                return;
            }
            
            timeDisplayLabel.Text = timeText;
            progressBar.Value = Math.Min(progress, 100);
        }

        public void SetFontSize(string fontSize)
        {
            Font currentFont = timeDisplayLabel.Font;
            float newSize = 48; // Default
            
            switch (fontSize.ToLower())
            {
                case "small":
                    newSize = 32;
                    break;
                case "medium":
                    newSize = 40;
                    break;
                case "large":
                    newSize = 48;
                    break;
                case "extra large":
                    newSize = 56;
                    break;
            }
            
            timeDisplayLabel.Font = new Font(currentFont.FontFamily, newSize, currentFont.Style);
        }

        public void SetWarningMode(bool warning)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(SetWarningMode), warning);
                return;
            }
            
            if (warning)
            {
                timeDisplayLabel.ForeColor = Color.FromArgb(220, 53, 69); // Red
                timeDisplayLabel.BackColor = Color.FromArgb(255, 243, 205); // Light yellow
            }
            else
            {
                timeDisplayLabel.ForeColor = Color.FromArgb(44, 62, 80); // Dark blue
                timeDisplayLabel.BackColor = Color.FromArgb(248, 249, 250); // Light gray
            }
        }

        public void UpdatePauseButton(bool isPaused)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(UpdatePauseButton), isPaused);
                return;
            }
            
            if (isPaused)
            {
                pauseButton.Text = "Resume";
                pauseButton.BackColor = Color.FromArgb(40, 167, 69); // Green
            }
            else
            {
                pauseButton.Text = "Pause";
                pauseButton.BackColor = Color.FromArgb(255, 140, 0); // Orange
            }
        }

        public void ResetTimer()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(ResetTimer));
                return;
            }
            
            setupPanel.Visible = true;
            timerPanel.Visible = false;
            
            timeDisplayLabel.ForeColor = Color.FromArgb(44, 62, 80);
            timeDisplayLabel.BackColor = Color.FromArgb(248, 249, 250);
            pauseButton.Text = "Pause";
            pauseButton.BackColor = Color.FromArgb(255, 140, 0);
            progressBar.Value = 0;
        }
        #endregion
    }
}