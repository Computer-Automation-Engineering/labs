using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Threading;

namespace PPClockVSTO
{
    /// <summary>
    /// PPClock VSTO Add-in for PowerPoint
    /// Professional countdown timer with native .NET integration
    /// No HTML/JavaScript required - pure C# implementation
    /// </summary>
    public partial class ThisAddIn
    {
        #region Fields
        private Timer countdownTimer;
        private int remainingSeconds;
        private int originalSeconds;
        private bool isPaused = false;
        private PPClockForm timerForm;
        private Office.CommandBarButton ppClockButton;
        #endregion

        #region Application Events
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                AddPPClockButton();
                System.Diagnostics.Debug.WriteLine("PPClock VSTO Add-in started successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting PPClock: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (countdownTimer != null)
                {
                    countdownTimer.Stop();
                    countdownTimer.Dispose();
                }
                
                if (timerForm != null)
                {
                    timerForm.Close();
                    timerForm.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PPClock shutdown error: {ex.Message}");
            }
        }
        #endregion

        #region Ribbon Integration
        private void AddPPClockButton()
        {
            try
            {
                // Add button to Home tab
                Office.CommandBar commandBar = Application.CommandBars["Ribbon"];
                
                if (commandBar != null)
                {
                    ppClockButton = (Office.CommandBarButton)commandBar.Controls.Add(
                        Office.MsoControlType.msoControlButton, 
                        Type.Missing, 
                        Type.Missing, 
                        Type.Missing, 
                        true);
                    
                    ppClockButton.Caption = "PPClock Timer";
                    ppClockButton.Tag = "PPClockTimer";
                    ppClockButton.TooltipText = "Professional countdown timer for presentations";
                    ppClockButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(PPClockButton_Click);
                    ppClockButton.Visible = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding PPClock button: {ex.Message}");
            }
        }

        private void PPClockButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ShowPPClockDialog();
        }
        #endregion

        #region PPClock Core Functionality
        private void ShowPPClockDialog()
        {
            try
            {
                if (timerForm == null || timerForm.IsDisposed)
                {
                    timerForm = new PPClockForm(this);
                }
                
                timerForm.Show();
                timerForm.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing PPClock: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void StartCountdown(int minutes, int seconds, string fontSize)
        {
            try
            {
                originalSeconds = (minutes * 60) + seconds;
                remainingSeconds = originalSeconds;
                isPaused = false;

                if (countdownTimer != null)
                {
                    countdownTimer.Stop();
                    countdownTimer.Dispose();
                }

                countdownTimer = new Timer(TimerTick, null, 1000, 1000);
                
                if (timerForm != null)
                {
                    timerForm.UpdateDisplay(FormatTime(remainingSeconds), CalculateProgress());
                    timerForm.SetFontSize(fontSize);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting countdown: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TimerTick(object state)
        {
            try
            {
                if (!isPaused && remainingSeconds > 0)
                {
                    remainingSeconds--;
                    
                    if (timerForm != null && !timerForm.IsDisposed)
                    {
                        timerForm.Invoke(new Action(() =>
                        {
                            timerForm.UpdateDisplay(FormatTime(remainingSeconds), CalculateProgress());
                            
                            if (remainingSeconds <= 10 && remainingSeconds > 0)
                            {
                                timerForm.SetWarningMode(true);
                            }
                            else if (remainingSeconds == 0)
                            {
                                TimerFinished();
                            }
                        }));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Timer tick error: {ex.Message}");
            }
        }

        private void TimerFinished()
        {
            try
            {
                if (countdownTimer != null)
                {
                    countdownTimer.Stop();
                    countdownTimer.Dispose();
                    countdownTimer = null;
                }

                MessageBox.Show("⏰ Time's Up!\n\nYour PPClock countdown has finished.", 
                    "PPClock - Timer Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                if (timerForm != null)
                {
                    timerForm.ResetTimer();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Timer finished error: {ex.Message}");
            }
        }

        public void PauseResumeTimer()
        {
            isPaused = !isPaused;
            
            if (timerForm != null)
            {
                timerForm.UpdatePauseButton(isPaused);
            }
        }

        public void StopTimer()
        {
            try
            {
                if (countdownTimer != null)
                {
                    countdownTimer.Stop();
                    countdownTimer.Dispose();
                    countdownTimer = null;
                }
                
                remainingSeconds = 0;
                isPaused = false;
                
                if (timerForm != null)
                {
                    timerForm.ResetTimer();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Stop timer error: {ex.Message}");
            }
        }

        private string FormatTime(int seconds)
        {
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int secs = seconds % 60;
            
            if (hours > 0)
            {
                return $"{hours:00}:{minutes:00}:{secs:00}";
            }
            else
            {
                return $"{minutes:00}:{secs:00}";
            }
        }

        private int CalculateProgress()
        {
            if (originalSeconds == 0) return 0;
            return (int)(((double)(originalSeconds - remainingSeconds) / originalSeconds) * 100);
        }
        #endregion

        #region PowerPoint Integration
        public void NextSlide()
        {
            try
            {
                PowerPoint.SlideShowWindow slideShow = Application.SlideShowWindows.Count > 0 
                    ? Application.SlideShowWindows[1] : null;
                
                if (slideShow != null)
                {
                    slideShow.View.Next();
                }
                else
                {
                    PowerPoint.Presentation presentation = Application.ActivePresentation;
                    if (presentation != null && presentation.SlideShowSettings != null)
                    {
                        presentation.SlideShowSettings.Run();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error navigating slides: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void PreviousSlide()
        {
            try
            {
                PowerPoint.SlideShowWindow slideShow = Application.SlideShowWindows.Count > 0 
                    ? Application.SlideShowWindows[1] : null;
                
                if (slideShow != null)
                {
                    slideShow.View.Previous();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error navigating slides: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void StartSlideshow()
        {
            try
            {
                PowerPoint.Presentation presentation = Application.ActivePresentation;
                if (presentation != null)
                {
                    presentation.SlideShowSettings.Run();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting slideshow: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void InsertTimerSlide()
        {
            try
            {
                PowerPoint.Presentation presentation = Application.ActivePresentation;
                if (presentation != null)
                {
                    PowerPoint.Slide newSlide = presentation.Slides.Add(
                        presentation.Slides.Count + 1, 
                        PowerPoint.PpSlideLayout.ppLayoutText);
                    
                    // Add title
                    newSlide.Shapes.Title.TextFrame.TextRange.Text = "PPClock Countdown Timer";
                    
                    // Add content
                    newSlide.Shapes.Placeholders[2].TextFrame.TextRange.Text = 
                        "• Professional countdown timer for presentations\n" +
                        "• Pause/Resume functionality\n" +
                        "• Multiple font sizes available\n" +
                        "• Integrated with PowerPoint controls\n" +
                        "• Built with VSTO .NET technology";
                    
                    MessageBox.Show("Timer slide inserted successfully!", "PPClock", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting slide: {ex.Message}", "PPClock Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public string GetSlideInfo()
        {
            try
            {
                PowerPoint.Presentation presentation = Application.ActivePresentation;
                if (presentation != null)
                {
                    int slideCount = presentation.Slides.Count;
                    int currentSlide = 1;
                    
                    try
                    {
                        PowerPoint.SlideShowWindow slideShow = Application.SlideShowWindows.Count > 0 
                            ? Application.SlideShowWindows[1] : null;
                        
                        if (slideShow != null)
                        {
                            currentSlide = slideShow.View.CurrentShowPosition;
                        }
                        else if (Application.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewSlide)
                        {
                            currentSlide = Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                        }
                    }
                    catch
                    {
                        // Use default value if can't determine current slide
                    }
                    
                    return $"Slide {currentSlide} of {slideCount}";
                }
                
                return "No presentation open";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
        #endregion

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}