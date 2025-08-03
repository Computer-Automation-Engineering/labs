' PPClock VBA Add-in for PowerPoint
' Professional countdown timer built entirely in VBA
' No external dependencies - runs natively in PowerPoint
' 
' Installation: Copy this code to PowerPoint VBA Editor (Alt+F11)
' Module: Insert > Module > Paste this code
' UserForm: Insert > UserForm > Design the interface

Option Explicit

' Global variables for timer functionality
Public TimerSeconds As Long
Public OriginalSeconds As Long
Public TimerRunning As Boolean
Public TimerPaused As Boolean
Public TimerForm As PPClockForm

' Main entry point - called from ribbon or macro
Public Sub ShowPPClock()
    On Error GoTo ErrorHandler
    
    ' Initialize timer if not already done
    If TimerForm Is Nothing Then
        Set TimerForm = New PPClockForm
    End If
    
    ' Show the timer form
    TimerForm.Show vbModeless
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error starting PPClock: " & Err.Description, vbCritical, "PPClock Error"
End Sub

' Start countdown timer
Public Sub StartCountdown(Minutes As Integer, Seconds As Integer, FontSize As String)
    On Error GoTo ErrorHandler
    
    ' Calculate total seconds
    OriginalSeconds = (Minutes * 60) + Seconds
    TimerSeconds = OriginalSeconds
    TimerRunning = True
    TimerPaused = False
    
    ' Update form display
    If Not TimerForm Is Nothing Then
        TimerForm.UpdateDisplay FormatTime(TimerSeconds), CalculateProgress()
        TimerForm.SetFontSize FontSize
        TimerForm.ShowTimerPanel
    End If
    
    ' Start the timer
    Application.OnTime Now + TimeValue("00:00:01"), "TimerTick"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error starting countdown: " & Err.Description, vbCritical, "PPClock Error"
End Sub

' Timer tick event - called every second
Public Sub TimerTick()
    On Error GoTo ErrorHandler
    
    ' Only proceed if timer is running and not paused
    If TimerRunning And Not TimerPaused And TimerSeconds > 0 Then
        TimerSeconds = TimerSeconds - 1
        
        ' Update display
        If Not TimerForm Is Nothing Then
            TimerForm.UpdateDisplay FormatTime(TimerSeconds), CalculateProgress()
            
            ' Warning mode for last 10 seconds
            If TimerSeconds <= 10 And TimerSeconds > 0 Then
                TimerForm.SetWarningMode True
            End If
        End If
        
        ' Check if timer finished
        If TimerSeconds <= 0 Then
            TimerFinished
        Else
            ' Schedule next tick
            Application.OnTime Now + TimeValue("00:00:01"), "TimerTick"
        End If
    ElseIf TimerRunning And TimerPaused Then
        ' Reschedule if paused
        Application.OnTime Now + TimeValue("00:00:01"), "TimerTick"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Silently handle errors in timer tick to prevent cascading issues
    Debug.Print "Timer tick error: " & Err.Description
End Sub

' Timer completion handler
Public Sub TimerFinished()
    On Error GoTo ErrorHandler
    
    TimerRunning = False
    TimerPaused = False
    
    ' Show completion message
    MsgBox "⏰ Time's Up!" & vbCrLf & vbCrLf & _
           "Your PPClock countdown has finished.", _
           vbInformation, "PPClock - Timer Finished"
    
    ' Reset form
    If Not TimerForm Is Nothing Then
        TimerForm.ResetTimer
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Timer finished error: " & Err.Description
End Sub

' Pause or resume timer
Public Sub PauseResumeTimer()
    On Error GoTo ErrorHandler
    
    If TimerRunning Then
        TimerPaused = Not TimerPaused
        
        ' Update form
        If Not TimerForm Is Nothing Then
            TimerForm.UpdatePauseButton TimerPaused
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error pausing timer: " & Err.Description, vbCritical, "PPClock Error"
End Sub

' Stop timer completely
Public Sub StopTimer()
    On Error GoTo ErrorHandler
    
    TimerRunning = False
    TimerPaused = False
    TimerSeconds = 0
    
    ' Reset form
    If Not TimerForm Is Nothing Then
        TimerForm.ResetTimer
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error stopping timer: " & Err.Description, vbCritical, "PPClock Error"
End Sub

' Format seconds into MM:SS or HH:MM:SS
Private Function FormatTime(TotalSeconds As Long) As String
    Dim Hours As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer
    
    Hours = TotalSeconds \ 3600
    Minutes = (TotalSeconds Mod 3600) \ 60
    Seconds = TotalSeconds Mod 60
    
    If Hours > 0 Then
        FormatTime = Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
    Else
        FormatTime = Format(Minutes, "00") & ":" & Format(Seconds, "00")
    End If
End Function

' Calculate progress percentage
Private Function CalculateProgress() As Integer
    If OriginalSeconds = 0 Then
        CalculateProgress = 0
    Else
        CalculateProgress = Int(((OriginalSeconds - TimerSeconds) / OriginalSeconds) * 100)
    End If
End Function

' PowerPoint Integration Functions
' =====================================

' Navigate to next slide
Public Sub NextSlide()
    On Error GoTo ErrorHandler
    
    Dim SlideShow As SlideShowWindow
    
    ' Check if slideshow is running
    If ActivePresentation.SlideShowSettings.ShowType <> ppShowTypeSpeaker Then
        If Application.SlideShowWindows.Count > 0 Then
            Set SlideShow = Application.SlideShowWindows(1)
            SlideShow.View.Next
        Else
            ' Start slideshow if not running
            ActivePresentation.SlideShowSettings.Run
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error navigating slides: " & Err.Description, vbExclamation, "PPClock Warning"
End Sub

' Navigate to previous slide
Public Sub PreviousSlide()
    On Error GoTo ErrorHandler
    
    Dim SlideShow As SlideShowWindow
    
    If Application.SlideShowWindows.Count > 0 Then
        Set SlideShow = Application.SlideShowWindows(1)
        SlideShow.View.Previous
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error navigating slides: " & Err.Description, vbExclamation, "PPClock Warning"
End Sub

' Start slideshow
Public Sub StartSlideshow()
    On Error GoTo ErrorHandler
    
    ActivePresentation.SlideShowSettings.Run
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error starting slideshow: " & Err.Description, vbExclamation, "PPClock Warning"
End Sub

' Insert timer slide into presentation
Public Sub InsertTimerSlide()
    On Error GoTo ErrorHandler
    
    Dim NewSlide As Slide
    Dim SlideIndex As Integer
    
    ' Insert new slide after current slide
    SlideIndex = ActivePresentation.Slides.Count + 1
    Set NewSlide = ActivePresentation.Slides.Add(SlideIndex, ppLayoutText)
    
    ' Add title
    NewSlide.Shapes.Title.TextFrame.TextRange.Text = "PPClock Countdown Timer"
    
    ' Add content
    With NewSlide.Shapes.Placeholders(2).TextFrame.TextRange
        .Text = "• Professional countdown timer for presentations" & vbCrLf & _
                "• Pause/Resume functionality available" & vbCrLf & _
                "• Multiple font sizes for visibility" & vbCrLf & _
                "• Integrated PowerPoint controls" & vbCrLf & _
                "• Built with native VBA technology"
    End With
    
    MsgBox "Timer slide inserted successfully!", vbInformation, "PPClock"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error inserting slide: " & Err.Description, vbExclamation, "PPClock Warning"
End Sub

' Get presentation information
Public Function GetSlideInfo() As String
    On Error GoTo ErrorHandler
    
    Dim SlideCount As Integer
    Dim CurrentSlide As Integer
    Dim SlideShow As SlideShowWindow
    
    SlideCount = ActivePresentation.Slides.Count
    CurrentSlide = 1
    
    ' Try to get current slide number
    If Application.SlideShowWindows.Count > 0 Then
        Set SlideShow = Application.SlideShowWindows(1)
        CurrentSlide = SlideShow.View.CurrentShowPosition
    ElseIf ActiveWindow.ViewType = ppViewSlide Then
        CurrentSlide = ActiveWindow.Selection.SlideRange.SlideIndex
    End If
    
    GetSlideInfo = "Slide " & CurrentSlide & " of " & SlideCount
    
    Exit Function
    
ErrorHandler:
    GetSlideInfo = "Error: " & Err.Description
End Function

' Utility function to add PPClock to ribbon (call once)
Public Sub AddPPClockToRibbon()
    ' This would require Ribbon XML customization
    ' For now, users can access via Developer > Macros > ShowPPClock
    MsgBox "PPClock VBA installed successfully!" & vbCrLf & vbCrLf & _
           "Access via:" & vbCrLf & _
           "• Developer tab > Macros > ShowPPClock" & vbCrLf & _
           "• Or assign to Quick Access Toolbar", _
           vbInformation, "PPClock Installation"
End Sub

' Clean up when closing
Public Sub CleanupPPClock()
    TimerRunning = False
    TimerPaused = False
    TimerSeconds = 0
    Set TimerForm = Nothing
End Sub