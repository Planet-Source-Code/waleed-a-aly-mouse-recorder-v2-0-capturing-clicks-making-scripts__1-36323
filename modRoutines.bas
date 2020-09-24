Attribute VB_Name = "modRoutines"

Option Explicit
Public Const SPS As Long = 50 'Recorded Samples Per Second

Public Cursor() As eMouseState
Public i As Long, j As Long, Samples As Long
Public pLB As Boolean, pMB As Boolean, pRB As Boolean
Public PlayOnly As Boolean, UnSaved As Boolean, Esc As Boolean
Public FN As String, RES As String, RT As String, HW As Boolean

Public Sub FreshForm()

    'Initialize Variables & Controls
    UnSaved = False: PlayOnly = False
    UpdateControls True, False, False, False
    UpdateInfo "", "", "", 0, True
    
    Debug.Print CurrentResolution
    'Setting Timers intervals to meet SPS requirement
    frmMain.tmrRecord.Interval = 1000 / SPS
    frmMain.tmrPlay.Interval = 1000 / SPS

End Sub

Public Sub UpdateInfo(ByVal iFN As String, ByVal iRES As String, ByVal iRT As String, ByVal iDU As Long, ByVal iHW As Boolean)

    Dim Info As String
    
    If iFN = "" Then iFN = "-"
    If iRES = "" Then iRES = "-"
    If iRT = "" Then iRT = "-"
    
    Info = "File Name   :  " + Left(iFN, 35) + vbCrLf
    Info = Info + "Recorded On :  " + iRT + vbCrLf
    Info = Info + "Resolution  :  " + iRES + vbCrLf
    Info = Info + "Duration    :  " + CStr(iDU) + " Sec" + vbCrLf
    Info = Info + "Hide Window :  " + CStr(iHW) + vbCrLf
    Info = Info + "File Saved  :  " + CStr(Not UnSaved)
    
    frmMain.lblInfo = Info

End Sub

Public Sub UpdateControls(cmdR As Boolean, cmdP As Boolean, tmrR As Boolean, tmrP As Boolean)

    With frmMain
        
        If PlayOnly Then
            .cmdRecord.Enabled = False
            .chkHide.Enabled = False
        Else
            .cmdRecord.Enabled = cmdR
            .chkHide.Enabled = cmdR
        End If
        .cmdPlay.Enabled = cmdP
        .tmrRecord.Enabled = tmrR
        .tmrPlay.Enabled = tmrP
        
    End With

End Sub

Public Function CurrentResolution() As String

    CurrentResolution = CStr(Screen.Width / Screen.TwipsPerPixelX) + " x " + CStr(Screen.Height / Screen.TwipsPerPixelY)

End Function

Public Sub SaveFile(ByVal FileName As String)
On Error GoTo Error

    Dim Count As Long, TimeNow As String
    TimeNow = CStr(Now)
    
    'Now Save Recorded Mouse Script
    Open FileName For Output Access Write Lock Write As #1
        Write #1, TimeNow, HW, RES, j
        For Count = 1 To j
            Write #1, Cursor(Count).Pos.X, Cursor(Count).Pos.Y, Cursor(Count).LButton, Cursor(Count).MButton, Cursor(Count).RButton
            DoEvents
        Next
    Close #1
    
    UnSaved = False                                   'Now Mouse Recorder Script is Saved
    UpdateInfo FileName, RES, TimeNow, j / SPS, HW    'Ubdate Info Label
    Exit Sub

Error:
Close #1
frmMain.ComDlg.FileName = "": FN = ""
MsgBox "Error, Cannot Open file for Save!", vbCritical, "Error"
End Sub

Public Sub LoadFile(ByVal FileName As String)
On Error GoTo Error

    Dim Count As Long
    
    'Now Load Recorded Mouse Script
    Open FN For Input Access Read Lock Write As #1
        Input #1, RT, HW, RES, j
        If j <= 0 Then GoTo Error
        ReDim Cursor(j)
        For Count = 1 To j
            Input #1, Cursor(Count).Pos.X, Cursor(Count).Pos.Y, Cursor(Count).LButton, Cursor(Count).MButton, Cursor(Count).RButton
            DoEvents
        Next
    Close #1
    
    UpdateInfo FN, RES, RT, j / SPS, HW
    
    PlayOnly = True
    UnSaved = False
    If HW Then frmMain.chkHide = 1 Else frmMain.chkHide = 0
    UpdateControls False, True, False, False
    
    Exit Sub

Error:
Close #1
frmMain.ComDlg.FileName = "": FN = ""
MsgBox "Error, Cannot Open file for PlayBack!", vbCritical, "Error"
End Sub

Public Sub Record()

    GetCursorPos Cursor(i).Pos                                  'Record Cursor Position
    Esc = CBool(GetAsyncKeyState(vbKeyEscape))                  'Monitor the Esc Key in Case User want to skip
    Cursor(i).LButton = CBool(GetAsyncKeyState(vbLeftButton))   'Left Button State
    Cursor(i).MButton = CBool(GetAsyncKeyState(vbMiddleButton)) 'Middle Button State
    Cursor(i).RButton = CBool(GetAsyncKeyState(vbRightButton))  'Right Button State
    
    'Prepare for next Position if Not finished yet Else Stop Recording
    If (i < Samples) And (Not Esc) Then
        i = i + 1
    Else
        j = i - 1
        UpdateControls True, True, False, False
        MsgBox "Recording finished.", vbInformation, "finished!"
        RES = CurrentResolution
        UpdateInfo FN, RES, CStr(Now), j / SPS, HW
        frmMain.Show
    End If

End Sub

Public Sub Play()

    'Position Cursor where it should be
    SetCursorPos Cursor(i).Pos.X, Cursor(i).Pos.Y
    
    'ReGenerate Left Mouse Button Events
    If (Not pLB) And (Cursor(i).LButton) Then mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    If (pLB) And (Not Cursor(i).LButton) Then mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    'ReGenerate Middle Mouse Button Events
    If (Not pMB) And (Cursor(i).MButton) Then mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
    If (pMB) And (Not Cursor(i).MButton) Then mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
    
    'ReGenerate Right Mouse Button Events
    If (Not pRB) And (Cursor(i).RButton) Then mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    If (pRB) And (Not Cursor(i).RButton) Then mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    'Monitor the Esc Key in Case User wants to skip
    Esc = CBool(GetAsyncKeyState(vbKeyEscape))
    
    'Prepare for next Position if Not finished yet Else Stop PlayBack
    If (i < j) And (Not Esc) Then
        pLB = Cursor(i).LButton     'Save previous LMB state
        pMB = Cursor(i).MButton     'Save previous MMB state
        pRB = Cursor(i).RButton     'Save previous RMB state
        i = i + 1                   'Next Sample
    Else
        UpdateControls True, True, False, False
        MsgBox "PlayBack finished.", vbInformation, "finished!"
        frmMain.Show
    End If

End Sub
