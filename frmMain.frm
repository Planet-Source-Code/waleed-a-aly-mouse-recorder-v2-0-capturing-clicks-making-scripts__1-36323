VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Recorder v2.0"
   ClientHeight    =   3435
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRecord 
      Left            =   2160
      Top             =   2340
   End
   Begin VB.Timer tmrPlay 
      Left            =   2580
      Top             =   2340
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "Hide window when recording (Recommended)"
      Height          =   195
      Left            =   780
      TabIndex        =   3
      Top             =   1500
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      Height          =   495
      Left            =   1740
      TabIndex        =   2
      Top             =   240
      Width           =   1755
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "PlayBack"
      Height          =   495
      Left            =   1740
      TabIndex        =   1
      Top             =   780
      Width           =   1755
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   2340
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      Filter          =   "Mouse Recorder Text Files|*.txt"
      InitDir         =   "C:\"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "To stop recording or playback at any time, just press ESC"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   4035
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Sa&ve As"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************'
'                                                       '
'   By:         Waleed A. Aly                           '
'   ASL:        [20 M Egypt]                            '
'   eMail:      wa_aly@tdcspace.dk                      '
'   Thanks to:  www.allapi.net                          '
'                                                       '
'     Please eMail me any Comments and|or Suggestions.  '
'   I hope you like my work and think is usefull !  :)  '
'   I'd love to know how many people are using my Code  '
'   so you can always eMail me if you are goin' to use  '
'   it :)                                               '
'                                      Thanks.          '
'                                                       '
'*******************************************************'

Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Form_Initialize()

    InitCommonControls  'XP Style Support

End Sub

Private Sub Form_Load()

    FreshForm           'Initiate Variables & Controls

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not UnSaved Then End
    If UnSaved And MsgBox("Script is not saved yet. Are you sure you want to exit ?", vbExclamation Or vbYesNo, "Script not saved") = vbYes Then End
    Cancel = 1

End Sub

Private Sub cmdRecord_Click()
On Error GoTo Error

    'Calculate total number of Samples to be Recorded
    Samples = SPS * Val(InputBox("Number of seconds to record :", "Do not Change Screen Resolution while Recording"))
    
    If Samples <= 0 Then Exit Sub            'Abort Recording if nothing to record
    i = 0                                    'Initiate Samples Counter
    ReDim Cursor(Samples)                    'Resize Cursor State Array
    UpdateControls False, False, True, False 'Update Controls
    UnSaved = True                           'Script is not saved yet
    HW = CBool(chkHide.Value)                'Save the Option for applying at PlayBack
    If HW Then Me.Hide                       'Should the Window Hide while Recording ?
    Exit Sub

Error:
MsgBox "Error, Recording Time Too Long!", vbCritical, "Error"
End Sub

Private Sub cmdPlay_Click()

    'Check to see whether current screen resolution matches the recorded file resolution
    If RES <> CurrentResolution Then
        If MsgBox("Your current screen resolution does not match the resolution of the file to be played back. Are you sure you want to Continue ?", vbCritical Or vbYesNo, "Resolution does not match") = vbNo Then Exit Sub
    End If
    
    i = 1                                    'Initiate PlayBack
    If j <= 0 Then Exit Sub                  'Abort PlayBack if nothing to play
    UpdateControls False, False, False, True 'Update Controls
    If HW Then Me.Hide                       'Should the Window Hide while PlayBack ?

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileNew_Click()

    FreshForm

End Sub

Private Sub mnuFileOpen_Click()

    ComDlg.ShowOpen             'Show file Open Dialog
    FN = ComDlg.FileName        'Get the Chosen File Name
    If FN = "" Then Exit Sub    'Make Sure User have selected a file
    LoadFile FN                 'Now Load the file

End Sub

Private Sub mnuFileSave_Click()

    If FN = "" Then mnuFileSaveAs_Click: Exit Sub
    SaveFile FN

End Sub

Private Sub mnuFileSaveAs_Click()

    ComDlg.ShowSave             'Show Save As Dialog
    FN = ComDlg.FileName        'Get the Chosen File Name
    If FN = "" Then Exit Sub    'Make Sure User have selected a file
    SaveFile FN                 'Save to the selected file

End Sub

Private Sub tmrRecord_Timer()

    Record

End Sub

Private Sub tmrPlay_Timer()

    Play

End Sub
