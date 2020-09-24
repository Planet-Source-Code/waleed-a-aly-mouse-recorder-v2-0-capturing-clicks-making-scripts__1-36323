VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   ScaleHeight     =   2970
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Click Here"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   2295
      UseMnemonic     =   0   'False
      Width           =   750
   End
   Begin VB.Label lblPurpose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Record Mouse Movements & Clicks!"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   825
      TabIndex        =   3
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00000000&
      Caption         =   "To Report any Bugs and | or Suggestions, or just to say thanx Click Here to eMail Me"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   660
      TabIndex        =   2
      Top             =   2100
      UseMnemonic     =   0   'False
      Width           =   2970
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "By Waleed A. Aly"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1500
      TabIndex        =   1
      Top             =   1140
      UseMnemonic     =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mouse Recorder v2.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   1020
      TabIndex        =   0
      Top             =   540
      UseMnemonic     =   0   'False
      Width           =   2205
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FFFFFF&
      Height          =   2970
      Left            =   0
      Top             =   0
      Width           =   4350
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub lblEMail_Click()

    ShellExecute Me.hWnd, "open", "mailto:wa_aly@tdcspace.dk", vbNullString, "C:\", 5

End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Screen.MousePointer = vbArrowQuestion
    lblEMail.Font.Underline = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Screen.MousePointer = vbDefault
    lblEMail.Font.Underline = False

End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Screen.MousePointer = vbDefault
    lblEMail.Font.Underline = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Click()

    Unload Me

End Sub

Private Sub lblAppName_Click()

    Unload Me

End Sub

Private Sub lblAuthor_Click()

    Unload Me

End Sub

Private Sub lblInfo_Click()

    Unload Me

End Sub

Private Sub lblPurpose_Click()

    Unload Me

End Sub
