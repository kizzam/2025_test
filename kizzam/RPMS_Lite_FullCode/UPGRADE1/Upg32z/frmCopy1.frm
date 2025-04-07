VERSION 5.00
Begin VB.Form frmCopy1 
   Caption         =   "Copy Files Phase 1A"
   ClientHeight    =   3195
   ClientLeft      =   675
   ClientTop       =   1290
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   480
      Top             =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(Phase 1A) Please Wait......copying files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmCopy1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim MyFile, MyPath, MyName
' Returns "WIN.INI" if it exists.
MyFile = Dir("C:\WINDOWS\WIN.INI")

    Cnt = 0
    Screen.MousePointer = 11
    Timer1.Interval = 64764  ' Set Timer interval.
    Screen.MousePointer = 11
    Label2.Caption = "Coping DB System 2...."
    'MyFile = Dir(WinSysDir$ & "\system\dao350.dll")
    'If UCase(MyFile) <> UCase("dao350.dll") Then
        'FileCopy SourcePath$ & "dao350.dll", WinSysDir$ & "\system\dao350.dll"
    'End If
    DoEvents
End Sub

Private Sub Timer1_Timer()
Cnt = Cnt + 1
If Cnt > 1 Then
    Unload Me
End If
End Sub
