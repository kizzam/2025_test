VERSION 5.00
Begin VB.Form frmCopy 
   Caption         =   "Copy File"
   ClientHeight    =   3195
   ClientLeft      =   375
   ClientTop       =   1905
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   480
      Top             =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(Phase 1) Please Wait......copying files"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Cnt = 0
    Screen.MousePointer = 11
    frmCopy.Refresh
    Timer1.Interval = 64764  ' Set Timer interval.
    Label2.Caption = "Coping RPMS.32z...."
    frmCopy.Refresh
    FileCopy SourcePath$ & "rpms.32z ", DestPath$ & "rpms.exe"
    DoEvents
End Sub

Private Sub Timer1_Timer()
Cnt = Cnt + 1
If Cnt > 1 Then
    Unload Me
End If
End Sub
