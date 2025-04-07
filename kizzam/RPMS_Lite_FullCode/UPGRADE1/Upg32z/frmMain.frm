VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Upgrade RPMS"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5580
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   5580
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton cmdFixRTrans 
      Caption         =   "Fix Race Trans (Y2k prob)"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   4320
   End
   Begin VB.CommandButton cmdUpgrade 
      Caption         =   "&Upgrade to V3.2z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Database:"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Windows Path:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Destination Path:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Source Path:"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Upgrade RPMS to v3.2z(c) - Phase 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Click on Upgrade to start upgrade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhelp1 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelp 
         Caption         =   "&Upgrade Steps"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCheck_Click()
    Dim db As Database
    Dim td As TableDef
    Dim Fld As Field
    Dim FldName As String
    Dim x As Integer
    AppName = "Upgrade V3.2 to V3.2z"
    Screen.MousePointer = 11
        
    'Update Table LOFT
    Set db = DBEngine.Workspaces(0).OpenDatabase(DBFullFileName$)
    Set td = db.TableDefs("Loft")
    FldName$ = "ownerfed"
    For x = 1 To 3
        ' Delete the newly added fields.
        FldName$ = "OwnerFed"
        FldName$ = InputBox("Field", "Input field", 1)

        If FieldExists(td, FldName$) Then
            'MsgBox "Field Exists"
            'AppendDeleteField Td, "DELETE", FldName$
            'AppendDeleteField Td, "APPEND", FldName$, dbText, 39
        Else
            'MsgBox "Field Doesn't Exists"
            'AppendDeleteField Td, "APPEND", FldName$, dbText, 39
        End If
        Label1.Caption = "Upgrade Completed"
        cmdUpgrade.Enabled = False
        Screen.MousePointer = 0
    Next x
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFixRTrans_Click()
    Fix_relDate
End Sub

Private Sub cmdUpgrade_Click()
    Dim db As Database
    Dim td As TableDef
    Dim Fld As Field
    Dim FldName As String
    Dim x As Integer
    Dim FldValue As String
    Dim CmdLine As String
    Dim result As Integer
        
    cmdUpgrade.Enabled = False
    
    result = MsgBox("Have you completed two(2) separate backups on separate diskettes?", vbYesNo, "MUST do BACKUP prior to this operation!")
    If result <> 6 Then
        Label2.Caption = "Upgrade ABORTED - MUST do backup first"
        MsgBox "See Manual re Backing up RPMS, after backup complete you can attempt upgrade!", vbCritical, "BACKUP MUST BE RUN PRIOR THIS UPGRADE"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Label2.Caption = "PHASE 1 Copying Files - Please Wait..."
    frmCopy.Show 1
    Screen.MousePointer = 0
        
    Label2.Caption = "PHASE 1A Copying Files - Please Wait..."
    frmCopy1.Show 1
    Screen.MousePointer = 0
    
    Label2.Caption = "PHASE 2 Modifying Database - Please Wait..."
    Screen.MousePointer = 0
    frmDBupg.Show 1
    DoEvents
    
    Label2.Caption = "PHASE 3 Creating Batch file - Please Wait..."
    CreateFile2
    
    Screen.MousePointer = 11
    Label2.Caption = "PHASE 4 - Install upgraded files" & Chr(10)
    RetVal = Shell(DestPath$ & "upg.bat", 1)
    
    Label2.Caption = "INSTALLATION COMPLETED" & Chr(10)
    Screen.MousePointer = 0

Exit Sub

ErrorSetup:
    MsgBox AppName + " is not properly installed.  Please re-run setup at a later time to install the " & AppName & " properly.", 48, dialogCaption$
    ChDrive Mid$(WinSysDir$, 1, 2) ' Set back to hard disk
    ChDir Right$(WinSysDir$, Len(WinSysDir$) - 2)
    End
    Exit Sub
End Sub

Private Sub Label4_Click()
Frame1.Visible = False
End Sub

Private Sub Form_Load()
Text1.Text = WinSysDir$
Text2.Text = SourcePath$
Text3.Text = DestPath$
Text4.Text = DBFullFileName$
End Sub

Private Sub mnuExit_Click()
cmdExit_Click
End Sub

Private Sub mnuhelp_Click()
Frame1.Visible = True
End Sub
