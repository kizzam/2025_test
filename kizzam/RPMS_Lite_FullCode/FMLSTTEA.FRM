VERSION 2.00
Begin Form FMlstTeam 
   Caption         =   "Master Team List - FMLSTTEAM"
   ClientHeight    =   6645
   ClientLeft      =   1575
   ClientTop       =   2175
   ClientWidth     =   9330
   Height          =   7155
   Left            =   1515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9330
   Top             =   1725
   Width           =   9450
   Begin SSPanel MsgPanel3 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   6120
      Width           =   2175
   End
   Begin SSPanel MsgPanel1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   6120
      Width           =   7215
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   6135
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      Begin SSCommand cmdTeam 
         Caption         =   "&Add Team"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   5280
         Width           =   1215
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   0   'False
         BoundReportHeading=   ""
         Connect         =   ""
         CopiesToPrinter =   1
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   0   'False
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   600
         MarginBottom    =   0
         MarginLeft      =   0
         MarginRight     =   0
         MarginTop       =   0
         PrinterCollation=   2  'default
         PrinterCopies   =   1
         PrinterDriver   =   ""
         PrinterName     =   ""
         PrinterPort     =   ""
         PrinterStartPage=   0
         PrinterStopPage =   0
         PrintFileCharSepQuote=   ""
         PrintFileCharSepSeparator=   ""
         PrintFileName   =   ""
         PrintFileType   =   2  'Text
         PrintFileUseRptDateFmt=   0   'False
         PrintFileUseRptNumberFmt=   0   'False
         ReportFileName  =   ""
         ReportSource    =   0  'Report File
         ReportTitle     =   ""
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   120
         UserName        =   ""
         WindowBorderStyle=   2  'Sizable
         WindowControlBox=   -1  'True
         WindowControls  =   -1  'True
         WindowHeight    =   300
         WindowLeft      =   100
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowParentHandle=   0
         WindowState     =   0  'normal
         WindowTitle     =   ""
         WindowTop       =   100
         WindowWidth     =   480
      End
      Begin CommonDialog CMDialog1 
         Left            =   120
         Top             =   120
      End
      Begin SSCommand CmdPrint 
         Caption         =   "&Print Team"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4680
         TabIndex        =   6
         Top             =   5280
         Width           =   1335
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Notes on teams"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   4455
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   3
         Top             =   720
         Width           =   3735
         Begin Label Label5 
            BackStyle       =   0  'Transparent
            Height          =   4095
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   3495
         End
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Show only birds that match the following"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   1455
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   16
         Top             =   720
         Width           =   3735
         Begin SSCheck ShowAll 
            Caption         =   "ALL on File"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   25
            Top             =   1200
            Value           =   -1  'True
            Width           =   1695
         End
         Begin CSComboBox CSCombo3 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   "	"
            ColWidth        =   FMLSTTEA.FRX:0000
            Contents        =   FMLSTTEA.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   240
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   24
            Text            =   "ALL"
            TextDisplay     =   0  'first visible
            Top             =   960
            ValueCol        =   -1
            Width           =   1455
         End
         Begin SSCheck showRef 
            Caption         =   "Reference Birds"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1920
            TabIndex        =   23
            Top             =   960
            Width           =   1695
         End
         Begin SSCheck showStock 
            Caption         =   "Stock Birds"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   22
            Top             =   720
            Width           =   1335
         End
         Begin SSCheck showRace 
            Caption         =   "Race Birds"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1920
            TabIndex        =   21
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin CSComboBox CSCombo2 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   "	"
            ColWidth        =   FMLSTTEA.FRX:0061
            Contents        =   FMLSTTEA.FRX:00A1
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   240
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   18
            TextDisplay     =   0  'first visible
            Top             =   480
            ValueCol        =   -1
            Width           =   1455
         End
         Begin SSCheck showCurrent 
            Caption         =   "Current birds"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1920
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ring Year:"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
      End
      Begin SSCommand CmdDelete 
         Caption         =   "<<"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4200
         TabIndex        =   2
         Top             =   3240
         Width           =   615
      End
      Begin SSCommand CmdAdd 
         Caption         =   ">>"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4200
         TabIndex        =   4
         Top             =   2040
         Width           =   615
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ","
         ColWidth        =   FMLSTTEA.FRX:01DE
         Contents        =   FMLSTTEA.FRX:021E
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   405
         Left            =   2520
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "description,code,status"
         ListDataSource  =   "Data1"
         SearchString    =   ""
         StartSearch     =   0
         TabIndex        =   0
         Text            =   "CSCombo1"
         TextDisplay     =   0  'first visible
         Top             =   240
         ValueCol        =   1
         Width           =   4335
      End
      Begin Data Data3 
         Caption         =   "Data3"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7800
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueGrid Table2 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data3"
         Editable        =   0   'False
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   2775
         HorzLines       =   0  'None
         Layout          =   FMLSTTEA.FRX:0220
         LayoutIndex     =   1
         Left            =   360
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   12
         Top             =   2400
         UseBookmarks    =   -1  'True
         Width           =   3735
         WrapCellPointer =   0   'False
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data2"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   4095
         HorzLines       =   0  'None
         Layout          =   FMLSTTEA.FRX:0D61
         LayoutIndex     =   1
         Left            =   4920
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   1
         Top             =   1080
         UseBookmarks    =   -1  'True
         Width           =   4335
         WrapCellPointer =   0   'False
      End
      Begin Data Data2 
         Caption         =   "Data2"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7800
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "teambird"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7800
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "teams"
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin SSCommand CmdUpdate 
         Caption         =   "&Update Team"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   3000
         TabIndex        =   5
         Top             =   5280
         Width           =   1335
      End
      Begin SSCommand CmdExit 
         Caption         =   "&Exit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6360
         TabIndex        =   8
         Top             =   5280
         Width           =   1335
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MASTER BIRD LIST"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   3735
      End
      Begin Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BIRDS IN TEAM"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   840
         Width           =   4335
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select team:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
   End
End
Option Explicit
Dim FGupdate
Dim nCntTeam As Long
Dim nCntMaster As Long

Sub CmdAdd_Click ()
    On Error GoTo FMErr:
    Dim ii As Integer
    Dim NL As String
    NL$ = Chr$(10)

    ' Process each selected row.
    DoEvents
    If Table2.BookmarkCount < 1 Then
        MsgBox "Must select BIRD to add to TEAM first - Select from Master bird list on left", 0, "Move to left until you get a tick & click to select record"
    Else
        For ii% = 0 To Table2.BookmarkCount - 1
            Data3.Recordset.Bookmark = Table2.BookmarkList(ii%)
            TeamMbrAdd
        Next ii%
    End If

    Table2.BookmarkCount = 0
    Data2.Refresh
    Exit Sub

FMErr:
    Select Case Err
    Case 3022
        MsgBox "Bird is ALREADY a MEMBER of this TEAM" & NL$ & Format$(Data3.Recordset("yr"), "") & "-" & Data3.Recordset("mark") & "-" & Data3.Recordset("ringno"), 16, "No Action for this Bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next

End Sub

Sub CmdAdd_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel1.Caption = "Add Selected bird (left pane) to team"
End Sub

Sub CmdDelete_Click ()

    On Error GoTo FMErr2:
    Dim NL As String
    Dim ii As Integer
    Dim reply As Integer

    NL$ = Chr(13)
    ' Process each selected row.
        If table1.BookmarkCount > 0 Then
            For ii% = 0 To table1.BookmarkCount - 1
                Data2.Recordset.Bookmark = table1.BookmarkList(ii%)
                reply% = MsgBox("Do you remove this BIRD from the Team" & NL$ & Data2.Recordset("yr") & "/" & Data2.Recordset("mark") & "/" & Data2.Recordset("ringno") & ".", 4, "CONFIRM DELETION")
                If reply% = 6 Then
                    Data2.Recordset.Delete
                    nCntTeam& = nCntTeam& - 1
                    Label2.Caption = Trim(Str(nCntTeam&)) & " BIRDS IN TEAM"
                End If
            Next ii%
        ElseIf nCntTeam& < 1 Then
            MsgBox "Team empty" & NL$ & "NO birds available to remove/delete from team" & NL$ & "see list on right", 0, "Cannot Delete"
        Else
            MsgBox "Must select bird record to REMOVE from team - Select from Team Bird List (in right panel)", 0, "Move to left && click to select bird record"
        End If

    table1.BookmarkCount = 0
    Data2.Refresh
    Exit Sub
FMErr2:
    Select Case Err
    Case 3021  'No Current Record
        Resume Next
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(data1.Recordset("yr"), "") & "-" & data1.Recordset("mark") & "-" & data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next

End Sub

Sub CmdDelete_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel1.Caption = "Remove Selected bird (right Pane) from team"
End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdPrint_Click ()
    Dim Formula As String
    
    If CSCombo1.Text = "" Then
        MsgBox "MUST Select TEAM to print", 0, "NOTHING TO PRINT"
        If CSCombo1.Enabled = True Then
            CSCombo1.SetFocus
        Else
            cmdTeam.SetFocus
        End If
        Exit Sub
    End If

    'If indLayoutL Then
    '    Report1.ReportFileName = DBPATH$ & "teamL.rpt"
    'ElseIf indLayoutP Then
    '    Report1.ReportFileName = DBPATH$ & "teamP.rpt"
    'ElseIf indLayoutL2 Then
    '    Report1.ReportFileName = DBPATH$ & "teamL2.rpt"
    'ElseIf indLayoutL3 Then
        Report1.ReportFileName = DBPATH$ & "teamL3.rpt"
    'Else
    '    Report1.ReportFileName = DBPATH$ & "teamP.rpt"
    'End If
    OrderBy     'Set Order of records in Report
    Formula$ = PrtShowWhat() 'set orderby and selection of records
    Report1.SelectionFormula = Formula$
    
' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr3
        CMDialog1.Copies = 1
        CMDialog1.FromPage = 1
        CMDialog1.ToPage = 1
        CMDialog1.CancelError = True
        CMDialog1.Action = 5
        DoEvents
        On Error GoTo 0
    End If

    If CMDialog1.PrinterDefault Then
        Report1.CopiesToPrinter = CMDialog1.Copies
    Else
        MsgBox "The printer you use MUST be selected as the default printer.", 16, "Printer Error"
        Exit Sub
    End If
    
        'If Trim(CSFrom.Text) <> "" Then
        'Formula$ = "{employee.EmployeeID} in '" & CSFrom.Text & "' to '" & CSTo.Text & "'"
        'Report1.SelectionFormula = Formula$
    'End If
    
    'ShowWhat
    Report1.Action = 1
    DoEvents 'allow report to print
    Exit Sub

PrinterErr3:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

Sub cmdTeam_Click ()
    On Error GoTo err_CT1
    Dim cTeam As String
    cTeam = CSCombo1.Value
    FMteam.Show 1
    data1.Refresh
    FreeLocks
    'table1.Refresh
    'DoEvents
    data1.Recordset.MoveLast
    If data1.Recordset.RecordCount < 1 Then
        CSCombo1.Enabled = False
        cmdUpdate.Enabled = False
    Else
        CSCombo1.Enabled = True
        cmdUpdate.Enabled = True
    End If
    data1.Recordset.MoveFirst
    CSCombo1.Value = cTeam
    Exit Sub

err_CT1:
    Select Case Err
    Case 3021
        Resume Next
    End Select
End Sub

Sub CmdUpdate_click ()

If FGupdate = False Then
    FGupdate = True
    Frame3D2.Visible = True
    Frame3D1.Visible = False
    Label3.Visible = False
    Table2.Visible = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    Table2.Enabled = False
    cmdAdd.Visible = False
    cmdDelete.Visible = False
    table1.SelectMode = 0
Else
    FGupdate = False
    Frame3D2.Visible = False
    Frame3D1.Visible = True
    Label3.Visible = True
    Table2.Visible = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    Table2.Enabled = True
    cmdAdd.Visible = True
    cmdDelete.Visible = True
    table1.SelectMode = 1
End If

End Sub

Sub CSCombo1_Change ()
Dim strSelection As String
Dim NL As String

On Error GoTo FMLstTeamERR:

NL$ = Chr(13)

strSelection$ = "SELECT * FROM teambird WHERE code = '" & CSCombo1.Value & "' ORDER BY val(ringno), ringno, mark, yr"
Data2.RecordSource = strSelection$
Data2.Refresh
DoEvents
Data2.Recordset.MoveLast
nCntTeam& = Data2.Recordset.RecordCount
Label2.Caption = Trim(Str(nCntTeam&)) & " BIRDS IN TEAM"
Data2.Recordset.MoveFirst
DoEvents
'Table2.Refresh
Exit Sub

FMLstTeamERR:
    Select Case Err
    Case 3021
        'MsgBox "No Current record" & NL$ & "so there!", 16, "ERROR"
        Resume Next
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(data1.Recordset("yr"), "") & "-" & data1.Recordset("mark") & "-" & data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next
End Sub

Sub CSCombo2_Change ()
setselections
End Sub

Sub CSCombo3_Change ()
setselections
End Sub

Sub Form_Load ()
Dim cnt1 As Integer

On Error GoTo FMerr3a
    
    Label5.Caption = "This screen allows TEAMs to be maintained." & Chr(10)
    Label5.Caption = Label5.Caption & "To maintain a team (i.e., add or remove) birds from a team, FIRST select the team from the 'Select Team' dropdown list above." & Chr(10)
    Label5.Caption = Label5.Caption & "Once team selected the birds that make up this team are listed in the right pane." & Chr(10)
    Label5.Caption = Label5.Caption & "To CHANGE this Team (ie add or remove birds from team), click the 'UPDATE TEAM' button below, this displays the master list of birds in the left pane." & Chr(10)
    Label5.Caption = Label5.Caption & "TO ADD bird to team, Select bird from Master list in left panel then click '>>' button to add to team list in the right panel." & Chr(10)
    Label5.Caption = Label5.Caption & "TO REMOVE bird from team, Select bird from Team list in right panel then click '<<' button to remove bird from team." & Chr(10)
    screen.MousePointer = 11
    data1.DatabaseName = DBFullFileName$
    data1.RecordSource = "SELECT * FROM teams ORDER BY description DESC"
    data1.Refresh
    data1.Recordset.MoveLast
    If data1.Recordset.RecordCount < 1 Then
        CSCombo1.Enabled = False
        cmdUpdate.Enabled = False
    Else
        CSCombo1.Enabled = True
        cmdUpdate.Enabled = True
    End If
    data1.Recordset.MoveFirst
    
    Data2.DatabaseName = DBFullFileName$
    Data2.RecordSource = "teambird"
    
    Data3.DatabaseName = DBFullFileName$

'setup screen
    'Table2.Enabled = False
    FGupdate = False
    CmdUpdate_click

    showcurrent.Value = True
    showrace.Value = True
    showAll.Value = False
    CSCombo2.Value = "ALL"
    CSCombo2.Text = "ALL"
    CSCombo3.Value = "ALL"
    CSCombo3.Text = "ALL"

    CSCombo1.Value = data1.Recordset("code")
    CSCombo1.Refresh
    CSCombo1_Change
    setselections
    screen.MousePointer = 0
Exit Sub

FMerr3a:
    Select Case Err
    Case 3021
        'MsgBox "Team Code Database Empty - Go to File/Team Codes - Create Team Code", 16, "CANNOT USE UNTIL TEAM  CREATED"
        'Unload Me
        'Exit Sub
        Resume Next
    Case 3022
        MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
    End Select
    
    Resume Next

End Sub

Sub OrderBy ()
    Dim cnt1 As Integer
    cnt1% = -1
    Report1.SortFields(0) = ""
    cnt1% = cnt1 + 1
    Report1.SortFields(cnt1%) = "+{@Intringno}"
    cnt1% = cnt1% + 1
    Report1.SortFields(cnt1%) = "+{teambird.yr}"
    cnt1% = cnt1% + 1
    Report1.SortFields(cnt1%) = "+{teambird.mark}"
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel1.Caption = ""
End Sub

Function PrtShowWhat () As String
    Dim s As String
    Dim mteam As String

    mteam$ = CSCombo1.Value

    s$ = "{teambird.code} = '" & mteam$ & "'"

    PrtShowWhat$ = s$

End Function

Sub setselections ()
    Dim s As String
    Dim o As String

    On Error GoTo ByErr9

    s$ = ShowWhat()
    o$ = "yr DESC, val(ringno), ringno, mark"

    If s$ = "" Then
        Data3.RecordSource = "SELECT * FROM master ORDER BY " & o$
    Else
        Data3.RecordSource = "SELECT * FROM master WHERE " & s$ & " ORDER BY " & o$
    End If
    Data3.Refresh
    Exit Sub
    
ByErr9:
    If Err = 380 Then
        Resume Next
    End If
    MsgBox "Error :" & Err & " " & Error$, 16, "Error"
    Exit Sub
End Sub

Sub ShowAll_Click (Value As Integer)
'If showAll.Value = True Then
'    showcurrent = False
'    showrace = False
'    showstock = False
'    showref = False
'End If
setselections
End Sub

Sub showCurrent_Click (Value As Integer)
showAll.Value = False
setselections
End Sub

Sub showRace_Click (Value As Integer)
showAll.Value = False
setselections
End Sub

Sub showRef_Click (Value As Integer)
showAll.Value = False
setselections
End Sub

Sub showStock_Click (Value As Integer)
showAll.Value = False
setselections
End Sub

Function ShowWhat () As String

    Dim s As String
    Dim cSex As String

    If showcurrent Then
       s$ = "indsoh = True"
    'Else
        s$ = "indsoh = false"
    End If

    If showref Then
       If s$ = "" Then
        s$ = "indref = True"
       Else
        s$ = s$ & " AND indref = True"
       End If
    End If
    If showrace Then
       If s$ = "" Then
        s$ = "indrace = True"
       Else
        s$ = s$ & " AND indrace = True"
       End If
    End If
    If showstock Then
       If s$ = "" Then
        s$ = "indstock = True"
       Else
        s$ = s$ & " AND indstock = True"
       End If
    Else
       If s$ = "" Then
        s$ = "indstock = False"
       Else
        s$ = s$ & " AND indstock = False"
       End If
    End If
    'If ShowSPb Then
     '  If s$ = "" Then
    '    s$ = "indsp2 = True"
     '  Else
    '    s$ = s$ & " AND indsp2 = True"
    '   End If
    'End If
    'If ShowSPa Then
    '   If s$ = "" Then
    '    s$ = "indsp1 = True"
     '  Else
    '    s$ = s$ & " AND indsp1 = True"
    '   End If
    'End If
    'If ShowBP2 Then
    '   If s$ = "" Then
    '    s$ = "indbpb = True"
    '   Else
     '   s$ = s$ & " AND indbpb = True"
    '   End If
    'End If
    'If ShowBP1 Then
    '   If s$ = "" Then
    '    s$ = "indbpa = True"
    '   Else
    '    s$ = s$ & " AND indbpa = True"
'       End If
'    End If

    If showAll Then
        s$ = ""
    End If

    If CSCombo3.Value = "ALL" Then
    ElseIf CSCombo3.Value = "Hens" Then
        cSex$ = "H"
    ElseIf CSCombo3.Value = "Cocks" Then
        cSex$ = "C"
    ElseIf CSCombo3.Value = "UnSexed" Then
        cSex$ = "U"
    End If

    If CSCombo3.Value <> "ALL" Then
       If s$ = "" Then
        s$ = "codesex = '" & cSex$ & "'"
       Else
        s$ = s$ & " AND codesex = '" & cSex$ & "'"
       End If
    End If
'    ElseIf ShowCocks Then
'       If s$ = "" Then
'        s$ = "codesex = 'C'"
'       Else
'        s$ = s$ & " AND codesex = 'C'"
'       End If
'    ElseIf ShowUnSexed Then
'       If s$ = "" Then
'        s$ = "codesex = 'U'"
'       Else
'        s$ = s$ & " AND codesex = 'U'"
'       End If
'    End If
    If CSCombo2.Text <> "ALL" Then
        If CSCombo2.Value = "PRE80" Then
            If s$ = "" Then
                s$ = "yr < 1980"
            Else
                s$ = s$ & " AND yr < 1980"
            End If
        Else
            If s$ = "" Then
                s$ = "yr = " & Val(CSCombo2.Value)
            Else
                s$ = s$ & " AND yr = " & Val(CSCombo2.Value)
            End If
        End If
    End If
 ShowWhat$ = s$
    
End Function

Sub TeamMbrAdd ()
On Error GoTo FMTeamMbrAddErr:
    Dim NL As String
    NL$ = Chr(13)
    Data2.Recordset.AddNew
    Data2.Recordset("Code") = CSCombo1.Value
    Data2.Recordset("yr") = Data3.Recordset("yr")
    Data2.Recordset("mark") = Data3.Recordset("mark")
    Data2.Recordset("ringno") = Data3.Recordset("ringno")
    Data2.Recordset("status") = "0"
    Data2.Recordset.Update
    nCntTeam& = nCntTeam& + 1
    Label2.Caption = Trim(Str(nCntTeam&)) & " BIRDS IN TEAM"
    Exit Sub

FMTeamMbrAddErr:
    Select Case Err
    Case 3022
        MsgBox "Bird ALREADY is a MEMBER of this TEAM" & NL$ & Format$(Data3.Recordset("yr"), "") & "-" & Data3.Recordset("mark") & "-" & Data3.Recordset("ringno"), 16, "Cannot ADD bird to team"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    Resume Next

End Sub

