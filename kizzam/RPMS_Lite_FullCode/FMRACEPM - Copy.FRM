VERSION 2.00
Begin Form FMracepmt 
   Caption         =   "Race List"
   ClientHeight    =   6345
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   9375
   Height          =   6855
   Left            =   30
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9375
   Top             =   -105
   Width           =   9495
   Begin SSPanel MsgPanel 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   5880
      Width           =   9375
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   5895
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9375
      Begin SSCommand ButtFind 
         Caption         =   "&Find"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8040
         TabIndex        =   10
         Top             =   3960
         Width           =   1215
      End
      Begin SSCommand ButtAdd 
         Caption         =   "&Add"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8040
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin SSCommand ButtProcess 
         Caption         =   "&Process"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8040
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   8040
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin SSCommand CmdPrintSum 
         Caption         =   "Print &Summary"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   5400
         Width           =   2055
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
         Left            =   7320
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
         Top             =   960
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
         Left            =   7320
         Top             =   240
      End
      Begin SSCommand CmdPrintRaces 
         Caption         =   "P&rint Details"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   5400
         Width           =   2055
      End
      Begin SSCommand CmdDelete 
         Caption         =   "&Delete"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8040
         TabIndex        =   13
         Top             =   4920
         Width           =   1215
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data1"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeadingHeight   =   1
         Height          =   3735
         HorzLines       =   0  'None
         Layout          =   FMRACEPM.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   7
         Top             =   1560
         UseBookmarks    =   -1  'True
         Width           =   7695
         WrapCellPointer =   0   'False
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   8160
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "race"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Show Only races/toss for"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   17
         Top             =   840
         Width           =   6855
         Begin CSComboBox CSCombo2 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FMRACEPM.FRX:0932
            Contents        =   FMRACEPM.FRX:0972
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   840
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            Sorted          =   -1  'True
            StartSearch     =   0
            TabIndex        =   5
            Text            =   "All Flights"
            TextDisplay     =   0  'first visible
            Top             =   240
            ValueCol        =   1
            Width           =   2055
         End
         Begin SSPanel Panel3D4 
            Alignment       =   4  'Right Justify - MIDDLE
            BackColor       =   &H00E0E0E0&
            BevelOuter      =   0  'None
            Caption         =   "Type:"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin SSPanel Panel3D2 
            Alignment       =   4  'Right Justify - MIDDLE
            BackColor       =   &H00E0E0E0&
            BevelOuter      =   0  'None
            Caption         =   "Year:"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3120
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FMRACEPM.FRX:09A8
            Contents        =   FMRACEPM.FRX:09E8
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   3840
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   6
            Text            =   "All"
            TextDisplay     =   0  'first visible
            Top             =   240
            ValueCol        =   -1
            Width           =   1575
         End
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Sort by races/toss by"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   16
         Top             =   240
         Width           =   6855
         Begin SSOption byCB 
            Caption         =   "Clock Bird"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5040
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin SSOption byPosition 
            Caption         =   "Club Position"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin SSOption byRcePnt 
            Caption         =   "Race Point"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1680
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin SSOption byDate 
            Caption         =   "Release Date"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
   End
End
Option Explicit

Sub ButtAdd_Click ()
    ' Advise the next form we are in add mode.
    FMrace.task.Text = "A"
    FMrace.Show 1
    Data1.Refresh '29/8
End Sub

Sub ButtCancel_Click ()
    Unload FMracepmt
End Sub

Sub ButtFind_Click ()
    Dim Key As String
    'bit of work to do here, composite key Yr/Mark/RingNo.
    Key$ = InputBox("Enter Race Code", "Find Race")

    Key$ = UCase(Key$)

    If Trim(Key$) <> "" Then
	Data1.Recordset.FindFirst "code_race >= '" & Key$ & "'"
    End If
End Sub

Sub ButtProcess_Click ()
    Dim ii As Integer
    ' Process each selected row.
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Race(s) selected", 0, "ERROR - Unable to Process"
    Else
	For ii% = 0 To Table1.BookmarkCount - 1
	    Data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
	    FMrace.task.Text = "M"
	    FMrace.mRaceCode.Text = Data1.Recordset("RaceCode")
	    FMrace.Key_reldate.Text = Data1.Recordset("reldate")
	    FMrace.Key_coderace.Text = Data1.Recordset("code_race")
	    FMrace.Key_codepnt.Text = Data1.Recordset("code_pnt")
	    FMrace.Mreldate.Text = Data1.Recordset("reldate")
	    FMrace.CSMcodepnt.Value = Data1.Recordset("code_pnt")
	    FMrace.Mcoderace.Text = Data1.Recordset("code_race")
	    FMrace.Show 1
	Next ii%
	Table1.BookmarkCount = 0
    End If
End Sub

Sub byCB_Click (Value As Integer)
setselections

End Sub

Sub byDate_Click (Value As Integer)
setselections
End Sub

Sub byPosition_Click (Value As Integer)
setselections

End Sub

Sub byRcePnt_Click (Value As Integer)
setselections
End Sub

Sub CmdDelete_Click ()

    Const IDYES = 6, IDNO = 7
    Dim Key As String
    Dim Key1 As String
    Dim Key2 As String
    Dim Key3 As String
    Dim Msg As String
    Dim Counter As Integer
    Dim Response As Integer
    Dim ActionQry As String
    Dim DB As Database
    Dim Tbl As Table
    Dim ii As Integer
    Dim NL As String

    NL$ = Chr(10)

    ' Process each selected row.
    If Table1.BookmarkCount < 1 Then
	MsgBox "No Race(s) selected", 0, "ERROR - Unable to Delete any"
    Else
	Response% = MsgBox("You are about to DELETE a race or races", 1, "Click OK to continue")
	If Response% = 1 Then
	    For ii% = 0 To Table1.BookmarkCount - 1
		Data1.Recordset.Bookmark = Table1.BookmarkList(ii%)
		Key$ = Data1.Recordset("RaceCode")
		Key1$ = Data1.Recordset("Code_Pnt")
		Key2$ = Data1.Recordset("Code_Race")
		Key3$ = Data1.Recordset("RelDate")
		'Ensure NOT system record, can't be changed
		'If Data1.Recordset("system") Then
		'    MsgBox "Cannot DELETE this code, It is set by the system", 16, "ERROR CANNOT DELETE"
		'    Screen.MousePointer = 0
		'    Exit Sub
		'Else
		'End If
		Msg$ = "Are you sure you want to delete this RACE " & NL$ & "Code: " & Key$ & NL$ & "From: " & Key1$ & NL$ & "Type: " & Key2$ & NL$ & "Date of Race: " & Key3$ & " ?"
		Response% = MsgBox(Msg$, 1, "Delete RACE")
		If Response% = 1 Then
		    screen.MousePointer = 11
		    'Ensure no birds in this race record, retain db integrity
		    Set DB = OpenDatabase(DBFullFileName$)    'Open Database
		    Set Tbl = DB.OpenTable("racetrans")     'Open Table
		    Tbl.Index = "RaceCode"                 'Open Index
		    Tbl.Seek "=", Key$ 'Ensure code to be deleted NOT in use
		    'Tbl.FindFirst "RaceCode = '" & Key$ & "'"
		    If Tbl.NoMatch Then
			Data1.Recordset.Delete
			'Data1.Refresh
			MsgBox "Record Deleted", 16, "Notification Only"
			Tbl.Close
			DB.Close
		    Else
			MsgBox "This RACE record has birds nominated" & NL$ & "Cannot delete until no nominations" & NL$ & "ie., " & Tbl("yr") & " " & Trim(Tbl("mark")) & " " & Trim(Tbl("ringno")), 16, "Notification CANNOT DELETE"
			Tbl.Close
			DB.Close
		    End If
		End If
	    Next ii%
	    Data1.Refresh
	End If
	Table1.BookmarkCount = 0
	screen.MousePointer = 0
	'ButtCancel.Default = True
    End If
    
End Sub

Sub CmdPrintRaces_Click ()
    Dim Formula As String
    
    If CSCombo2.Value = "T" Then 'Toss
	Report1.ReportFileName = DBPATH$ & "toss1.rpt"
    Else
	Report1.ReportFileName = DBPATH$ & "races1.rpt"
    End If
    
    OrderByPrt     'Set Order of records in Report
    Formula$ = ShowWhatPrt() 'set orderby and selection of records
    Report1.SelectionFormula = Formula$
    
    'Check the range is valid.
    'If Trim(CSFrom.Text) <> "" Or Trim(CSTo.Text) <> "" Then
     '   If Trim(CSFrom.Text) = "" Or Trim(CSTo.Text) = "" Then
      '      MsgBox "Both FROM and TO employee numbers are required.", 16, "Range Error"
       '     Exit Sub
	'End If
    'End If

    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
	On Error GoTo PrinterErr4
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

PrinterErr4:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

Sub CmdPrintSum_Click ()
frmRptFmt.Show 1
End Sub

Sub CSCombo1_Click ()
setselections

End Sub

Sub CSCombo2_Click ()
setselections

End Sub

Sub Form_Activate ()
'setselections
DoEvents
FreeLocks
DoEvents
End Sub

Sub Form_Load ()
screen.MousePointer = 11
Data1.DatabaseName = DBFullFileName$
'Data1.RecordSource = "race"
Data1.RecordSource = "SELECT * FROM race ORDER BY #reldate# DESC"
'Data1.Refresh
setselections
screen.MousePointer = 0
End Sub

Function OrderBy () As String
    If byDate Then
       OrderBy$ = "reldate"
    ElseIf ByRcePnt Then
	OrderBy$ = "code_pnt"
    ElseIf ByPosition Then
	OrderBy$ = "ClubPos"
    ElseIf ByCb Then
	OrderBy$ = "yr, ringno, mark"
    Else
       OrderBy$ = "reldate"
    End If
End Function

Sub OrderByPrt ()
    Dim Cnt1 As Integer
    Cnt1% = -1
    Report1.SortFields(0) = ""

    If byDate Then
	'OrderBy$ = "yr, mark, ring_no"
	Cnt1% = Cnt1 + 1
	Report1.SortFields(Cnt1%) = "+{race.reldate}"
    ElseIf ByRcePnt Then
	'OrderBy$ = "yr, ring_no,mark"
	Cnt1% = Cnt1 + 1
	Report1.SortFields(Cnt1%) = "+{race.code_pnt}"
    ElseIf ByPosition Then
	'OrderBy$ = "ring_no, yr, mark"
	Cnt1% = Cnt1 + 1
	Report1.SortFields(Cnt1%) = "+{race.clubpos}"
    ElseIf ByCb Then
	'OrderBy$ = "mark, yr, ring_no"
	Cnt1% = Cnt1 + 1
	Report1.SortFields(Cnt1%) = "+{race.clubvel}"
    Else
       'OrderBy$ = "yr, mark, ring_no"
	Cnt1% = Cnt1 + 1
	Report1.SortFields(Cnt1%) = "+{race.reldate}"
    End If
End Sub

Sub setselections ()
    Dim s As String
    Dim o As String
    Dim cType As String
    cType$ = ""

    On Error GoTo ByErr9

    s$ = ShowWhat(cType$)
    o$ = OrderBy()
    
    If s$ = "" Then
	Data1.RecordSource = "SELECT * FROM race ORDER BY " & o$ & " DESC"
    Else
	Data1.RecordSource = "SELECT * FROM race WHERE " & s$ & " ORDER BY " & o$ & " DESC"
    End If
    Data1.Refresh
    FreeLocks
    DoEvents
    Exit Sub
    
ByErr9:
    If Err = 380 Then
	Resume Next
    End If
    MsgBox "Error :" & Err & " " & Error$, 16, "Error"
    Exit Sub
End Sub

Function ShowWhat (cType As String) As String

    Dim s As String
    s$ = ""

    If cType$ = "R" Then 'Crystal Report Selection
    Else
	If Trim(CSCombo2.Text) = "All Flights" Then
	    s$ = ""
	Else
	    If s$ = "" Then
	       s$ = "racestatus = '" & Trim(CSCombo2.Value) & "'"
	    Else
	       s$ = s$ & " AND racestatus = '" & Trim(CSCombo2.Value) & "'"
	    End If
	End If
	If Trim(CSCombo1.Text) = "All" Then
	Else
	    If s$ = "" Then
		s$ = "Year(RelDate) = " & Val(CSCombo1.Value)
	    Else
		s$ = s$ & " AND Year(RelDate) = " & Val(CSCombo1.Value)
	    End If
	End If
    End If
    ShowWhat$ = s$

End Function

Function ShowWhatPrt () As String

    Dim s As String
    Dim mclass As String
    Dim mYear As Integer

    mclass$ = Trim(CSCombo2.Value)

    If mclass$ = "" Or mclass$ = "All" Then
	s$ = ""
    Else
	s$ = "{race.racestatus} = '" & mclass$ & "'"
    End If

    If CSCombo1.Value = "" Or CSCombo1.Value = "All" Then
    Else
	mYear% = CSCombo1.Value
	If s$ = "" Then
	    's$ = " Year({race.reldate}) = '" & mYear% & "'"
	    s$ = " Year({race.reldate}) = " & mYear%
	Else
	    's$ = s$ & " AND Year({race.reldate}) = '" & mYear% & "'"
	    s$ = s$ & " AND Year({race.reldate}) = " & mYear%
	End If
    End If

    ShowWhatPrt$ = s$

End Function

Sub Table1_UnboundFetch (Row As Long, Col As Integer, Bookmark As String, Value As String)
Dim curbm As String
    curbm$ = Data1.Recordset.Bookmark   ' Save current bookmark
    Data1.Recordset.Bookmark = Bookmark ' Position Recordset to BookMark
    ' Compute Value = "Quantity" * "Price" for the grid to display
    If Not IsNull(Data1.Recordset("yr")) Then
	If Not IsNull(Data1.Recordset("mark")) Then
	    If Not IsNull(Data1.Recordset("ringno")) Then
		Value = Trim(Str(Data1.Recordset("yr")) & "/" & Data1.Recordset("mark") & "/" & Data1.Recordset("ringno"))
	    Else
		Value = "Not Available"
	    End If
	Else
	    Value = "Not Available"
	End If
    Else
	Value = "Not Available"
    End If
    Data1.Recordset.Bookmark = curbm$   ' Restore original bookmark
End Sub

