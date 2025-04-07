VERSION 2.00
Begin Form frmRaceMap 
   Caption         =   "Bird Race history"
   ClientHeight    =   3465
   ClientLeft      =   1725
   ClientTop       =   1890
   ClientWidth     =   6480
   Height          =   3870
   Left            =   1665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6480
   Top             =   1545
   Width           =   6600
   Begin SSPanel msgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   6495
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin CommonDialog CMDialog1 
         Left            =   5880
         Top             =   2400
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
         Left            =   5400
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
         Top             =   2520
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
      Begin SSCommand cmdQueryCT 
         Caption         =   "&Query"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin TextBox txtDateTo 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1800
         Width           =   1575
      End
      Begin TextBox txtDateFrm 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1440
         Width           =   1575
      End
      Begin SSCommand cmdAbort 
         Caption         =   "&Abort"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin Data Data4 
         Caption         =   "Data4"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\EMPTY\RPEXTRA.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "RaceTrans"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin SSCommand cmdPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin Data Data3 
         Caption         =   "Data3"
         Connect         =   ""
         DatabaseName    =   "C:\RPIGEONS\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "racetrans"
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Data Data2 
         Caption         =   "Data2"
         Connect         =   ""
         DatabaseName    =   "C:\RPIGEONS\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "race"
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   2400
         Width           =   1215
      End
      Begin Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print Race History for selected team between nominated dates (Note: Max 14 races per report)"
         Height          =   495
         Left            =   600
         TabIndex        =   7
         Top             =   840
         Width           =   4695
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date To:"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date From:"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdPrint_Click ()
'Dim strSelection As String
'Dim DteFrom As Variant
'Dim DteTo As Variant
'Dim DteFromUS As String
'Dim DteToUS As String
'Dim strCrit As String

'Dim MySet As Dynaset
'Dim SQLStmt As String

'DteFrom = CVDate(txtDateFrm.Text)
'DteTo = CVDate(txtDateTo.Text)
'DteFromUS = Format(txtDateFrm, "mm/dd/yyyy")
'DteToUS = Format(txtDateTo, "mm/dd/yyyy")

'Set DBFile = OpenDatabase(DBFullFileName$)   ' Open database.
' Set text for the SQL statement.
'strCrit$ = "SELECT DISTINCTROW racetrans.reldate, racetrans.RaceCode,"
'strCrit$ = strCrit$ & " racetrans.code_pnt, racetrans.yr, racetrans.mark, racetrans.ringno, racetrans.CodeRating"
'strCrit$ = strCrit$ & " FROM racetrans WHERE ((racetrans.reldate>#" & DteFromUS & "#))"
'strCrit$ = strCrit$ & " And racetrans.reldate<#" & DteToUS & "#))"
'strCrit$ = strCrit$ & " ORDER BY racetrans.RaceCode, racetrans.yr, racetrans.mark, racetrans.ringno"

' Create the new Dynaset.
'Set MySet = DBFile.CreateDynaset(strCrit$, 4)

'***
'Dim sCriteria As String
'Dim dFrom As Variant
'Dim dTo As Variant
'Dim dDate As String
'Dim dFromUSA As String
'Dim dToUSA As String
'Dim OrderBy As String
'Dim FlagExit As String
'Dim scrVelocity As Double
'Dim scrDistance As Double
'Dim Formula As String
'Dim Status As Integer
'Dim FlagAbort As Integer

'On Error GoTo FMErr2aa

'screen.MousePointer = 11
'msgPanel.Caption = "Please wait while History Generated ..."
'cmdAbort.Visible = True
'cmdPrint.Visible = False

'FlagExit$ = "N"
'OrderBy$ = ""
'FlagAbort = 0

'Data4.DatabaseName = DBPATH$ & "rpextra.mdb"
'Data4.RecordSource = "raceTrans"
'Data4.Refresh
'sCriteria$ = "DELETE FROM raceTrans"
'Data7.Recordset.Execute sCriteria$
'Data4.Database.Execute sCriteria$
'DoEvents
'FreeLocks

'sCriteria$ = "SELECT * FROM racetrans WHERE RaceStatus = 'R'"

    'dDate$ = strMth(GLbscMonth%) & ", 01 " & CSCombo1.Text
    'dFrom = Format(dDate$, "ddddd")
    'dTo = DateAdd("yyyy", 1, dFrom)

    'Take off 1 day
    'dFrom = DateAdd("d", -1, dFrom)

    'VB internally uses USA format for SQL queries
    'dFromUSA = "#" & Month(dFrom) & "/" & Day(dFrom) & "/" & Year(dFrom) & "#"
    'dToUSA = "#" & Month(dTo) & "/" & Day(dTo) & "/" & Year(dTo) & "#"

    'sCriteria$ = sCriteria$ & " AND RelDate > " & dFromUSA & " AND RelDate < " & dToUSA & ""

  'FlagExit$ = "N"
'Status = 0

    'msgPanel.Caption = "Please wait ... report being generated ..."
    'Report1.ReportFileName = DBPATH$ & "kpi.rpt"
    'Formula$ = "{stats.season}='" & CSCombo1.Text & "'"  'set orderby and selection of records
    'Report1.SelectionFormula = Formula$
    
    ' Show printer setup dialog box if destination is printer.
    'On Error GoTo Printererr3a
    'If Report1.Destination = 1 Then
	'CMDialog1.Copies = 1
	'CMDialog1.FromPage = 1
	'CMDialog1.ToPage = 1
	'CMDialog1.CancelError = True
	'CMDialog1.Action = 5
	'DoEvents
	'On Error GoTo 0
    'End If

    'If CMDialog1.PrinterDefault Then
	'Report1.CopiesToPrinter = CMDialog1.Copies
    'Else
	'MsgBox "The printer you use MUST be selected as the default printer.", 16, "Printer Error"
	'Exit Sub
    'End If
    
    'ShowWhat
    'Report1.Action = 1
    'DoEvents 'allow report to print
    'GoTo shutdown

'FMErr2aa:
    'Select Case Err
    'Case 3021
	'MsgBox "Database Empty", 16, "INFORMATION ONLY"
	'Status = -1
	'Resume Next
    'Case 3022
	'MsgBox "Record already exists", 16, "Add record Error"
	'Resume Next
    'Case 3026
	'MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
	'cmdAbort.Visible = False
	'cmdPrint.Visible = True
	'msgPanel.Caption = ""
	'screen.MousePointer = 0
	'Exit Sub
    'Case 3043
	'MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
	'GoTo shutdown
    'Case Else
	'MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
	'cmdAbort.Visible = False
	'cmdPrint.Visible = True
	'msgPanel.Caption = ""
	'screen.MousePointer = 0
	'Exit Sub
    'End Select
    
    'Resume Next

'Printererr3a:
    'MsgBox Error$, 16, "Report Printing Error"
    'Exit Sub

'shutdown:
    'cmdAbort.Visible = False
    'cmdPrint.Visible = True
    'msgPanel.Caption = ""
    'screen.MousePointer = 0

End Sub

Sub cmdQueryCT_Click ()
'Dim SqlSment As String
'Dim MySet As Dynaset

'SqlSment$ = "TRANSFORM Last([2001RacesTrans].CodeRating) AS LastOfCodeRating"
'SqlSment$ = SqlSment$ & " SELECT [2001RacesTrans].yr, [2001RacesTrans].mark, [2001RacesTrans].ringno"
'SqlSment$ = SqlSment$ & " FROM 2001RacesTrans"
'SqlSment$ = SqlSment$ & " GROUP BY [2001RacesTrans].yr, [2001RacesTrans].mark, [2001RacesTrans].ringno"
'SqlSment$ = SqlSment$ & " PIVOT Format([reldate],'Short Date')"
'Set MySet = DBFile.CreateDynaset(SqlSment$, 4)

End Sub

Sub Form_Load ()
    Data2.DatabaseName = DBFullFileName$
    Data2.RecordSource = "race"
    Data2.Refresh
    Data3.DatabaseName = DBFullFileName$
    Data3.RecordSource = "racetrans"
    Data3.Refresh
End Sub

Sub txtDateFrm_LostFocus ()
    'if year is entered as 00 then assume 2000 otherwise accept
    'still have slight problem with 29/02/00 assumes 1900 for isdate function & reports as invalid
    'didn't worry about this

    If Trim(txtDateFrm.Text) = "" Then
	txtDateFrm.Text = "  /  /  "
	MsgBox "WARNING - Date can't be empty", 0, "Date From required"
	txtDateFrm.SetFocus
	txtDateFrm.SelStart = 0
	txtDateFrm.SelLength = Len(txtDateFrm.Text)
	Exit Sub
    End If
    
    'Fix potential leap year problem otherwise won't pass isdate() check
    If txtDateFrm.Text = "29/2/00" Or txtDateFrm.Text = "29/2/00" Then
	txtDateFrm.Text = DateAdd("yyyy", 100, txtDateFrm.Text)
    End If
    
    If IsDate(txtDateFrm.Text) Then 'Must be Valid
	txtDateFrm.Text = DateFix(txtDateFrm.Text)
    Else
	MsgBox "WARNING - From Date ", 0, "Date FROM format incorrect"
	txtDateFrm.SetFocus
	txtDateFrm.SelStart = 0
	txtDateFrm.SelLength = Len(txtDateFrm.Text)
    End If
End Sub

Sub txtDateTo_LostFocus ()
    'if year is entered as 00 then assume 2000 otherwise accept
    'still have slight problem with 29/02/00 assumes 1900 for isdate function & reports as invalid
    'didn't worry about this

    If Trim(txtDateTo.Text) = "" Then
	txtDateTo.Text = "  /  /  "
	MsgBox "WARNING - Date can't be empty", 0, "Date TO required"
	txtDateTo.SetFocus
	txtDateTo.SelStart = 0
	txtDateTo.SelLength = Len(txtDateTo.Text)
	Exit Sub
    End If
    
    'Fix potential leap year problem otherwise won't pass isdate() check
    If txtDateTo.Text = "29/2/00" Or txtDateTo.Text = "29/2/00" Then
	txtDateTo.Text = DateAdd("yyyy", 100, txtDateTo.Text)
    End If
    
    If IsDate(txtDateTo.Text) Then 'Must be Valid
	txtDateTo.Text = DateFix(txtDateTo.Text)
    Else
	MsgBox "WARNING - From Date ", 0, "Date TO format incorrect"
	txtDateTo.SetFocus
	txtDateTo.SelStart = 0
	txtDateTo.SelLength = Len(txtDateTo.Text)
    End If
End Sub

