VERSION 2.00
Begin Form frmRptCard 
   Caption         =   "Report Card"
   ClientHeight    =   4305
   ClientLeft      =   1380
   ClientTop       =   2340
   ClientWidth     =   7290
   Height          =   4815
   Left            =   1320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   7290
   Top             =   1890
   Width           =   7410
   Begin SSPanel MsgPanel 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   7335
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin Data Data5 
         Caption         =   "Data5"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "BloodLine"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin CommonDialog CMDialog1 
         Left            =   1080
         Top             =   3480
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   0   'False
         BoundReportHeading=   ""
         Connect         =   ""
         CopiesToPrinter =   1
         DataSource      =   "Data7"
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   0   'False
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   1680
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
         Top             =   3480
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
      Begin Data Data7 
         Caption         =   "Data7"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RPEXTRA.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Stats"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin SSCommand cmdPrint2 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         Height          =   855
         Left            =   2160
         TabIndex        =   7
         Top             =   2880
         Width           =   1335
      End
      Begin SSCommand cmdAbort 
         Caption         =   "&Abort"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   2160
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin Data Data6 
         Caption         =   "Data6"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Ratings"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Data Data4 
         Caption         =   "Data4"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "colour"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1860
      End
      Begin Data Data3 
         Caption         =   "Data3"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "race"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Data Data2 
         Caption         =   "Data2"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "racetrans"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   3840
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FRMRPTCA.FRX:0000
         Contents        =   FRMRPTCA.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   405
         Left            =   2880
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         TabIndex        =   1
         Text            =   "ALL"
         TextDisplay     =   0  'first visible
         Top             =   720
         ValueCol        =   -1
         Width           =   1695
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Season is defined as 12 months from month nominated in loft setup eg from main menu check File/Owner."
         Height          =   1335
         Left            =   4920
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ensure Printer is connected and on line ready to print. This report prints direct to printer."
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1095
         Left            =   2040
         TabIndex        =   3
         Top             =   1440
         Width           =   2655
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For Season:"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
   End
End
Option Explicit
Dim FlagAbort As Integer

Sub cmdAbort_Click ()
FlagAbort% = -1
End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdPrint2_Click ()

Dim sCriteria As String
Dim dFrom As Variant
Dim dTo As Variant
Dim dDate As String
Dim dFromUSA As String
Dim dToUSA As String
Dim OrderBy As String
Dim FlagExit As String
Dim scrVelocity As Double
Dim scrDistance As Double
Dim Formula As String
Dim Status As Integer

On Error GoTo FMErr2aa

screen.MousePointer = 11
msgPanel.Caption = "Please wait while Statistics Generated ..."
cmdAbort.Visible = True
cmdPrint2.Visible = False

FlagExit$ = "N"
OrderBy$ = ""
FlagAbort = 0

Data7.DatabaseName = DBPATH$ & "rpextra.mdb"
Data7.RecordSource = "stats"
Data7.Refresh
sCriteria$ = "DELETE FROM stats WHERE Season = '" & CSCombo1.Text & "'"
'Data7.Recordset.Execute sCriteria$
Data7.Database.Execute sCriteria$
DoEvents
FreeLocks

'sCriteria$ = "SELECT * FROM stats WHERE Season = '" & CSCombo1.Text & "';"
'Data7.RecordSource = sCriteria$
'Data7.Refresh
'While Not Data7.Recordset.EOF
'    Data7.Recordset.Delete
'    Data7.Recordset.MoveNext
'Wend

sCriteria$ = "SELECT * FROM racetrans WHERE RaceStatus = 'R'"

If CSCombo1.Text = "ALL" Then
    sCriteria$ = sCriteria$ & ""
Else
    dDate$ = strMth(GLbscMonth%) & ", 01 " & CSCombo1.Text
    dFrom = Format(dDate$, "ddddd")
    dTo = DateAdd("yyyy", 1, dFrom)

    'Take off 1 day
    dFrom = DateAdd("d", -1, dFrom)

    'VB internally uses USA format for SQL queries
    dFromUSA = "#" & Month(dFrom) & "/" & Day(dFrom) & "/" & Year(dFrom) & "#"
    dToUSA = "#" & Month(dTo) & "/" & Day(dTo) & "/" & Year(dTo) & "#"

    sCriteria$ = sCriteria$ & " AND RelDate > " & dFromUSA & " AND RelDate < " & dToUSA & ""
End If

'Stats
'Get fastest velocity for loft
Status = 0
msgPanel.Caption = "Please wait ... selecting fastest velocity ..."
OrderBy$ = " ORDER BY velocity DESC"
Data2.RecordSource = sCriteria$ & OrderBy$
Data2.Refresh
Data2.Recordset.MoveFirst
If Status = 0 Then
    While Not Data2.Recordset.EOF And FlagExit = "N" And Not Data2.Recordset("velocity") < 1
        DoEvents
        If FlagAbort Then
            GoTo shutdown
        End If
        scrVelocity = Data2.Recordset("velocity")
        If scrVelocity > 0 Then
            Data7.Recordset.AddNew
            Data7.Recordset("Season") = CSCombo1.Text
            Data7.Recordset("InfoReported") = "Fastest Velocity"
            Data7.Recordset("StatType") = "STAT"
            Data7.Recordset("yr") = Data2.Recordset("yr")
            Data7.Recordset("mark") = Data2.Recordset("mark")
            Data7.Recordset("ringno") = Data2.Recordset("ringno")
            Data7.Recordset("code_pnt") = Data2.Recordset("code_pnt")
            Data7.Recordset("velocity") = Data2.Recordset("velocity")
            Data7.Recordset("RaceCode") = Data2.Recordset("RaceCode")
            Data7.Recordset("code_race") = Data2.Recordset("code_race")
            Data7.Recordset("posLoft") = Data2.Recordset("posLoft")
            Data7.Recordset("posClub") = Data2.Recordset("posClub")
            Data7.Recordset("posFed") = Data2.Recordset("posFed")
            Data7.Recordset("posSec") = Data2.Recordset("posSec")
            Data7.Recordset("posClub") = Data2.Recordset("posClub")
            Data7.Recordset("CodeRating") = Data2.Recordset("CodeRating")
            Data7.Recordset("Distance") = Data2.Recordset("Distance")
            Data7.Recordset.Update
        End If
        Data2.Recordset.MoveNext
        If Data2.Recordset("Velocity") < scrVelocity Then
            FlagExit$ = "Y"
        End If
    Wend
    Status = 0
    FlagExit$ = "N"
End If

'Get fastest velocity for longest distance for loft
msgPanel.Caption = "Please wait ... selecting longest distance ..."
OrderBy$ = " ORDER BY distance DESC, Velocity DESC"
Data2.RecordSource = sCriteria$ & OrderBy$
Data2.Refresh
Data2.Recordset.MoveFirst
If Status = 0 Then
  While Not Data2.Recordset.EOF And Not FlagExit = "Y" And Not Data2.Recordset("distance") < 1
    DoEvents
    If FlagAbort Then
        GoTo shutdown
    End If
    scrDistance = Data2.Recordset("Distance")
    Data7.Recordset.AddNew
    Data7.Recordset("Season") = CSCombo1.Text
    Data7.Recordset("InfoReported") = "Longest Distance"
    Data7.Recordset("StatType") = "STAT"
    Data7.Recordset("yr") = Data2.Recordset("yr")
    Data7.Recordset("mark") = Data2.Recordset("mark")
    Data7.Recordset("ringno") = Data2.Recordset("ringno")
    Data7.Recordset("code_pnt") = Data2.Recordset("code_pnt")
    Data7.Recordset("velocity") = Data2.Recordset("velocity")
    Data7.Recordset("RaceCode") = Data2.Recordset("RaceCode")
    Data7.Recordset("code_race") = Data2.Recordset("code_race")
    Data7.Recordset("posLoft") = Data2.Recordset("posLoft")
    Data7.Recordset("posClub") = Data2.Recordset("posClub")
    Data7.Recordset("posFed") = Data2.Recordset("posFed")
    Data7.Recordset("posSec") = Data2.Recordset("posSec")
    Data7.Recordset("CodeRating") = Data2.Recordset("CodeRating")
    Data7.Recordset("Distance") = Data2.Recordset("Distance")
    Data7.Recordset.Update
    Data2.Recordset.MoveNext
    If Data2.Recordset("Distance") < scrDistance Then
        FlagExit = "Y"
    End If
  Wend
  Status = 0
  FlagExit$ = "N"
End If

'Get Pos=1 for Section
msgPanel.Caption = "Please wait ... selecting Section results ..."
OrderBy$ = " AND PosSec < 11 AND PosSec > 0 ORDER BY PosSec ASC, yr, mark, ringno"
Data2.RecordSource = sCriteria$ & OrderBy$
Data2.Refresh
Data2.Recordset.MoveFirst
If Status = 0 Then
  While Not Data2.Recordset.EOF And Not FlagExit = "Y"
    DoEvents
    If FlagAbort Then
        GoTo shutdown
    End If
    Data7.Recordset.AddNew
    Data7.Recordset("Season") = CSCombo1.Text
    Data7.Recordset("InfoReported") = "Section Top 10"
    Data7.Recordset("StatType") = "KPI"
    Data7.Recordset("yr") = Data2.Recordset("yr")
    Data7.Recordset("mark") = Data2.Recordset("mark")
    Data7.Recordset("ringno") = Data2.Recordset("ringno")
    Data7.Recordset("code_pnt") = Data2.Recordset("code_pnt")
    Data7.Recordset("velocity") = Data2.Recordset("velocity")
    Data7.Recordset("RaceCode") = Data2.Recordset("RaceCode")
    Data7.Recordset("code_race") = Data2.Recordset("code_race")
    Data7.Recordset("posLoft") = Data2.Recordset("posLoft")
    Data7.Recordset("posClub") = Data2.Recordset("posClub")
    Data7.Recordset("posFed") = Data2.Recordset("posFed")
    Data7.Recordset("posSec") = Data2.Recordset("posSec")
    Data7.Recordset("CodeRating") = Data2.Recordset("CodeRating")
    Data7.Recordset("Distance") = Data2.Recordset("Distance")
    Data7.Recordset.Update
    Data2.Recordset.MoveNext
  Wend
  FlagExit$ = "N"
End If
Status = 0

'Get Pos < 11 for fed
msgPanel.Caption = "Please wait ... selecting Federation results ..."
OrderBy$ = " AND PosFed < 11 AND PosFed > 0 ORDER BY PosFed ASC, yr, mark, ringno"
Data2.RecordSource = sCriteria$ & OrderBy$
Data2.Refresh
Data2.Recordset.MoveFirst
If Status = 0 Then
  While Not Data2.Recordset.EOF And Not FlagExit = "Y"
    DoEvents
    If FlagAbort Then
        GoTo shutdown
    End If
    Data7.Recordset.AddNew
    Data7.Recordset("Season") = CSCombo1.Text
    Data7.Recordset("InfoReported") = "Fed Top 10"
    Data7.Recordset("StatType") = "KPI"
    Data7.Recordset("yr") = Data2.Recordset("yr")
    Data7.Recordset("mark") = Data2.Recordset("mark")
    Data7.Recordset("ringno") = Data2.Recordset("ringno")
    Data7.Recordset("code_pnt") = Data2.Recordset("code_pnt")
    Data7.Recordset("velocity") = Data2.Recordset("velocity")
    Data7.Recordset("RaceCode") = Data2.Recordset("RaceCode")
    Data7.Recordset("code_race") = Data2.Recordset("code_race")
    Data7.Recordset("posLoft") = Data2.Recordset("posLoft")
    Data7.Recordset("posClub") = Data2.Recordset("posClub")
    Data7.Recordset("posFed") = Data2.Recordset("posFed")
    Data7.Recordset("posSec") = Data2.Recordset("posSec")
    Data7.Recordset("CodeRating") = Data2.Recordset("CodeRating")
    Data7.Recordset("Distance") = Data2.Recordset("Distance")
    Data7.Recordset.Update
    Data2.Recordset.MoveNext
  Wend
  FlagExit$ = "N"
End If
Status = 0

'Get Pos=1 for club
msgPanel.Caption = "Please wait ... selecting Club results ..."
OrderBy$ = " AND PosClub > 1 AND PosClub < 6 ORDER BY PosClub ASC, yr, mark, ringno"
Data2.RecordSource = sCriteria$ & OrderBy$
Data2.Refresh
Data2.Recordset.MoveFirst
If Status = 0 Then
  While Not Data2.Recordset.EOF And Not FlagExit = "Y"
    DoEvents
    If FlagAbort Then
        GoTo shutdown
    End If
    Data7.Recordset.AddNew
    Data7.Recordset("Season") = CSCombo1.Text
    Data7.Recordset("InfoReported") = "Club Top 5"
    Data7.Recordset("StatType") = "KPI"
    Data7.Recordset("yr") = Data2.Recordset("yr")
    Data7.Recordset("mark") = Data2.Recordset("mark")
    Data7.Recordset("ringno") = Data2.Recordset("ringno")
    Data7.Recordset("code_pnt") = Data2.Recordset("code_pnt")
    Data7.Recordset("velocity") = Data2.Recordset("velocity")
    Data7.Recordset("RaceCode") = Data2.Recordset("RaceCode")
    Data7.Recordset("code_race") = Data2.Recordset("code_race")
    Data7.Recordset("posLoft") = Data2.Recordset("posLoft")
    Data7.Recordset("posClub") = Data2.Recordset("posClub")
    Data7.Recordset("posFed") = Data2.Recordset("posFed")
    Data7.Recordset("posSec") = Data2.Recordset("posSec")
    Data7.Recordset("CodeRating") = Data2.Recordset("CodeRating")
    Data7.Recordset("Distance") = Data2.Recordset("Distance")
    Data7.Recordset.Update
    Data2.Recordset.MoveNext
  Wend
  FlagExit$ = "N"
End If
Status = 0

'Get Pos=1 for Loft
msgPanel.Caption = "Please wait ... selecting Clock Birds ..."
OrderBy$ = " AND PosLoft = 1 ORDER BY yr, mark, ringno"
Data2.RecordSource = sCriteria$ & OrderBy$
Data2.Refresh
Data2.Recordset.MoveFirst
If Status = 0 Then
While Not Data2.Recordset.EOF And Not FlagExit = "Y"
    DoEvents
    If FlagAbort Then
        GoTo shutdown
    End If
    Data7.Recordset.AddNew
    Data7.Recordset("Season") = CSCombo1.Text
    Data7.Recordset("InfoReported") = "Clock Birds"
    Data7.Recordset("StatType") = "KPI"
    Data7.Recordset("yr") = Data2.Recordset("yr")
    Data7.Recordset("mark") = Data2.Recordset("mark")
    Data7.Recordset("ringno") = Data2.Recordset("ringno")
    Data7.Recordset("code_pnt") = Data2.Recordset("code_pnt")
    Data7.Recordset("velocity") = Data2.Recordset("velocity")
    Data7.Recordset("RaceCode") = Data2.Recordset("RaceCode")
    Data7.Recordset("code_race") = Data2.Recordset("code_race")
    Data7.Recordset("posLoft") = Data2.Recordset("posLoft")
    Data7.Recordset("posClub") = Data2.Recordset("posClub")
    Data7.Recordset("posFed") = Data2.Recordset("posFed")
    Data7.Recordset("posSec") = Data2.Recordset("posSec")
    Data7.Recordset("CodeRating") = Data2.Recordset("CodeRating")
    Data7.Recordset("Distance") = Data2.Recordset("Distance")
    Data7.Recordset.Update
    Data2.Recordset.MoveNext
  Wend
  FlagExit$ = "N"
End If
Status = 0

    msgPanel.Caption = "Please wait ... report being generated ..."
    Report1.ReportFileName = DBPATH$ & "kpi.rpt"
    Formula$ = "{stats.season}='" & CSCombo1.Text & "'"  'set orderby and selection of records
    Report1.SelectionFormula = Formula$
    
    ' Show printer setup dialog box if destination is printer.
    On Error GoTo Printererr3a
    If Report1.Destination = 1 Then
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
    
    'ShowWhat
    Report1.Action = 1
    DoEvents 'allow report to print
    GoTo shutdown

FMErr2aa:
    Select Case Err
    Case 3021
        'MsgBox "Database Empty", 16, "INFORMATION ONLY"
        Status = -1
        Resume Next
    Case 3022
        MsgBox "Record already exists", 16, "Add record Error"
        Resume Next
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
        cmdAbort.Visible = False
        cmdPrint2.Visible = True
        msgPanel.Caption = ""
        screen.MousePointer = 0
        Exit Sub
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
        GoTo shutdown
    Case Else
        MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
        cmdAbort.Visible = False
        cmdPrint2.Visible = True
        msgPanel.Caption = ""
        screen.MousePointer = 0
        Exit Sub
    End Select
    
    Resume Next

Printererr3a:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub

shutdown:
    cmdAbort.Visible = False
    cmdPrint2.Visible = True
    msgPanel.Caption = ""
    screen.MousePointer = 0
End Sub

Sub Form_Activate ()
screen.MousePointer = 0
End Sub

Sub Form_Load ()
Data1.DatabaseName = DBFullFileName$
Data1.RecordSource = "master"
Data2.DatabaseName = DBFullFileName$
Data2.RecordSource = "racetrans"
Data3.DatabaseName = DBFullFileName$
Data3.RecordSource = "race"
Data4.DatabaseName = DBFullFileName$
Data4.RecordSource = "colour"
Data5.DatabaseName = DBFullFileName$
Data5.RecordSource = "bloodline"
Data6.DatabaseName = DBFullFileName$
Data6.RecordSource = "ratings"
Data7.DatabaseName = DBPATH$ & "rpextra.mdb"
Data7.RecordSource = "stats"
label2.Caption = "Ensure Printer is connected and on line ready to print. This report prints straight to printer."
cmdAbort.Visible = False
cmdPrint2.Visible = True
End Sub

