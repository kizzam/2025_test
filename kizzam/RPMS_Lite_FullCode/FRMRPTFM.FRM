VERSION 2.00
Begin Form frmRptFmt 
   Caption         =   "Format Report"
   ClientHeight    =   3600
   ClientLeft      =   1965
   ClientTop       =   1875
   ClientWidth     =   5895
   Height          =   4110
   Left            =   1905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5895
   Top             =   1425
   Width           =   6015
   Begin SSPanel Panel3D1 
      Alignment       =   4  'Right Justify - MIDDLE
      BackColor       =   &H00E0E0E0&
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin CommonDialog CMDialog1 
         Left            =   5280
         Top             =   3000
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
         Left            =   4680
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
         Top             =   3000
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
      Begin SSFrame Frame3D2 
         Caption         =   "Include"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   2655
         Left            =   120
         ShadowStyle     =   1  'Raised
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   "	"
            ColWidth        =   FRMRPTFM.FRX:0000
            Contents        =   FRMRPTFM.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   120
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   16
            Text            =   "All"
            TextDisplay     =   0  'first visible
            Top             =   1800
            ValueCol        =   -1
            Width           =   1455
         End
         Begin SSCheck CheckToss 
            Caption         =   "Toss"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   855
         End
         Begin SSCheck CheckRaces 
            Caption         =   "Races"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Value           =   -1  'True
            Width           =   975
         End
         Begin Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Width           =   1455
         End
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelInner      =   2  'Raised
         Caption         =   "HELP: Summary of Race results will be printed in the order you nominate on the left."
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1695
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
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   3000
         Width           =   1695
      End
      Begin SSCommand cmdPrint 
         Caption         =   "&Print Report"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   3000
         Width           =   1575
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Sort by"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   2655
         Left            =   1920
         ShadowStyle     =   1  'Raised
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         Begin SSOption byPosition 
            Caption         =   "Position"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin SSOption byRcePnt 
            Caption         =   "Race Point"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   1455
         End
         Begin SSOption byDate 
            Caption         =   "Release Date"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin SSOption byBloodline 
            Caption         =   "Bloodline"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1095
         End
         Begin SSOption byDam 
            Caption         =   "Dam"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1095
         End
         Begin SSOption bySire 
            Caption         =   "Sire"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1215
         End
         Begin SSOption byVelocity 
            Caption         =   "Velocity"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1200
            Width           =   975
         End
         Begin SSOption byCB 
            Caption         =   "Clock Bird"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   960
            Width           =   1455
         End
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdPrint_Click ()
    Dim Formula As String
    Dim Cnt1 As Integer
    Dim s As String
    Dim mClass As String
    Dim mYear As Integer
    
    Cnt1% = 0

    If CheckRaces Then 'Toss
        Report1.ReportFileName = DBPATH$ & "racelst1.rpt"
    Else
        Report1.ReportFileName = DBPATH$ & "racelst1.rpt"
    End If
    
    'OrderByPrt     'Set Order of records in Report
    
    Report1.SortFields(0) = ""
    If byDate Then
        'OrderBy$ = "yr, mark, ring_no"
        Report1.SortFields(Cnt1%) = "+{race.reldate}"
        Cnt1% = Cnt1 + 1
    ElseIf ByRcePnt Then
        'OrderBy$ = "yr, ring_no,mark"
        Report1.SortFields(Cnt1%) = "+{race.code_pnt}"
        Cnt1% = Cnt1 + 1
    ElseIf ByPosition Then
        'OrderBy$ = "ring_no, yr, mark"
        Report1.SortFields(Cnt1%) = "+{race.clubpos}"
        Cnt1% = Cnt1 + 1
    ElseIf ByCb Then
        'OrderBy$ = "mark, yr, ring_no"
        Report1.SortFields(Cnt1%) = "+{race.ringno}"
        Cnt1% = Cnt1 + 1
    ElseIf ByVelocity Then
        'OrderBy$ = "mark, yr, ring_no"
        Report1.SortFields(Cnt1%) = "+{race.clubvel}"
        Cnt1% = Cnt1 + 1
    ElseIf BySire Then
        'OrderBy$ = "mark, yr, ring_no"
        Report1.SortFields(Cnt1%) = "+{master.sireringno}"
        Cnt1% = Cnt1 + 1
    ElseIf ByDam Then
        'OrderBy$ = "mark, yr, ring_no"
        Report1.SortFields(Cnt1%) = "+{master.damringno}"
        Cnt1% = Cnt1 + 1
    ElseIf ByBloodline Then
        'OrderBy$ = "mark, yr, ring_no"
        Report1.SortFields(Cnt1%) = "+{bloodline.desc}"
        Cnt1% = Cnt1 + 1
    Else
       'OrderBy$ = "yr, mark, ring_no"
        Report1.SortFields(Cnt1%) = "+{race.reldate}"
        Cnt1% = Cnt1 + 1
    End If

    If CheckRaces And CheckToss Then
        s$ = ""
    ElseIf CheckRaces Then
        s$ = "{race.racestatus} = 'R'"
    Else s$ = "{race.racestatus} = 'Y'"
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

    Formula$ = s$

    
    'Formula$ = ShowWhatPrt() 'set orderby and selection of records
    'Formula$ = "{employee.EmployeeID} in '" & CSFrom.Text & "' to '" & CSTo.Text & "'"

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
        On Error GoTo PrinterErr6
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
    
    Report1.Action = 1
    DoEvents 'allow report to print
    Exit Sub

PrinterErr6:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

