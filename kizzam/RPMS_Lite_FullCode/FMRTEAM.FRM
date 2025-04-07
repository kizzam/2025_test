VERSION 2.00
Begin Form FMRTeam 
   Caption         =   "Report Team selection (fmrteam.rpt)"
   ClientHeight    =   5625
   ClientLeft      =   885
   ClientTop       =   870
   ClientWidth     =   7755
   Height          =   6135
   Left            =   825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7755
   Top             =   420
   Width           =   7875
   Begin SSPanel MsgPanel 
      BackColor       =   &H00C0FFC0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   5280
      Width           =   7815
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0FFC0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin SSFrame Frame3D2 
         Caption         =   "Report Layout Selection"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   2175
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   16
         Top             =   2280
         Width           =   5415
         Begin SSOption indLayoutL6 
            Caption         =   "Landscape - Report on Birds with Team Status = ""L"""
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1800
            Width           =   4935
         End
         Begin SSOption indLayoutL5 
            Caption         =   "Landscape - Report on Birds with Team Status <> ""L"""
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1560
            Width           =   4935
         End
         Begin SSOption indLayoutL4 
            Caption         =   "Landscape - Bredby Report (Handy for Breeders Plates)"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1320
            Width           =   5055
         End
         Begin SSOption indLayoutL3 
            Caption         =   "Landscape - LARGE PRINT && Boxes"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3855
         End
         Begin SSOption indLayOutL2 
            Caption         =   "Landscape - LARGE PRINT"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   840
            Width           =   3735
         End
         Begin SSOption indLayoutP 
            Caption         =   "Portrait - less information (TeamsP.rpt)"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   4215
         End
         Begin SSOption indLayoutL 
            Caption         =   "Landscape - more information (TeamsL.rpt)"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   4215
         End
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "teams"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMRTEAM.FRX:0000
         Contents        =   FMRTEAM.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   405
         Left            =   2160
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "description,code"
         ListDataSource  =   "Data1"
         SearchString    =   ""
         StartSearch     =   0
         TabIndex        =   1
         TextDisplay     =   0  'first visible
         Top             =   480
         ValueCol        =   1
         Width           =   4095
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Sort Birds by"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   14
         Top             =   960
         Width           =   5415
         Begin SSOption bySire 
            Caption         =   "by Sire"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
         End
         Begin SSOption byMarkNoYr 
            Caption         =   "by Mark/No./Yr"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin SSOption byYrNoMark 
            Caption         =   "by Yr/No./Mark"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   720
            Width           =   1695
         End
         Begin SSOption byNoYrMark 
            Caption         =   "by No./Yr/Mark"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin SSOption byMarkYrNo 
            Caption         =   "by Mark/Yr/No."
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   1935
         End
         Begin SSOption byYrMarkNo 
            Caption         =   "by Yr/Mark/No."
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   0   'False
         BoundReportHeading=   ""
         Connect         =   ""
         CopiesToPrinter =   1
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   -1  'True
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   120
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
         Top             =   3960
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
         Left            =   240
         Top             =   3120
      End
      Begin SSCommand CmdExit 
         Caption         =   "&Exit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   3840
         TabIndex        =   12
         Top             =   4560
         Width           =   1815
      End
      Begin SSCommand CmdReport 
         Caption         =   "&Do Report"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1920
         TabIndex        =   11
         Top             =   4560
         Width           =   1815
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Team for PRINTING:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

Sub CmdReport_Click ()
    Dim Formula As String
    
    If CSCombo1.Text = "" Then
        MsgBox "MUST Select TEAM to print", 0, "NOTHING TO PRINT"
        CSCombo1.SetFocus
        Exit Sub
    End If

    If indLayoutL Then
        Report1.ReportFileName = DBPATH$ & "teamL.rpt"
    ElseIf indLayoutP Then
        Report1.ReportFileName = DBPATH$ & "teamP.rpt"
    ElseIf indLayoutL2 Then
        Report1.ReportFileName = DBPATH$ & "teamL2.rpt"
    ElseIf indLayoutL3 Then
        Report1.ReportFileName = DBPATH$ & "teamL3.rpt"
    ElseIf indLayoutL4 Then
        Report1.ReportFileName = DBPATH$ & "teamL4.rpt"
    ElseIf indLayoutL5 Then
        Report1.ReportFileName = DBPATH$ & "teamL5.rpt"
    ElseIf indLayoutL6 Then
        Report1.ReportFileName = DBPATH$ & "teamL6.rpt"
    Else
        Report1.ReportFileName = DBPATH$ & "teamP.rpt"
    End If
    OrderBy     'Set Order of records in Report
    Formula$ = ShowWhat() 'set orderby and selection of records
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
        On Error GoTo PrinterErr2
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

PrinterErr2:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub


End Sub

Sub Form_Activate ()
    ByNoYrMark.Value = -1
End Sub

Sub Form_Load ()
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "Select * from teams order by code DESC"
    Data1.Refresh
    DoEvents
End Sub

Sub OrderBy ()
    Dim Cnt1 As Integer
    Cnt1% = -1
    Report1.SortFields(0) = ""

    If ByYrMarkNo Then
        'OrderBy$ = "yr, mark, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{teambird.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.ringno}"
    ElseIf ByYrNoMark Then
        'OrderBy$ = "yr, ring_no,mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{teambird.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.ringno}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.mark}"
    ElseIf ByNoYrMark Then
        'OrderBy$ = "ring_no, yr, mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.ringno}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.mark}"
    ElseIf ByMarkYrNo Then
        'OrderBy$ = "mark, yr, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{teambird.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.ringno}"
    ElseIf BySire Then
        'OrderBy$ = "mark, yr, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.sireringno}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.sireyr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.siremark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
    Else
       'OrderBy$ = "yr, mark, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{teambird.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{teambird.ringno}"
    End If
End Sub

Function ShowWhat () As String

    Dim s As String
    Dim mteam As String

    mteam$ = CSCombo1.Value

    s$ = "{teambird.code} = '" & mteam$ & "'"

    ShowWhat$ = s$

End Function

