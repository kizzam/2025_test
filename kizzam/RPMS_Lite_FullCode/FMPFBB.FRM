VERSION 2.00
Begin Form FMpfbb 
   Caption         =   "Report - Points flown by bird"
   ClientHeight    =   5295
   ClientLeft      =   210
   ClientTop       =   1650
   ClientWidth     =   9000
   Height          =   5700
   Left            =   150
   LinkTopic       =   "Form6"
   ScaleHeight     =   5295
   ScaleWidth      =   9000
   Top             =   1305
   Width           =   9120
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4920
      Width           =   9015
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   720
         Options         =   0
         ReadOnly        =   -1  'True
         RecordSource    =   "racetrans"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1140
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
         Height          =   495
         Left            =   4680
         TabIndex        =   16
         Top             =   4320
         Width           =   2535
      End
      Begin SSCommand ButtPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   4320
         Width           =   2535
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   -1  'True
         BoundReportHeading=   "Points Flown by Bird"
         Connect         =   ""
         CopiesToPrinter =   1
         DataSource      =   "Data1"
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   -1  'True
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   360
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
         ReportSource    =   1  'Bound TrueGrid Control
         ReportTitle     =   "Points Flown"
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   4440
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
         Caption         =   "Order By"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   4320
         ShadowStyle     =   1  'Raised
         TabIndex        =   14
         Top             =   720
         Width           =   4575
         Begin SSOption scrIndVel 
            Caption         =   "Velocity"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin SSOption scrIndRating 
            Caption         =   "Rating"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin SSOption scrIndRel 
            Caption         =   "Release Date"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Show Only"
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   13
         Top             =   720
         Width           =   3975
         Begin SSCheck scrIndRace 
            Caption         =   "Race"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin SSCheck scrIndToss 
            Caption         =   "Toss"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin TextBox scrRingno 
         Height          =   285
         Left            =   7320
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin TextBox scrMark 
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin TextBox scrYr 
         Height          =   285
         Left            =   6000
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         BackColor       =   &H0080FFFF&
         DataSource      =   "Data1"
         Editable        =   0   'False
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeadingHeight   =   1
         Height          =   2535
         HorzLines       =   0  'None
         Layout          =   FMPFBB.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   6
         Top             =   1440
         UseBookmarks    =   -1  'True
         Width           =   8655
         WrapCellPointer =   0   'False
      End
      Begin TextBox Text1 
         Enabled         =   0   'False
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1200
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   2415
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bird : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
End
Option Explicit
Dim FgCntRaces As Long

Sub ButtCancel_Click ()
    Unload FMpfbb
End Sub

Sub ButtPrint_Click ()
Report1.Action = 1
DoEvents
End Sub

Sub Form_Activate ()
Text1.Text = scrYr.Text & "/" & scrMark.Text & "/" & scrRingno.Text
'Note put setselections in scrRingno.change event
'to stop getting corrupt database
End Sub

Sub Form_Load ()
Data1.DatabaseName = DBFullFileName$
End Sub

Function OrderBy () As String
    If scrIndRel Then
       OrderBy$ = " relDate"
    ElseIf scrIndRating Then
        OrderBy$ = " CodeRating"
    ElseIf scrIndVel Then
        OrderBy$ = " velocity"
    End If
End Function

Sub scrIndRace_Click (Value As Integer)
setselections
End Sub

Sub scrIndRating_Click (Value As Integer)
setselections

End Sub

Sub scrIndRel_Click (Value As Integer)
setselections

End Sub

Sub scrIndToss_Click (Value As Integer)
setselections
End Sub

Sub scrIndVel_Click (Value As Integer)
setselections

End Sub

Sub scrRingno_Change ()
setselections

End Sub

Sub setselections ()
    Dim s As String
    Dim o As String

    On Error GoTo ByErr9b

    s$ = ShowWhat()
    o$ = OrderBy()

    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = s$ & " ORDER BY " & o$
    Data1.Refresh
    Data1.Recordset.MoveLast
    FgCntRaces& = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst
    MsgPanel.Caption = "Count of races listed = " & Str(FgCntRaces&)
    Exit Sub

ByErr9b:
    If Err = 3021 Then 'No Current record
        MsgPanel.Caption = "No records for this selection"
        Exit Sub
    End If
    If Err = 380 Then
        Resume Next
    End If
    MsgBox "Error :" & Err & " " & Error$, 16, "Error"
    Resume Next
End Sub

Function ShowWhat () As String
    
    Dim s As String

    'Standard select
    s$ = "SELECT * FROM racetrans WHERE yr = "
    s$ = s$ & Val(scrYr.Text) & " AND mark = '" & scrMark.Text & "' AND ringno = '" & scrRingno.Text & "'"

    If scrIndToss And scrIndRace Then
        's$ = s$ & " AND RaceStatus = 'T' or RaceStatus = 'R'"
    ElseIf scrIndToss Then
        s$ = s$ & " AND RaceStatus = 'T'"
    ElseIf scrIndRace Then
        s$ = s$ & " AND RaceStatus = 'R'"
    End If
    ShowWhat$ = s$
End Function

