VERSION 2.00
Begin Form FMresults 
   Caption         =   "Maintain Race Results"
   ClientHeight    =   6585
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   9450
   Height          =   6990
   Left            =   45
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   9450
   Top             =   30
   Width           =   9570
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6240
      Width           =   9495
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin SSFrame Frame3D2 
         Caption         =   "Sort"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   67
         Top             =   360
         Width           =   2295
         Begin SSOption Option3D2 
            Caption         =   "Velocity"
            Height          =   195
            Left            =   1200
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin SSOption Option3D1 
            Caption         =   "Position"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   975
         End
      End
      Begin CommonDialog CMDialog1 
         Left            =   480
         Top             =   5640
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
         Top             =   5520
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
      Begin SSCommand CmdPrevious 
         Caption         =   "&Previous Bird"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3840
         TabIndex        =   61
         Top             =   3120
         Width           =   1575
      End
      Begin SSCommand CmdNext 
         Caption         =   "&Next Bird"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         TabIndex        =   60
         Top             =   3120
         Width           =   1575
      End
      Begin TextBox mRaceStatus 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8640
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   120
         Width           =   735
      End
      Begin TextBox mBirdsEntered 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         TabIndex        =   57
         Top             =   480
         Width           =   615
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   8040
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "racetrans"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSCommand CmdPoints 
         Caption         =   "Points &Flown"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1200
         TabIndex        =   42
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin SSCommand CmdPrint 
         Caption         =   "&Print Results"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   48
         Top             =   5760
         Width           =   1935
      End
      Begin TextBox scrRaceClass 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         TabIndex        =   43
         Text            =   "Text7"
         Top             =   120
         Width           =   975
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Results for Bird -"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   2055
         Left            =   120
         ShadowStyle     =   1  'Raised
         TabIndex        =   35
         Top             =   3480
         Width           =   9135
         Begin TextBox scrDistance 
            Height          =   285
            Left            =   7680
            TabIndex        =   70
            Top             =   960
            Width           =   975
         End
         Begin SSCommand cmdCalcVel 
            Caption         =   "C"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8760
            TabIndex        =   66
            Top             =   1320
            Width           =   255
         End
         Begin TextBox scrClkDate 
            Height          =   285
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1320
            Width           =   1215
         End
         Begin TextBox scrFTsecs 
            Height          =   285
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   25
            Text            =   "Text2"
            Top             =   1680
            Width           =   495
         End
         Begin TextBox scrFTmins 
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   1680
            Width           =   375
         End
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ","
            ColWidth        =   FMRESULT.FRX:0000
            Contents        =   FMRESULT.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   315
            Left            =   5520
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   "Desc,Level,Code"
            ListDataSource  =   "Data2"
            SearchString    =   ""
            Sorted          =   -1  'True
            StartSearch     =   0
            TabIndex        =   26
            Text            =   "CSCombo1"
            TextDisplay     =   1  'all visible
            Top             =   1680
            ValueCol        =   2
            Width           =   2655
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
            RecordSource    =   "Ratings"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1140
         End
         Begin TextBox scrPosComb 
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   7
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrOOComb 
            Height          =   285
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrOOLoft 
            Height          =   285
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrPosLoft 
            Height          =   285
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   3
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrOOClub 
            Height          =   285
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrOOSec 
            Height          =   285
            Left            =   3600
            MaxLength       =   5
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrOOFed 
            Height          =   285
            Left            =   3000
            MaxLength       =   5
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrOONat 
            Height          =   285
            Left            =   4200
            MaxLength       =   5
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrPosINat 
            Height          =   285
            Left            =   4800
            MaxLength       =   4
            TabIndex        =   15
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrNotes 
            Height          =   375
            Left            =   5520
            MaxLength       =   24
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   480
            Width           =   3495
         End
         Begin TextBox scrPosClub 
            Height          =   285
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   5
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrPosSec 
            Height          =   285
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   11
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrPosFed 
            Height          =   285
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   9
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrPosNat 
            Height          =   285
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   13
            Top             =   480
            Width           =   495
         End
         Begin TextBox scrClkSecs 
            Height          =   285
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   20
            Top             =   1320
            Width           =   495
         End
         Begin TextBox scrClkMins 
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   19
            Top             =   1320
            Width           =   375
         End
         Begin TextBox scrOOINat 
            Height          =   285
            Left            =   4800
            MaxLength       =   5
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin TextBox scrVelocity 
            Height          =   285
            Left            =   7680
            MaxLength       =   6
            TabIndex        =   22
            Top             =   1320
            Width           =   975
         End
         Begin TextBox scrFThrs 
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   23
            Top             =   1680
            Width           =   375
         End
         Begin TextBox scrClkHrs 
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   18
            Top             =   1320
            Width           =   375
         End
         Begin Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Distance:"
            Height          =   255
            Left            =   6480
            TabIndex        =   71
            Top             =   960
            Width           =   1095
         End
         Begin Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Clock Date:"
            Height          =   255
            Left            =   4320
            TabIndex        =   65
            Top             =   1320
            Width           =   1095
         End
         Begin Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Secs"
            Height          =   255
            Left            =   3720
            TabIndex        =   64
            Top             =   1680
            Width           =   495
         End
         Begin Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Mins"
            Height          =   255
            Left            =   2640
            TabIndex        =   63
            Top             =   1680
            Width           =   495
         End
         Begin Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Hrs"
            Height          =   255
            Left            =   1800
            TabIndex        =   62
            Top             =   1680
            Width           =   375
         End
         Begin Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Out of:"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   840
            Width           =   975
         End
         Begin Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Rating:"
            Height          =   255
            Left            =   4680
            TabIndex        =   55
            Top             =   1680
            Width           =   735
         End
         Begin Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Comb:"
            Height          =   255
            Left            =   2400
            TabIndex        =   54
            Top             =   240
            Width           =   495
         End
         Begin Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Loft:"
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   240
            Width           =   495
         End
         Begin Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Notes:"
            Height          =   255
            Left            =   5520
            TabIndex        =   52
            Top             =   240
            Width           =   615
         End
         Begin Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Secs"
            Height          =   255
            Left            =   3720
            TabIndex        =   51
            Top             =   1320
            Width           =   495
         End
         Begin Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Mins"
            Height          =   255
            Left            =   2640
            TabIndex        =   50
            Top             =   1320
            Width           =   495
         End
         Begin Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Hrs"
            Height          =   255
            Left            =   1800
            TabIndex        =   49
            Top             =   1320
            Width           =   375
         End
         Begin Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Nat:"
            Height          =   255
            Left            =   4200
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
         Begin Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Fed:"
            Height          =   255
            Left            =   3000
            TabIndex        =   46
            Top             =   240
            Width           =   495
         End
         Begin Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Int'Nat:"
            Height          =   255
            Left            =   4800
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sec:"
            Height          =   255
            Left            =   3600
            TabIndex        =   41
            Top             =   240
            Width           =   495
         End
         Begin Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Club:"
            Height          =   255
            Left            =   1800
            TabIndex        =   40
            Top             =   240
            Width           =   495
         End
         Begin Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Flying Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   1095
         End
         Begin Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Position:"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   480
            Width           =   975
         End
         Begin Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Velocity:"
            Height          =   255
            Left            =   6840
            TabIndex        =   37
            Top             =   1320
            Width           =   855
         End
         Begin Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Clock Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1095
         End
      End
      Begin SSCommand CmdUpdate 
         Caption         =   "&Enter Results"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   34
         Top             =   3120
         Width           =   1575
      End
      Begin TextBox mRelPoint 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   480
         Width           =   1335
      End
      Begin TextBox mRaceCode 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
      Begin TextBox mRelDate 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   120
         Width           =   1335
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         BackColor       =   &H0080FFFF&
         DataSource      =   "Data1"
         Editable        =   0   'False
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   1  'By row
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         HeadingHeight   =   1
         Height          =   2175
         HorzLines       =   0  'None
         Layout          =   FMRESULT.FRX:0042
         LayoutIndex     =   1
         Left            =   120
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   2
         Top             =   840
         UseBookmarks    =   -1  'True
         Width           =   9135
         WrapCellPointer =   0   'False
      End
      Begin SSCommand CmdExit 
         Caption         =   "&Exit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   5760
         Width           =   1935
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Count of Birds entered:"
         Height          =   255
         Left            =   5280
         TabIndex        =   58
         Top             =   480
         Width           =   2175
      End
      Begin Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Class:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   5640
         TabIndex        =   44
         Top             =   120
         Width           =   1815
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Race Point:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
   End
End
Dim TblRace As Table
Dim TblPoint As Table
Dim DB As Database
Dim GFchangeind As Integer
Dim GFEntryCnt As Integer
Dim PntLng As Double
Dim PntLat As Double
Dim PntDist As Double
Dim PntDistGeod As Double

Sub cmdCalcVel_Click ()
Dim reply As Integer
Dim CalcA As Double
Dim CalcB As Double
Dim CalcC As Double
Dim CalcD As Double
Dim CalcE As Double
Dim MLatDeg As Double
Dim MLngDeg As Double
Dim RLatDeg As Double
Dim RLngDeg As Double
Dim MLatRad As Double
Dim MLngRad As Double
Dim RLatRad As Double
Dim RLngRad As Double
Dim DistGCM As Double
Dim MFT As Double
Dim msg As String
Dim DistVincenty As Double
Dim WGS84_f As Double
Dim WGS84_eps As Double

WGS84_f = 298.257223563
WGS84_eps = 6378137

If (Not IsNumeric(scrFTHrs.Text)) Or (Not IsNumeric(scrFTMins.Text)) Or (Not IsNumeric(scrFTSecs.Text)) Then
    reply% = MsgBox("Enter Flying time in Hours, Minutes && Seconds", 0, "Cannot calculate")
    Exit Sub
End If

If IsEmpty(scrFTHrs.Text) Or IsEmpty(scrFTMins.Text) Or IsEmpty(scrFTSecs.Text) Then
    reply% = MsgBox("Enter Flying time in Hours, Minutes && Seconds", 0, "Cannot calculate")
    Exit Sub
End If
  
If IsNull(scrFTHrs.Text) Or IsNull(scrFTMins.Text) Or IsNull(scrFTSecs.Text) Then
    reply% = MsgBox("Enter Flying time in Hours, Minutes && Seconds", 0, "Cannot calculate")
    Exit Sub
End If

MFT# = CDbl(scrFTHrs.Text) * 60
MFT# = MFT# + CDbl(scrFTMins.Text)
MFT# = MFT# + CDbl(scrFTSecs.Text) / 60

If PntDist# <> 0 Then
    If MFT# > 0 Then
        msg$ = "CLUB Distance provided = '" & Int(PntDist# * 1000) / 1000 & "' metres" & Chr(13) & "Velocity = '" & Int((PntDist# / MFT#) * 1000) / 1000 & "' metres per minute"
        reply% = MsgBox(msg$, 4, "Velocity calculated from Club supplied distance")
        If reply% = 6 Then
            scrVelocity.Text = CStr(Int((PntDist# / MFT#) * 1000) / 1000)
        End If
    End If
Else
    If PntLat = 0 Or PntLng = 0 Then
        reply% = MsgBox("Race point requires Coordinates entered for Longitude && Latitude", 0, "Cannot calculate")
        Exit Sub
    End If
    If GLLoftLat = 0 Or GLLoftLng = 0 Then
        reply% = MsgBox("Loft requires Coordinates entered for Longitude && Latitude", 0, "Cannot calculate")
        Exit Sub
    End If

    'Convert to degrees decimal
    MLatDeg# = GLLoftLat#
    MLngDeg# = GLLoftLng#

    'Convert to Radians
    MLatRad# = MakRad(MLatDeg#)
    MLngRad# = MakRad(MLngDeg#)

    'Set distance to Zero
    DistVincenty# = 0
        
    'Convert to degrees decimal
    RLatDeg# = PntLat#
    RLngDeg# = PntLng#
        
    'Convert Degrees to Radians
    RLatRad# = MakRad(RLatDeg#)
    RLngRad# = MakRad(RLngDeg#)

    DistVincenty# = fnVincenty(MLatRad#, MLngRad#, RLatRad#, RLngRad#, "WGS84", WGS84_f, WGS84_eps)
    If GLDistType = "Y" Then
        DistVincenty# = DistVincenty# * 1.09361  'Convert Metres to Yards
    End If

    scrDistance.Text = (Int(DistVincenty# * 1000) + .005) / 1000

    If MFT# > 0 Then
        If GLDistType = "Y" Then
            msg$ = "Calculated Geodesic Distance = '" & Int(DistVincenty# * 1000) / 1000 & "' yards" & Chr(13) & "Velocity = '" & Int((DistVincenty# / MFT#) * 1000) / 1000 & " yards per minute"
        Else
            msg$ = "Calculated Geodesic Distance = '" & Int(DistVincenty# * 1000) / 1000 & "' metres" & Chr(13) & "Velocity = '" & Int((DistVincenty# / MFT#) * 1000) / 1000 & " metres per minute"
        End If
        reply% = MsgBox(msg$, 4, "Calculated Distance & Velocity based on recorded coordinates")
        If reply% = 6 Then
            scrVelocity.Text = CStr(Int((DistVincenty# / MFT#) * 1000) / 1000)
        End If
    End If

End If

End Sub

Sub cmdExit_Click ()
Dim reply As Integer

If CmdUpdate.Caption = "&Accept Results" Then
    If GFchangeind% = -1 Then
        reply% = MsgBox("Do you want to abort changes?", 36, "ABORT CHANGES")
        If reply% = 6 Then
            Table1_RowChange
            CmdPrevious.Enabled = True
            CmdNext.Enabled = True
            CmdPoints.Enabled = True
            CmdPrint.Enabled = True
            CmdExit.Caption = "&Exit"
            CmdUpdate.Caption = "&Enter Results"
            Setforbrowse
            GFchangeind% = 0
            Exit Sub
        Else
            scrPosLoft.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "No Changes effected", 64, "Information Only"
        Setforbrowse
        Exit Sub
    End If
Else
    Unload FMresults
End If

End Sub

Sub CmdNext_Click ()
On Error GoTo BOKERR5

If Data1.Recordset.EOF Then
    MsgBox "Reached EOF", 64, "Information Only"
Else
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then
        Data1.Recordset.MoveLast
        Table1_RowChange
        MsgBox "Reached LAST record", 64, "Information Only"
    End If
End If

Exit Sub

BOKERR5:
    Select Case Err
    Case 3021
        MsgBox "No Birds Nominated for this Race" & NL$ & "Birds MUST be entered if RESULTS are to be recorded", 0, "RESULTS NOT ALLOWED"
        Unload Me
        Exit Sub
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
        Exit Sub
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
        Exit Sub
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMresult.frm"
        Exit Sub
    End Select
    Resume Next
End Sub

Sub CmdPrevious_Click ()

On Error GoTo BOKERR6

If Data1.Recordset.BOF Then
    MsgBox "Reached BOF", 64, "Information Only"
Else
    Data1.Recordset.MovePrevious
    If Data1.Recordset.BOF Then
        Data1.Recordset.MoveFirst
        Table1_RowChange
        MsgBox "On FIRST record already", 64, "Information Only"
    End If
End If
Exit Sub

BOKERR6:
    Select Case Err
    Case 3021
        MsgBox "No Birds Nominated for this Race" & NL$ & "Birds MUST be entered if RESULTS are to be recorded", 0, "RESULTS NOT ALLOWED"
        Unload Me
        Exit Sub
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
        Exit Sub
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
        Exit Sub
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMresult.frm"
        Exit Sub
    End Select
    Resume Next
End Sub

Sub cmdPrint_Click ()
'print results1.rpt
    Dim Formula As String
    Dim Cnt1 As Integer
    
    Cnt1 = -1

    Report1.ReportFileName = DBPATH$ & "results1.rpt"
    If Option3D1 Then 'Position
        'OrderBy$ = "yr, mark, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(0) = "+{racetrans.PosLoft}"
    ElseIf Option3D2 Then 'Velocity
        Cnt1% = Cnt1 + 1
        Report1.SortFields(0) = "+{racetrans.Velocity}"
    Else
        Cnt1% = Cnt1 + 1
        Report1.SortFields(0) = "+{@RingNoInt}"
    End If
    Formula$ = "{racetrans.RaceCode}=" & mRaceCode.Text  'set orderby and selection of records
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
    
    'ShowWhat
    Report1.Action = 1
    DoEvents 'allow report to print
    Exit Sub

PrinterErr3:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub


End Sub

Sub CmdUpdate_click ()

On Error GoTo ResultErr

If CmdUpdate.Caption = "&Accept Results" Then
    If GFchangeind% = 0 Then
        MsgBox "No Changes effected", 64, "Information Only"
        Setforbrowse
        Exit Sub
    Else    'GFchangeind% = 1
        'Validate DOB
        'if year is entered as 00 then assume 2000 otherwise accept
        'still have slight problem with 29/02/00 assumes 1900 for isdate function & reports as invalid
        'didn't worry about this

        If Trim(scrClkDate.Text) = "" Then
            scrClkDate.Text = "  /  /  "
            MsgBox "WARNING - Date can't be empty", 0, "Date Of Birth required"
            scrClkDate.SetFocus
            scrClkDate.SelStart = 0
            scrClkDate.SelLength = Len(scrClkDate.Text)
            Exit Sub
        End If
    
        'Fix potential leap year problem otherwise won't pass isdate() check
        If scrClkDate.Text = "29/2/00" Or scrClkDate.Text = "29/02/00" Then
            scrClkDate.Text = DateAdd("yyyy", 100, scrClkDate.Text)
        End If
    
        If IsDate(scrClkDate.Text) Then 'Must be Valid
            scrClkDate.Text = DateFix(scrClkDate.Text)
        Else
            MsgBox "WARNING - Clock-in Date Invalid", 0, "Clock-in Date format incorrect"
            scrClkDate.SetFocus
            scrClkDate.SelStart = 0
            scrClkDate.SelLength = Len(scrClkDate.Text)
            Exit Sub
        End If

        If CVDate(scrClkDate.Text) < CVDate(mRelDate.Text) Then
            MsgBox "Clock in date cannot be prior to release date!", 32, "INVALID CLOCK IN DATE"
            scrClkDate.SetFocus
            Exit Sub
        End If

        'Update record
        Data1.Recordset.Edit
        Data1.Recordset("PosLoft") = Val(scrPosLoft.Text)
        Data1.Recordset("PosClub") = Val(ScrPosClub.Text)
        Data1.Recordset("PosComb") = Val(ScrPosComb.Text)
        Data1.Recordset("PosFed") = Val(scrPosFed.Text)
        Data1.Recordset("PosSec") = Val(scrPosSec.Text)
        Data1.Recordset("PosNat") = Val(scrPosNat.Text)
        Data1.Recordset("PosINat") = Val(scrPosINat.Text)
        Data1.Recordset("ClkHr") = Val(scrClkHrs.Text)
        Data1.Recordset("ClkMin") = Val(scrClkMins.Text)
        Data1.Recordset("ClkSec") = Val(scrClkSecs.Text)
        Data1.Recordset("FThr") = Val(scrFTHrs.Text)
        Data1.Recordset("ClkDate") = CVDate(scrClkDate.Text)
        Data1.Recordset("FTmin") = Val(scrFTMins.Text)
        Data1.Recordset("FTsec") = Val(scrFTSecs.Text)
        Data1.Recordset("FlyTimeHrs") = Val(scrFTHrs.Text) + Val(scrFTMins.Text) / 60 + Val(scrFTSecs.Text) / 3600
        Data1.Recordset("Velocity") = Val(scrVelocity.Text)
        Data1.Recordset("Distance") = Val(scrDistance.Text)
        Data1.Recordset("CodeRating") = CSCombo1.Value
        Data1.Recordset("Notes") = Trim(ScrNotes.Text)
        Data1.Recordset.Update
        Setforbrowse
    End If
Else 'weren't updating so quit form requested
    setforupd
End If

Exit Sub

ResultErr:
    Select Case Err
    Case 0
        Resume Next
    Case 3021
        MsgBox "No Birds Nominated for this Race" & NL$ & "Birds MUST be entered if RESULTS are to be recorded", 0, "RESULTS NOT ALLOWED"
        Unload Me
        Exit Sub
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
        Exit Sub
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
        Exit Sub
    Case Else
        MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMresult.frm"
        Exit Sub
    End Select
    Resume Next
End Sub

Sub CSCombo1_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub Form_Activate ()
    On Error GoTo BOKERR
    
    Dim sCriteria As String
    Dim NL As String

    NL$ = Chr(13)

    Data1.DatabaseName = DBFullFileName$
    StrCriteria1$ = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY ringno, yr, mark"
    DoEvents
    Data1.RecordSource = StrCriteria1$
    Data1.Refresh
    FreeLocks
    DoEvents

    TblRace.Seek "=", Trim(mRaceCode.Text)
    DoEvents
    If TblRace.NoMatch Then
        MsgBox "Fatal Error #30001", 4, "Program Aborting"
        Stop
    Else
        If TblRace("RaceStatus") = "R" Then
            mRaceStatus.Text = "Race"
        Else
            mRaceStatus.Text = "Toss"
        End If

        scrRaceClass.Text = TblRace("Code_Race")
        
        If IsNull(TblRace("BiLoft")) Then
            ScrOOLoft.Text = ""
        Else
            ScrOOLoft.Text = TblRace("BiLoft")
        End If
        If IsNull(TblRace("BiClub")) Then
            ScrOOClub.Text = ""
        Else
            ScrOOClub.Text = TblRace("BIClub")
        End If
        If IsNull(TblRace("BiComb")) Then
            ScrOOComb.Text = ""
        Else
            ScrOOComb.Text = TblRace("BiComb")
        End If
        If IsNull(TblRace("BiFed")) Then
            ScrOOFed.Text = ""
        Else
            ScrOOFed.Text = TblRace("BiFed")
        End If

        If IsNull(TblRace("BiSec")) Then
            ScrOOSec.Text = ""
        Else
            ScrOOSec.Text = TblRace("BiSec")
        End If
        If IsNull(TblRace("BiNat")) Then
            ScrOONat.Text = ""
        Else
            ScrOONat.Text = TblRace("BiNat")
        End If
        If IsNull(TblRace("BiINat")) Then
            ScrOOINat.Text = ""
        Else
            ScrOOINat.Text = TblRace("BiINat")
        End If
    End If
    
    Data1.Recordset.MoveLast
    FreeLocks
    GFEntryCnt% = Data1.Recordset.RecordCount
    If GFEntryCnt% < 1 Then
        MsgBox "No Birds Nominated so results cannot be entered", 0, "CANNOT PROCEED WITH RESULTS"
        Unload Me
        Exit Sub
    End If
    mBirdsEntered.Text = GFEntryCnt%
    mRelDate.Text = Format(Data1.Recordset("reldate"), "ddddd")
    mRelPoint.Text = Data1.Recordset("code_pnt")
    CmdPrevious.Enabled = True
    CmdNext.Enabled = True
    CmdPoints.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Caption = "&Exit"
    Data1.Recordset.MoveFirst
    DoEvents
    FreeLocks
    DoEvents
    Table1.Refresh
    DoEvents

    TblPoint.Seek "=", TblRace("code_pnt")
    If TblPoint.NoMatch Then
        MsgBox "Fatal Error #30002", 4, "(FMResult) Program Aborting"
        Stop
    Else
        If IsNull(TblPoint("Lat_deg") Or TblPoint("Lat_min") Or TblPoint("Lat_sec")) Then
            PntLat = 0
            PntLng = 0
            Exit Sub
        ElseIf IsNull(TblPoint("Lng_deg") Or TblPoint("Lng_min") Or TblPoint("Lng_sec")) Then
            PntLat = 0
            PntLng = 0
            Exit Sub
        Else
            PntLat = CDbl((TblPoint("Lat_deg") + TblPoint("Lat_min") / 60 + (TblPoint("Lat_Sec") / 3600)))
            PntLng = CDbl((TblPoint("Lng_deg") + TblPoint("Lng_min") / 60 + (TblPoint("Lng_Sec") / 3600)))
        End If
        If IsNull(TblPoint("DistSpec")) Then
            PntDist# = 0
            PntDistGeod# = TblPoint("DistCalcGEOD")
        Else
            PntDist# = TblPoint("DistSpec")
            Exit Sub
        End If
    End If

    Exit Sub

BOKERR:
    Select Case Err
    Case 3021
        MsgBox "No Birds Nominated for this Race" & NL$ & "Birds MUST be entered if RESULTS are to be recorded", 0, "RESULTS NOT ALLOWED"
        Unload Me
        Exit Sub
    Case 3022
        MsgBox "Bird ALREADY selected for race" & NL$ & Format$(Data1.Recordset("yr"), "") & "-" & Data1.Recordset("mark") & "-" & Data1.Recordset("ringno"), 16, "Cannot add bird"
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

Sub Form_Load ()
DoEvents
FreeLocks
Set DB = OpenDatabase(DBFullFileName$, True, False)
DoEvents
FreeLocks

Data1.DatabaseName = DBFullFileName$
'Data1.RecordSource = "racetrans"
'Data1.Refresh
'DoEvents
'FreeLocks

Data2.DatabaseName = DBFullFileName$
Data2.RecordSource = "Ratings"
DoEvents
FreeLocks

Set TblRace = DB.OpenTable("Race")
TblRace.Index = "RaceCode"
DoEvents
FreeLocks
Set TblPoint = DB.OpenTable("Points")
TblPoint.Index = "Points"
DoEvents
FreeLocks

Setforbrowse

End Sub

Sub Form_Unload (Cancel As Integer)
TblRace.Close
TblPoint.Close
DB.Close
End Sub

Sub Option3D1_Click (Value As Integer)
    Data1.RecordSource = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY posloft DESC, ringno, yr, mark"
    Data1.Refresh
    FreeLocks
    DoEvents
End Sub

Sub Option3D2_Click (Value As Integer)
    Data1.RecordSource = "SELECT * FROM racetrans WHERE RaceCode = " & mRaceCode.Text & " ORDER BY velocity DESC, ringno, yr, mark"
    Data1.Refresh
    FreeLocks
    DoEvents
End Sub

Sub scrClkDate_Change ()
GFchangeind% = -1
End Sub

Sub scrClkDate_GotFocus ()
scrClkDate.SelStart = 0
scrClkDate.SelLength = Len(scrClkDate)
End Sub

Sub scrClkDate_LostFocus ()
    Dim dd As String, mm As String
    Dim yy As String
    Dim X As Integer

    If Len(scrClkDate.Text) > 0 Then
        X% = InStr(scrClkDate.Text, "/")
        If X% = 0 Then                      'Entered Date doesn't contain /
            If Trim(Len(scrClkDate.Text)) = 4 Then
                scrClkDate.Text = "0" & Trim(scrClkDate.Text)
                If Len(scrClkDate.Text) = 5 Then
                    scrClkDate.Text = Left$(Trim(scrClkDate.Text), 2) & "0" & Right$(scrClkDate.Text, 3)
                End If
            End If
            If Len(scrClkDate.Text) = 5 Then
                scrClkDate.Text = "0" & scrClkDate.Text
            End If
            If Len(scrClkDate.Text) = 7 Then
                scrClkDate.Text = "0" & scrClkDate.Text
            End If
            If Len(scrClkDate.Text) = 6 Then
                dd$ = Mid$(scrClkDate.Text, 1, 2)
                mm$ = Mid$(scrClkDate.Text, 3, 2)
                yy$ = Mid$(scrClkDate.Text, 5, 2)
            ElseIf Len(scrClkDate.Text) = 8 Then
                dd$ = Mid$(scrClkDate.Text, 1, 2)
                mm$ = Mid$(scrClkDate.Text, 3, 2)
                yy$ = Mid$(scrClkDate.Text, 5, 4)
            End If
            If Len(yy$) = 2 Then
                If Val(yy$) > 80 Then
                    yy$ = "19" & yy$
                Else
                    yy$ = "20" & yy$
                End If
            End If
            scrClkDate.Text = dd$ & "/" & mm$ & "/" & yy$
        End If
    End If
    If Year(Now) < "2000" Then
    End If
End Sub

Sub scrClkHrs_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrClkHrs_GotFocus ()
scrClkHrs.SelStart = 0
scrClkHrs.SelLength = Len(scrClkHrs)
End Sub

Sub scrClkMins_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrClkMins_GotFocus ()
scrClkMins.SelStart = 0
scrClkMins.SelLength = Len(scrClkMins)
End Sub

Sub scrClkSecs_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrClkSecs_GotFocus ()
scrClkSecs.SelStart = 0
scrClkSecs.SelLength = Len(scrClkSecs)

End Sub

Sub scrFThrs_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrFThrs_GotFocus ()
scrFTHrs.SelStart = 0
scrFTHrs.SelLength = Len(scrFTHrs)

End Sub

Sub scrFTmins_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrFTmins_GotFocus ()
scrFTMins.SelStart = 0
scrFTMins.SelLength = Len(scrFTMins)

End Sub

Sub scrFTsecs_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrFTsecs_GotFocus ()
scrFTSecs.SelStart = 0
scrFTSecs.SelLength = Len(scrFTSecs)

End Sub

Sub scrNotes_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrOOClub_GotFocus ()
ScrOOClub.SelStart = 0
ScrOOClub.SelLength = Len(ScrOOClub)
End Sub

Sub scrOOComb_GotFocus ()
ScrOOComb.SelStart = 0
ScrOOComb.SelLength = Len(ScrOOComb)

End Sub

Sub scrOOLoft_GotFocus ()
ScrOOLoft.SelStart = 0
ScrOOLoft.SelLength = Len(ScrOOLoft)
End Sub

Sub scrPosClub_Change ()
GFchangeind% = -1
End Sub

Sub scrPosClub_GotFocus ()
ScrPosClub.SelStart = 0
ScrPosClub.SelLength = Len(ScrPosClub)
End Sub

Sub scrPosComb_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"
End Sub

Sub scrPosComb_GotFocus ()
ScrPosComb.SelStart = 0
ScrPosComb.SelLength = Len(ScrPosComb)
End Sub

Sub scrPosFed_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrPosFed_GotFocus ()
scrPosFed.SelStart = 0
scrPosFed.SelLength = Len(scrPosFed)

End Sub

Sub scrPosINat_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrPosINat_GotFocus ()
scrPosINat.SelStart = 0
scrPosINat.SelLength = Len(scrPosINat)

End Sub

Sub scrPosLoft_Change ()
GFchangeind% = -1
End Sub

Sub scrPosLoft_GotFocus ()
scrPosLoft.SelStart = 0
scrPosLoft.SelLength = Len(scrPosLoft)
End Sub

Sub scrPosNat_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrPosNat_GotFocus ()
scrPosNat.SelStart = 0
scrPosNat.SelLength = Len(scrPosNat)

End Sub

Sub scrPosSec_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"
End Sub

Sub scrPosSec_GotFocus ()
scrPosSec.SelStart = 0
scrPosSec.SelLength = Len(scrPosSec)

End Sub

Sub scrVelocity_Change ()
GFchangeind% = -1
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"

End Sub

Sub scrVelocity_GotFocus ()
scrVelocity.SelStart = 0
scrVelocity.SelLength = Len(scrVelocity)

End Sub

Sub Setforbrowse ()
Table1.Enabled = True
scrPosLoft.Enabled = False
ScrOOLoft.Enabled = False
ScrPosClub.Enabled = False
ScrOOClub.Enabled = False
ScrPosComb.Enabled = False
ScrOOComb.Enabled = False
scrPosSec.Enabled = False
ScrOOSec.Enabled = False
scrPosFed.Enabled = False
ScrOOFed.Enabled = False
scrPosNat.Enabled = False
ScrOONat.Enabled = False
scrPosINat.Enabled = False
ScrOOINat.Enabled = False
ScrNotes.Enabled = False
scrClkHrs.Enabled = False
scrClkMins.Enabled = False
scrClkSecs.Enabled = False
scrFTHrs.Enabled = False
scrFTMins.Enabled = False
scrFTSecs.Enabled = False
CSCombo1.Enabled = False 'Rating
scrVelocity.Enabled = False
scrClkDate.Enabled = False
CmdPrevious.Enabled = True
CmdNext.Enabled = True
CmdPoints.Enabled = True
CmdPrint.Enabled = True
CmdUpdate.Enabled = True
cmdCalcVel.Enabled = False
CmdExit.Enabled = True
CmdExit.Caption = "&Exit"
CmdUpdate.Caption = "&Enter Results"
GFchangeind% = 0
'Table1.SetFocus
End Sub

Sub setforupd ()
Table1.Enabled = False
scrPosLoft.Enabled = True
ScrPosClub.Enabled = True
ScrPosComb.Enabled = True
scrPosSec.Enabled = True
scrPosFed.Enabled = True
scrPosNat.Enabled = True
scrPosINat.Enabled = True
ScrNotes.Enabled = True
scrClkHrs.Enabled = True
scrClkMins.Enabled = True
scrClkSecs.Enabled = True
scrFTHrs.Enabled = True
scrFTMins.Enabled = True
scrFTSecs.Enabled = True
CSCombo1.Enabled = True 'ratings
scrVelocity.Enabled = True
scrClkDate.Enabled = True
CmdPrevious.Enabled = False
CmdNext.Enabled = False
CmdPoints.Enabled = False
CmdPrint.Enabled = False
CmdExit.Caption = "&Cancel Changes"
CmdUpdate.Caption = "&Accept Results"
If scrClkDate.Text = "" Then
    scrClkDate.Text = Format$(Now, "ddddd")
End If
cmdCalcVel.Enabled = True
GFchangeind% = 0
scrPosLoft.SetFocus
End Sub

Sub Table1_RowChange ()
Frame3D1.Caption = "Results for Bird - " & Str$(Data1.Recordset("yr")) & "/" & Trim(Data1.Recordset("mark")) & "/" & Trim(Data1.Recordset("RingNo"))
If IsNull(Data1.Recordset("PosLoft")) Then
    scrPosLoft.Text = ""
Else
    scrPosLoft.Text = Str(Data1.Recordset("PosLoft"))
End If
If IsNull(Data1.Recordset("PosClub")) Then
    ScrPosClub.Text = ""
Else
    ScrPosClub.Text = Str(Data1.Recordset("PosClub"))
End If
'
If IsNull(Data1.Recordset("PosComb")) Then
    ScrPosComb.Text = ""
Else
    ScrPosComb.Text = Str(Data1.Recordset("PosComb"))
End If
If IsNull(Data1.Recordset("PosFed")) Then
    scrPosFed.Text = ""
Else
    scrPosFed.Text = Str(Data1.Recordset("PosFed"))
End If
If IsNull(Data1.Recordset("PosSec")) Then
    scrPosSec.Text = ""
Else
    scrPosSec.Text = Str(Data1.Recordset("PosSec"))
End If
If IsNull(Data1.Recordset("PosNat")) Then
    scrPosNat.Text = ""
Else
    scrPosNat.Text = Str(Data1.Recordset("PosNat"))
End If
If IsNull(Data1.Recordset("PosINat")) Then
    scrPosINat.Text = ""
Else
    scrPosINat.Text = Str(Data1.Recordset("PosINat"))
End If

If IsNull(Data1.Recordset("ClkDate")) Then
    scrClkDate.Text = ""
Else
    scrClkDate.Text = Format$(Data1.Recordset("ClkDate"), "ddddd")
End If
If IsNull(Data1.Recordset("ClkHr")) Then
    scrClkHrs.Text = ""
Else
    scrClkHrs.Text = Trim(Str(Data1.Recordset("ClkHr")))
End If
If IsNull(Data1.Recordset("ClkMin")) Then
    scrClkMins.Text = ""
Else
    scrClkMins.Text = Trim(Str(Data1.Recordset("ClkMin")))
End If
If IsNull(Data1.Recordset("ClkSec")) Then
    scrClkSecs.Text = ""
Else
    scrClkSecs.Text = Trim(Str(Data1.Recordset("ClkSec")))
End If
If IsNull(Data1.Recordset("FThr")) Then
    scrFTHrs.Text = ""
Else
    scrFTHrs.Text = Trim(Str(Data1.Recordset("FThr")))
End If
If IsNull(Data1.Recordset("FTmin")) Then
    scrFTMins.Text = ""
Else
    scrFTMins.Text = Trim(Str(Data1.Recordset("FTmin")))
End If
If IsNull(Data1.Recordset("FTsec")) Then
    scrFTSecs.Text = ""
Else
    scrFTSecs.Text = Trim(Str(Data1.Recordset("FTsec")))
End If

If IsNull(Data1.Recordset("notes")) Then
    ScrNotes.Text = ""
Else
    ScrNotes.Text = Trim(Data1.Recordset("notes"))
End If

If IsNull(Data1.Recordset("velocity")) Then
    scrVelocity.Text = ""
Else
    scrVelocity.Text = Str(Data1.Recordset("velocity"))
End If
If IsNull(Data1.Recordset("CodeRating")) Then
    CSCombo1.Value = ""
Else
    CSCombo1.Text = ""
    CSCombo1.Value = Data1.Recordset("CodeRating")
    CSCombo1.Refresh
End If

GFchangeind% = 0
CmdPrevious.Enabled = True
CmdNext.Enabled = True
CmdPoints.Enabled = True
CmdPrint.Enabled = True
CmdExit.Caption = "&Exit"
End Sub

