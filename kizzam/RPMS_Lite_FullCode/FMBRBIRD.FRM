VERSION 2.00
Begin Form FMBRbird 
   Caption         =   "Report Selection (Breeding)"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9600
   Height          =   7335
   Left            =   -15
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9600
   Top             =   -90
   Width           =   9720
   Begin SSPanel MsgPanel 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   6480
      Width           =   9615
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9615
      Begin SSCommand ButtCancel 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   5040
         TabIndex        =   34
         Top             =   5520
         Width           =   2295
      End
      Begin SSCommand ButtDisplay 
         Caption         =   "Do &Report"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   2400
         TabIndex        =   33
         Top             =   5520
         Width           =   2295
      End
      Begin SSFrame Frame3D5 
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   42
         Top             =   720
         Width           =   8175
         Begin TextBox bscmonth 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5760
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   240
            Width           =   2055
         End
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FMBRBIRD.FRX:0000
            Contents        =   FMBRBIRD.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   1680
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   0
            Text            =   "ALL"
            TextDisplay     =   0  'first visible
            Top             =   240
            ValueCol        =   -1
            Width           =   1335
         End
         Begin SSOption MindRing 
            Caption         =   "Life Ring Year"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3360
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin SSOption Mindbreed 
            Caption         =   "Breeding Season Year"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3360
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select :"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
      End
      Begin SSFrame Frame3D4 
         Caption         =   "Report Type"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   41
         Top             =   4560
         Width           =   5655
         Begin SSOption mTypeNotes 
            Caption         =   "Notes"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4440
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin SSOption mtl 
            Caption         =   "Breeding (LS)"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1200
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin SSOption Mtypeall 
            Caption         =   "All"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin SSOption Mtypemaster 
            Caption         =   "Master"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2760
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
         End
         Begin SSOption Mtypebreed 
            Caption         =   "Breeding"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin CommonDialog CMDialog1 
         Left            =   240
         Top             =   1080
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Report Output"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   6600
         ShadowStyle     =   1  'Raised
         TabIndex        =   40
         Top             =   4560
         Width           =   2415
         Begin SSOption ToPrinter 
            Caption         =   "Printer"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin SSOption ToScreen 
            Caption         =   "Screen"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin CrystalReport Report1 
         BoundReportFooter=   0   'False
         BoundReportHeading=   ""
         Connect         =   ""
         CopiesToPrinter =   1
         DataSource      =   "Data1"
         Destination     =   0  'Window
         DetailCopies    =   1
         DiscardSavedData=   0   'False
         EMailCCList     =   ""
         EMailMessage    =   ""
         EMailSubject    =   ""
         EMailToList     =   ""
         EMailVIMBCCList =   ""
         GroupSelectionFormula=   ""
         Left            =   240
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
         Top             =   600
         UserName        =   ""
         WindowBorderStyle=   2  'Sizable
         WindowControlBox=   -1  'True
         WindowControls  =   -1  'True
         WindowHeight    =   480
         WindowLeft      =   0
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowParentHandle=   0
         WindowState     =   0  'normal
         WindowTitle     =   "Breeding Report"
         WindowTop       =   0
         WindowWidth     =   600
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "C:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "master"
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Report Order"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   1335
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   39
         Top             =   1560
         Width           =   8175
         Begin SSOption byDOB 
            Caption         =   "by Date of Birth"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   840
            Width           =   1815
         End
         Begin SSOption ByMarkNoYr 
            Caption         =   "by Mark/Ring no/Yr"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   600
            Width           =   2415
         End
         Begin SSOption ByGender 
            Caption         =   "by Gender"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin SSOption ByCage 
            Caption         =   "by Cage"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6840
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin SSOption ByMarkYrNo 
            Caption         =   "by Mark/Yr/Ring no."
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   2415
         End
         Begin SSOption ByNoYrMark 
            Caption         =   "by Ring No./Yr/Mark"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   840
            Width           =   2535
         End
         Begin SSOption ByYrNoMark 
            Caption         =   "by Yr/Ring no./Mark"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   600
            Width           =   2415
         End
         Begin SSOption ByYrMarkNo 
            Caption         =   "by Yr/Mark/Ring no."
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin SSOption ByDam 
            Caption         =   "by Dam"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   840
            Width           =   1215
         End
         Begin SSOption BySire 
            Caption         =   "by Sire"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   600
            Width           =   1095
         End
      End
      Begin SSFrame Frame3D3 
         Caption         =   "Select ONLY Birds to be included in Report by placing 'X' in box"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   35
         Top             =   3000
         Width           =   8175
         Begin SSOption Mallsex 
            Caption         =   "All"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6600
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin SSOption Munsexed 
            Caption         =   "UnSexed"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6600
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1215
         End
         Begin SSOption Mhens 
            Caption         =   "Hens"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6600
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   840
            Width           =   855
         End
         Begin SSOption Mcocks 
            Caption         =   "Cocks"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6600
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   600
            Width           =   975
         End
         Begin SSCheck Mindbredforstock 
            Caption         =   "Bred for Stock"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            TabIndex        =   21
            Top             =   360
            Width           =   1815
         End
         Begin SSCheck Mindstock 
            Caption         =   "Stock Bird"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1575
         End
         Begin SSCheck Mindrace 
            Caption         =   "Race Bird"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin SSCheck Mindsoh 
            Caption         =   "Current"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin SSCheck Mindsp2 
            Caption         =   "Special 2"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   1080
            Width           =   1695
         End
         Begin SSCheck Mindsp1 
            Caption         =   "Special 1"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin SSCheck Mindbpb 
            Caption         =   "Breeders Plate (B)"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
         Begin SSCheck Mindbpa 
            Caption         =   "Breeders Plate (A)"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   17
            Top             =   360
            Width           =   1815
         End
         Begin SSCheck Mindref 
            Caption         =   "Reference"
            Font3D          =   0  'None
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
         End
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BREEDING REPORT SELECTION CRITERIA"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Top             =   240
         Width           =   7455
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    Unload FMBRbird
End Sub

Sub ButtDisplay_Click ()
    Dim Formula As String
    
    If Mtypebreed Then
        Report1.ReportFileName = DBPATH$ & "breeder.rpt"
    ElseIf Mtypemaster Then
        Report1.ReportFileName = DBPATH$ & "master.rpt"
    ElseIf Mtypeall Then
        Report1.ReportFileName = DBPATH$ & "all.rpt"
    ElseIf Mtl Then
        Report1.ReportFileName = DBPATH$ & "breederl.rpt"
    ElseIf MtypeNotes Then
        Report1.ReportFileName = DBPATH$ & "brNotes.rpt"
    End If
    OrderBy 'Orders records in report
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
        On Error GoTo PrinterErr1
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

PrinterErr1:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

Sub Form_Activate ()
    ByYrMarkNo.Value = -1
End Sub

Sub Form_Load ()
    Dim Mend As Integer
    
    If GLbscmonth% = 1 Then
        Mend% = 12
    Else
        Mend% = GLbscmonth% - 1
    End If
    Data1.DatabaseName = DBFullFileName$
    Data1.RecordSource = "master"
    bscmonth.Text = StrMth(GLbscmonth%)
    bscmonth.Text = bscmonth.Text & " to " & StrMth(Mend%)
End Sub

Sub OrderBy ()
    Dim Cnt1 As Integer
    Cnt1% = -1
    Report1.SortFields(0) = ""

    If ByYrMarkNo Then
        'OrderBy$ = "yr, mark, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    ElseIf ByYrNoMark Then
        'OrderBy$ = "yr, ring_no,mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
    ElseIf ByNoYrMark Then
        'OrderBy$ = "ring_no, yr, mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
    ElseIf ByMarkYrNo Then
        'OrderBy$ = "mark, yr, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    ElseIf ByMarkNoYr Then
        'OrderBy$ = "mark, ring_no, yr"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
    ElseIf BySire Then
        'OrderBy$ = "sire_yr, sire_mark, sire_ring, ddob, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.sireyr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.siremark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.sireringno}"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.ddob}"
        'Report1.SortFields(Cnt1%) = "+{master.dob}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    ElseIf ByDam Then
        'OrderBy$ = "dam_yr, dam_mark, dam_ring, dob, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.damyr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.dammark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.damringno}"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.dob}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    ElseIf ByCage Then
        'OrderBy$ = "code_cage, yr, ring_no, mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.codecage}"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    ElseIf ByGender Then
        'OrderBy$ = "sex, yr, ring_no, mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.codesex}"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    ElseIf ByDOB Then
        'OrderBy$ = "dob, yr, ring_no, mark"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.ddob}"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    Else
       'OrderBy$ = "yr, mark, ring_no"
        Cnt1% = Cnt1 + 1
        Report1.SortFields(Cnt1%) = "+{master.yr}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.mark}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{@ringnoInt}"
        Cnt1% = Cnt1% + 1
        Report1.SortFields(Cnt1%) = "+{master.ringno}"
    End If
End Sub

Sub setselections ()

    Dim s As String
    Dim o As String

    On Error GoTo ByErr9

    s$ = ShowWhat()
   'o$ = OrderBy()

    If s$ = "" Then
       Data1.RecordSource = "SELECT * FROM master ORDER BY " & o$
    Else
       Data1.RecordSource = "SELECT * FROM master WHERE " & s$ & " ORDER BY " & o$
    End If
    Data1.Refresh
    Exit Sub
    
ByErr9:
    If Err = 380 Then
        Resume Next
    End If
    MsgBox "Error :" & Err & " " & Error$, 16, "Error"
    Exit Sub

End Sub

Function ShowWhat () As String

    Dim s As String
    Dim myr As String
    Dim miyr As Long
    Dim sdate As Long
    Dim edate As Long
    Dim zdate As String

    If Trim(CSCombo1.Value) = "ALL" Or Trim(CSCombo1.Value) = "" Then
        s$ = ""
    Else
        'Will need adjusting for Internalional use.
        myr$ = CSCombo1.Value
        zdate$ = "01/" & GLbscmonth% & "/" & LTrim(myr$)
        sdate& = DateValue(zdate$) 'Start date
        miyr& = Val(myr$)
        miyr& = miyr& + 1
        zdate$ = "01/" & GLbscmonth% & "/" & miyr&
        edate& = DateValue(zdate$) 'End date
        'edate& = Format(zdate$, "ddddd")'End date
        'myr$ = RTrim(LTrim(Str(miyr%)))
        'edate& = DateValue("01/04/" & myr$)
        's$ = "{CVDate(master.dob)} > " & sdate&
        's$ = s$ & " and {CVDate(master.dob)} < " & edate&
        
        If Mindbreed Then
            s$ = "{master.dob_li} > " & sdate&  'note should be >=
            s$ = s$ & " and {master.dob_li} < " & edate&
        Else
            s$ = "{master.yr} = " & CSCombo1.Value
        End If

    End If

    If Mindsoh Then
        If s$ = "" Then
        s$ = "{master.indsoh} = True"
       Else
        s$ = s$ & " and {master.indsoh} = True"
       End If
    End If
    
    If Mindrace Then
       If s$ = "" Then
        s$ = "{master.indrace} = True"
       Else
        s$ = s$ & " and {master.indrace} = True"
       End If
    End If
    If Mindstock Then
        If s$ = "" Then
        s$ = "{master.indstock} = True"
       Else
        s$ = s$ & " and {master.indstock} = True"
       End If
    End If

    If Mindref Then
       If s$ = "" Then
        s$ = "{master.indref} = True"
       Else
        s$ = s$ & " and {master.indref} = True"
       End If
    End If
    If Mindsp1 Then
       If s$ = "" Then
        s$ = "{master.indsp1} = True"
       Else
        s$ = s$ & " and {master.indsp1} = True"
       End If
    End If
    If Mindsp2 Then
       If s$ = "" Then
        s$ = "{master.indsp2} = True"
       Else
        s$ = s$ & " and {master.indsp2} = True"
       End If
    End If
    If Mindbpb Then
       If s$ = "" Then
        s$ = "{master.indbpb} = True"
       Else
        s$ = s$ & " and {master.indbpb} = True"
       End If
    End If
    If Mindbpa Then
       If s$ = "" Then
        s$ = "{master.indbpa} = True"
       Else
        s$ = s$ & " and {master.indbpa} = True"
       End If
    End If

    If Mhens Then
       If s$ = "" Then
          s$ = "{master.codesex} = 'H'"
       Else
          s$ = s$ & " and {master.codesex} = 'H'"
       End If
    ElseIf Mcocks Then
       If s$ = "" Then
          s$ = "{master.codesex} = 'C'"
       Else
          s$ = s$ & " and {master.codesex} = 'C'"
       End If
    ElseIf Munsexed Then
       If s$ = "" Then
          s$ = "{master.codesex} = 'U'"
       Else
          s$ = s$ & " and {master.codesex} = 'U'"
       End If
    End If

    ShowWhat$ = s$

End Function

Sub ToPrinter_Click (Value As Integer)
    Report1.Destination = 1
End Sub

Sub ToScreen_Click (Value As Integer)
    Report1.Destination = 0
End Sub

