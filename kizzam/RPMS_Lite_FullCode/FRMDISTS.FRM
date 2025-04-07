VERSION 2.00
Begin Form frmDistSelect 
   Caption         =   "Distance Calculation & Print"
   ClientHeight    =   4080
   ClientLeft      =   1605
   ClientTop       =   1605
   ClientWidth     =   5655
   Height          =   4590
   Left            =   1545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5655
   Top             =   1155
   Width           =   5775
   Begin SSPanel msgPanel 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3720
      Width           =   5655
   End
   Begin SSPanel Panel3D1 
      Alignment       =   4  'Right Justify - MIDDLE
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
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
         Left            =   5040
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
         Top             =   2640
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
         Left            =   5040
         Top             =   3120
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
         TabIndex        =   10
         Top             =   3000
         Width           =   1335
      End
      Begin SSCommand cmdPrint 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
         Width           =   1335
      End
      Begin CSComboBox CSComboSeries 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FRMDISTS.FRX:0000
         Contents        =   FRMDISTS.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   405
         Left            =   1920
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TextDisplay     =   0  'first visible
         Top             =   2400
         ValueCol        =   -1
         Width           =   2055
      End
      Begin Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Race Points to report on"
         Height          =   1095
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
         Begin SSOption optMyPoints 
            Caption         =   "Only those Race Points Marker as 'My Points'"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   600
            Width           =   4455
         End
         Begin SSOption optAllPoints 
            Caption         =   "All Points on File"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   4455
         End
      End
      Begin Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Report Type"
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   4335
         Begin SSOption optAll 
            Caption         =   "All Methods"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin SSOption optGeod 
            Caption         =   "Geoid Distances"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Series:"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdPrint_Click ()
    screen.MousePointer = 11
    msgPanel.Caption = "Please WAIT while distances are recalculated...."
    CalcDist 'KM procedure
    msgPanel.Caption = "Report now being formatted... Please WAIT..."
    If Trim(GLOWNER) = "Demonstration Copy" Then
        MsgBox "Sorry this is a DEMO version, Please register software to access this feature - Print Order for this software.", 16, "Please register software to access this feature"
        Exit Sub
    End If
    
    If CSComboSeries.Value = "" Then CSComboSeries.Value = "All"
    
    If UCase(CSComboSeries.Value) <> "ALL" Then
        Report1.SelectionFormula = "{points.series} = '" & CSComboSeries & "'"
    End If

    If Not optAllPoints Then
        If Report1.SelectionFormula = "" Then
            Report1.SelectionFormula = "{points.MyPoint} = 'Y'"
        Else
            Report1.SelectionFormula = "{points.series} = '" & CSComboSeries & "' AND {points.MyPoint} = 'Y'"
        End If
    End If

    If GLDistType = "Y" Then
        If optGeod Then
            Report1.ReportFileName = DBPATH$ & "dist_ydG.rpt"
        Else
            Report1.ReportFileName = DBPATH$ & "dist_yds.rpt"
        End If
    Else
        If optGeod Then
            Report1.ReportFileName = DBPATH$ & "distancG.rpt"
        Else
            Report1.ReportFileName = DBPATH$ & "distance.rpt"
        End If
    End If
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr6c
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
        screen.MousePointer = 0
        msgPanel.Caption = ""
        Exit Sub
    End If
    
    Report1.Action = 1
    msgPanel.Caption = ""
    DoEvents 'allow report to print
    'Report1.SortFields(0) = ""
    'Report1.SelectionFormula = ""
    msgPanel.Caption = ""
    Exit Sub

PrinterErr6c:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub
End Sub

Sub Form_Load ()
    If CSComboSeries.Value = "" Then CSComboSeries.Value = "All"
End Sub

