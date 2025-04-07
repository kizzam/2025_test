VERSION 2.00
Begin Form FMRmaster 
   Caption         =   "Report Master List (FMRmaster)"
   ClientHeight    =   6645
   ClientLeft      =   255
   ClientTop       =   375
   ClientWidth     =   9090
   Height          =   7050
   Left            =   195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9090
   Top             =   30
   Width           =   9210
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Caption         =   "More reports are available in each module"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   6240
      Width           =   9135
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9135
      Begin SSFrame Frame3D7 
         Caption         =   "Performance"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   1215
         Left            =   6480
         ShadowStyle     =   1  'Raised
         TabIndex        =   33
         Top             =   4800
         Width           =   2415
         Begin SSCommand cmdRCard 
            Caption         =   "Report &Card"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   2055
         End
      End
      Begin SSPanel MsgPanelBox 
         BevelInner      =   2  'Raised
         ForeColor       =   &H00FF0000&
         Height          =   4575
         Left            =   3360
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   5655
      End
      Begin SSFrame Frame3D6 
         Caption         =   "Team Reports"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   1095
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   31
         Top             =   2880
         Width           =   2895
         Begin SSCommand CmdPrtTeams 
            Caption         =   "Select &Team Reports"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   2415
         End
      End
      Begin SSFrame Frame3D5 
         Caption         =   "Breeding Reports"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   2655
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   30
         Top             =   120
         Width           =   2895
         Begin SSCommand cmdPrtLoft 
            Caption         =   "Current birds by Loft/Cage"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   2415
         End
         Begin SSCommand cmdPrtBreedA 
            Caption         =   "Quick Breeding by Matings"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   720
            Width           =   2415
         End
         Begin SSCommand CmdPrtRace 
            Caption         =   "&Quick Race Bird List"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   1680
            Width           =   2415
         End
         Begin SSCommand CmdPrtStock 
            Caption         =   "Quick &Stock Bird List"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   1200
            Width           =   2415
         End
         Begin SSCommand CmdPrtBreeding 
            Caption         =   "Select &Breeding Reports"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   2415
         End
      End
      Begin SSFrame Frame3D4 
         Caption         =   "Race Reports"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   2655
         Left            =   3360
         ShadowStyle     =   1  'Raised
         TabIndex        =   29
         Top             =   1800
         Width           =   2895
         Begin SSCommand CmdPrtTosses 
            Caption         =   "Quick &Toss List"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   2040
            Width           =   2415
         End
         Begin SSCommand CmdPrtRceList 
            Caption         =   "Q&uick Race List"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   2415
         End
         Begin SSCommand CmdPrtRaces 
            Caption         =   "Print R&ace List"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   2415
         End
         Begin SSCommand CmdPrtResults 
            Caption         =   "Print &Race Results"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   2415
         End
      End
      Begin SSFrame Frame3D3 
         Caption         =   "Miscellaneous Reports"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   2055
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   28
         Top             =   4080
         Width           =   2895
         Begin SSCommand cmdSaleRpt 
            Caption         =   "Produce &Sale Catalogue"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   2415
         End
         Begin SSCommand CmdPrtOrder 
            Caption         =   "Print &Order for Software"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   9
            Top             =   1440
            Width           =   2415
         End
         Begin SSCommand CmdDistances 
            Caption         =   "&Distances"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   2415
         End
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Contact List Reports"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   1575
         Left            =   3360
         ShadowStyle     =   1  'Raised
         TabIndex        =   27
         Top             =   120
         Width           =   2895
         Begin SSCommand CmdPrtCon1 
            Caption         =   "Print &Concise Contact List"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   2415
         End
         Begin SSCommand CmdPrtFullCon 
            Caption         =   "Print Full Contact &Details"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   2415
         End
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
         Left            =   960
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
         Top             =   5880
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
         Left            =   360
         Top             =   5880
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Reference File listings"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   4575
         Left            =   6480
         ShadowStyle     =   1  'Raised
         TabIndex        =   26
         Top             =   120
         Width           =   2415
         Begin SSCommand CmdPrtTeam 
            Caption         =   "Print Team C&odes"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   22
            Top             =   3960
            Width           =   2055
         End
         Begin SSCommand CmdPrtRatings 
            Caption         =   "Print Ra&tings"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   21
            Top             =   3360
            Width           =   2055
         End
         Begin SSCommand CmdPrtPoints 
            Caption         =   "Print Race &Points"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   20
            Top             =   2760
            Width           =   2055
         End
         Begin SSCommand CmdPrtBL 
            Caption         =   "Print &Bloodlines"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   2160
            Width           =   2055
         End
         Begin SSCommand CmdPrtColours 
            Caption         =   "C&olour Codes"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
         Begin SSCommand CmdPrtCages 
            Caption         =   "Print Ca&ges"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   2055
         End
         Begin SSCommand CmdPrtES 
            Caption         =   "Print &Eye Sign"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   240
            TabIndex        =   18
            Top             =   1560
            Width           =   2055
         End
      End
      Begin SSCommand CmdExit 
         Caption         =   "&Exit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   3360
         TabIndex        =   0
         Top             =   4800
         Width           =   2895
      End
   End
End
Option Explicit

Sub cmdDistances_Click ()
    frmDistSelect.Show 1
End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub CmdPrtBL_Click ()
FMPcodes.task.Text = "B"
FMPcodes.Show 1
End Sub

Sub cmdPrtBreedA_Click ()
fmdsprpt.Show 1
End Sub

Sub CmdPrtBreeding_Click ()
FMBRbird.Show 1
End Sub

Sub CmdPrtCages_Click ()
FMPcodes.task.Text = "A"
FMPcodes.Show 1
End Sub

Sub CmdPrtColours_Click ()
FMPcodes.task.Text = "C"
FMPcodes.Show 1
End Sub

Private Sub CmdPrtCon1_Click ()
    
    Report1.ReportFileName = DBPATH$ & "conshort.rpt"
    
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
    
    Report1.Action = 1
    DoEvents 'allow report to print
    Exit Sub

PrinterErr1:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
End Sub

Sub CmdPrtES_Click ()
FMPcodes.task.Text = "E"
FMPcodes.Show 1
End Sub

Private Sub CmdPrtFullCon_Click ()
    
    Report1.ReportFileName = DBPATH$ & "con50916.rpt"
    
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
    
    Report1.Action = 1
    DoEvents 'allow report to print
    Exit Sub

PrinterErr2:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
    

End Sub

Sub cmdPrtLoft_Click ()
    Report1.ReportFileName = DBPATH$ & "BreederA.rpt"
    
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr1cA
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

PrinterErr1cA:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
End Sub

Sub CmdPrtOrder_Click ()
    PrtOrder
End Sub

Sub CmdPrtPoints_Click ()
FMPcodes.task.Text = "R"
FMPcodes.Show 1
End Sub

Sub CmdPrtRace_Click ()
    Report1.ReportFileName = DBPATH$ & "GenRace.rpt"
    
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr1d
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

PrinterErr1d:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
    


End Sub

Sub CmdPrtRaces_Click ()
Dim Msg As String
Dim NL As String
NL$ = Chr(13)

If MsgPanelBox.Visible = True Then
    MsgPanelBox.Visible = False
Else
    Msg$ = "Run these reports from the RACE MODULE." & NL$
    Msg$ = Msg$ & " Click on Detail or Summary report button available on the RACE LIST screen." & NL$
    Msg$ = Msg$ & " Select Race or Toss from Race List screen and also Year of races if required." & NL$
    Msg$ = Msg$ & " These reports also show results on a race basis."
    MsgPanelBox.Caption = Msg$
    MsgPanelBox.Visible = True
End If

End Sub

Sub CmdPrtRatings_Click ()
FMPcodes.task.Text = "I"
FMPcodes.Show 1
End Sub

Sub CmdPrtRceList_Click ()
    Report1.ReportFileName = DBPATH$ & "racelst1.rpt"
    Report1.SortFields(0) = "-{race.reldate}"
    Report1.SelectionFormula = "{race.racestatus} = 'R'"
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr6a
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

        'Report1.SelectionFormula = Formula$
    'End If
    
    Report1.Action = 1
    DoEvents 'allow report to print
    Report1.SortFields(0) = ""
    Report1.SelectionFormula = ""
    Exit Sub

PrinterErr6a:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub
End Sub

Sub CmdPrtResults_Click ()
Dim Msg As String
Dim NL As String
NL$ = Chr(13)

If MsgPanelBox.Visible = True Then
    MsgPanelBox.Visible = False
Else
    Msg$ = "Run this report from the RACE MODULE." & NL$
    Msg$ = Msg$ & " The Race Result report is available on the RACE RESULT ENTRY screen." & NL$
    Msg$ = Msg$ & " Select Race from Race List Screen, Select Results Entry from the Race detail screen."
    Msg$ = Msg$ & " Click on the print button to get results of all birds sent." & NL$
    MsgPanelBox.Caption = Msg$
    MsgPanelBox.Visible = True
End If

End Sub

Sub CmdPrtStock_Click ()
    Report1.ReportFileName = DBPATH$ & "GenStock.rpt"
    
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr1c
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

PrinterErr1c:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
   
End Sub

Sub CmdPrtTeam_Click ()
FMPcodes.task.Text = "T"
FMPcodes.Show 1
End Sub

Sub CmdPrtTeams_Click ()
FMRteam.Show 1
End Sub

Sub CmdPrtTosses_Click ()
    Report1.ReportFileName = DBPATH$ & "racelst1.rpt"
    Report1.SortFields(0) = "-{race.reldate}"
    Report1.SelectionFormula = "{race.racestatus} = 'T'"
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr6b
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

        'Report1.SelectionFormula = Formula$
    'End If
    
    Report1.Action = 1
    DoEvents 'allow report to print
    Report1.SortFields(0) = ""
    Report1.SelectionFormula = ""
    Exit Sub

PrinterErr6b:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

Sub cmdRCard_Click ()
screen.MousePointer = 11
frmRptCard.Show 1
End Sub

Sub cmdSaleRpt_Click ()
frmRptSale.Show 1
End Sub

Sub Panel3D1_MouseMove (button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanelBox.Visible = False
End Sub

