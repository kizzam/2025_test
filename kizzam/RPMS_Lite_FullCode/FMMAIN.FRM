VERSION 2.00
Begin Form FMmain 
   Caption         =   "The Racing Pigeon Management System"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   945
   ClientWidth     =   11670
   ClipControls    =   0   'False
   Height          =   9255
   Icon            =   FMMAIN.FRX:0000
   Left            =   60
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11670
   Top             =   195
   Width           =   11790
   Begin TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   20
      Top             =   4200
      Width           =   9015
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   8055
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11655
      Begin SSPanel Panel3D3 
         BackColor       =   &H00C0E0FF&
         BevelInner      =   2  'Raised
         Caption         =   "NEW RPMS for Windows 8 && 10 Available NOW. All your data will be fully compatible/transferrable."
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   11175
      End
      Begin CommonDialog CMDialog1 
         Left            =   10080
         Top             =   1320
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
         Left            =   9240
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
         Top             =   1320
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
         Caption         =   "Modules"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   2655
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   19
         Top             =   4920
         Width           =   10575
         Begin SSCommand CmdLstTeams 
            Caption         =   "&Teams"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   3840
            TabIndex        =   3
            Top             =   600
            Width           =   2895
         End
         Begin SSCommand Butt3DBreeding 
            Caption         =   "&Breeding"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   360
            TabIndex        =   1
            Top             =   600
            Width           =   2895
         End
         Begin SSCommand Butt3DRacing 
            Caption         =   "&Racing"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   855
            Left            =   360
            TabIndex        =   2
            Top             =   1560
            Width           =   2895
         End
         Begin SSCommand Butt3DExit 
            Caption         =   "E&xit"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   855
            Left            =   7200
            TabIndex        =   0
            Top             =   1560
            Width           =   2895
         End
         Begin SSCommand Butt3DReports 
            Caption         =   "Rep&orts"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   855
            Left            =   7200
            TabIndex        =   5
            Top             =   600
            Width           =   2895
         End
         Begin SSCommand Butt3Contacts 
            Caption         =   "&Contacts"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   3840
            TabIndex        =   4
            Top             =   1560
            Width           =   2895
         End
      End
      Begin SSPanel Panel3D4 
         BackColor       =   &H00C0E0FF&
         BevelInner      =   2  'Raised
         Caption         =   "The Racing Pigeon Management System (Lite)"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Times New Roman"
         FontSize        =   24
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   975
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   11175
      End
      Begin SSPanel Panel3D1 
         BackColor       =   &H00C0E0FF&
         BevelInner      =   2  'Raised
         Caption         =   "Kieran Malone (Copyright 1990-2016) All Rights Reserved"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   2280
         TabIndex        =   15
         Top             =   1680
         Width           =   6615
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Software Licensed to"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Microsoft Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   2535
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   8
         Top             =   2160
         Width           =   10575
         Begin TextBox Text8 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5400
            TabIndex        =   18
            Top             =   1560
            Width           =   4215
         End
         Begin TextBox Text7 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5400
            TabIndex        =   17
            Top             =   1200
            Width           =   4215
         End
         Begin TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   16
            Top             =   840
            Width           =   4215
         End
         Begin TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   13
            Top             =   480
            Width           =   4215
         End
         Begin TextBox Text3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   12
            Top             =   1200
            Width           =   4215
         End
         Begin TextBox Text4 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   11
            Top             =   1560
            Width           =   4215
         End
         Begin TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5400
            TabIndex        =   10
            Top             =   480
            Width           =   4215
         End
         Begin TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5400
            TabIndex        =   9
            Top             =   840
            Width           =   4215
         End
      End
      Begin Image Image4 
         Height          =   15
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin SSPanel MsgPanel 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   8040
      Width           =   11655
   End
   Begin Menu mnuLoft 
      Caption         =   "&File"
      Begin Menu mnuOwner 
         Caption         =   "&Owner"
      End
      Begin Menu mnuRacePoints 
         Caption         =   "&Race Points"
      End
      Begin Menu mnuColour 
         Caption         =   "&Colour"
      End
      Begin Menu mnuBlines 
         Caption         =   "&Bloodlines"
      End
      Begin Menu mnuEye 
         Caption         =   "&Eye Sign"
      End
      Begin Menu mnuBage 
         Caption         =   "Ca&ges"
      End
      Begin Menu mnuSC 
         Caption         =   "&Future Use"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin Menu mnuTeam1 
         Caption         =   "&Team Codes"
      End
      Begin Menu mnuRatings 
         Caption         =   "&Ratings"
      End
      Begin Menu mnuNotes 
         Caption         =   "&Notes"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin Menu e 
         Caption         =   "-"
      End
      Begin Menu mnuCompress 
         Caption         =   "Co&mpress"
      End
      Begin Menu mnuRepair 
         Caption         =   "R&epair Database"
      End
      Begin Menu FileBackup 
         Caption         =   "&Backup Data"
      End
      Begin Menu d 
         Caption         =   "-"
      End
      Begin Menu mnuRegSoftware 
         Caption         =   "Register &Software"
      End
      Begin Menu f 
         Caption         =   "-"
      End
      Begin Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin Menu mnuContact 
      Caption         =   "&Contacts"
      Begin Menu mnuContacts 
         Caption         =   "&Contacts"
      End
   End
   Begin Menu mnuBreed 
      Caption         =   "&Breeding"
      Begin Menu mnuBRMaint 
         Caption         =   "&Breeding"
      End
      Begin Menu mnuMatings 
         Caption         =   "&Matings"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin Menu mnuRace 
      Caption         =   "&Racing"
      Begin Menu mnuRrace 
         Caption         =   "&Race"
      End
      Begin Menu mnuCalc 
         Caption         =   "&Calculator"
      End
   End
   Begin Menu mnuReports 
      Caption         =   "R&eports"
      Begin Menu mnugr 
         Caption         =   "&General reports"
      End
      Begin Menu mnub 
         Caption         =   "-"
      End
      Begin Menu mnuERacePoints 
         Caption         =   "&Race Points"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin Menu mnuRcol 
         Caption         =   "&Colours"
      End
      Begin Menu mnuCages 
         Caption         =   "C&ages"
      End
      Begin Menu mnuEyesign 
         Caption         =   "&Eyesign"
      End
      Begin Menu mnudash1 
         Caption         =   "-"
      End
      Begin Menu mnuRBird 
         Caption         =   "&Bird Register"
      End
      Begin Menu mnuTeam 
         Caption         =   "&Team Reports"
      End
      Begin Menu mnuSaleCat 
         Caption         =   "&Sale Catalogue"
      End
      Begin Menu mnuCSB 
         Caption         =   "C&urrent Stock Bird List"
      End
      Begin Menu mnuRaceBirds 
         Caption         =   "C&urrent Race Bird list"
      End
      Begin Menu b 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin Menu RERaceF 
         Caption         =   "B&ird Map"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin Menu c 
         Caption         =   "-"
      End
      Begin Menu mnuContConcise 
         Caption         =   "Co&ntacts - Concise detail"
      End
      Begin Menu mnuContFull 
         Caption         =   "Contacts - &Full detail"
      End
   End
   Begin Menu mnuHelp 
      Caption         =   "&Help"
      Begin Menu mnuHHelp 
         Caption         =   "&Help Backups"
      End
      Begin Menu mnuHOrder 
         Caption         =   "Print &Order for Software"
      End
      Begin Menu mnuSys 
         Caption         =   "Print &System details"
      End
      Begin Menu mnuClub 
         Caption         =   "About &rpClub Club System"
      End
      Begin Menu mnuAbout 
         Caption         =   "&About RPMS"
      End
   End
End
Option Explicit

Sub Butt3Contacts_Click ()
mnuContacts_Click
End Sub

Sub Butt3Contacts_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Click for Contact details"
End Sub

Sub Butt3DBreeding_Click ()
mnuBRMaint_Click
End Sub

Sub Butt3DBreeding_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Click to start Breeding module"
End Sub

Sub Butt3DExit_Click ()
Dim reply As Integer
Dim r As String

'Check of last backup over 21 days if so prompt
If Now > DateAdd("d", 30, GLLastBU) Then
    reply% = MsgBox("Do you wish to do backup before shutting down?", 36, "Your last backup was over 30 days ago")
    If reply% = 6 Then
	r$ = DB_Backup()
    End If
End If
Unload FMMain
End
End Sub

Sub Butt3DExit_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Click to Shutdown or Exit the RPMS program"
End Sub

Sub Butt3DRacing_Click ()
    Butt3DRacing.Enabled = False
    FMracepmt.Show 1
    Butt3DRacing.Enabled = True
    screen.MousePointer = 0
End Sub

Sub Butt3DRacing_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Click to work with Racing records"
End Sub

Sub Butt3DReports_Click ()
'mnuRBird_Click
FMRmaster.Show 1
End Sub

Sub Butt3DReports_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Click for Reporting functions"
End Sub

Sub CmdLstTeams_Click ()
FMlstTeam.Show
End Sub

Sub CmdLstTeams_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Click to enter TEAM module"
End Sub

Sub FileBackup_Click ()
    Dim Y As String
    Y$ = DB_Backup()
End Sub

Sub Form_Activate ()
    SetupScreen
End Sub

Sub Frame3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Integer
Randomize   ' Seed random number generator.
x1% = Int(4 * Rnd + 1)    ' Generate first die value.

Select Case x1%
    Case 1
	msgPanel.Caption = "These are the details of the licenced user"
    Case 2
	msgPanel.Caption = "Winners are Grinners!"
    Case 3
	msgPanel.Caption = "Each copy of this software requires a separate licence"
    Case Else
	msgPanel.Caption = GLOwner$ & " - Striving for the top"
End Select

End Sub

Sub Image4_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Townsville (1160 Klms) Winner 1990 - 1st WZIC Club, 7th North Section, 16th QPF Federation"
End Sub

Sub mnuAbout_Click ()
    FMabout.Show
End Sub

Sub mnuBage_Click ()
    FMcage.Show 1
    'Load FMaddcage
End Sub

Sub mnuBlines_Click ()
FMBline.Show 1
End Sub

Sub mnuBRMaint_Click ()
    
    'Load prompt form.
    'Load FMbrreg

    'Advise it that we are in maintenance mode.
    FMbrreg.Task.Text = "M"

    FMbrreg.Show 1
    screen.MousePointer = 0

End Sub

Sub mnuCages_Click ()
FMPcodes.task.Text = "A"
FMPcodes.Show 1
End Sub

Sub mnuCalc_Click ()
    frmCalc.Show 1
End Sub

Sub mnuClub_Click ()
frmAbtClub.Show 1
End Sub

Sub mnuColour_Click ()
    FMcolor.Show 1
    'Load FMaddcolour
End Sub

Sub mnuCompress_Click ()
   Dim reply As Integer
   reply% = DB_Repair("COMPRESS")
    If reply% = -1 Then
	MsgBox "Attempted COMPRESS appears successful", 16, "COMPRESS completed"
    Else
	MsgBox "COMPRESS not successful", 16, "WARNING"
    End If
   screen.MousePointer = 0
End Sub

Sub mnuContacts_Click ()
FMlstcontact.Show
screen.MousePointer = 0
End Sub

Sub mnuContConcise_Click ()
    
    Report1.ReportFileName = DBPath$ & "conshort.rpt"
    
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

Sub mnuContFull_Click ()
    
    Report1.ReportFileName = DBPath$ & "con50916.rpt"
    
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

Sub mnuCSB_Click ()
    Report1.ReportFileName = DBPath$ & "GenStock.rpt"
    
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
	On Error GoTo PrinterErr1a
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

PrinterErr1a:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
   
End Sub

Sub mnuERacePoints_Click ()
FMPcodes.task.Text = "R"
FMPcodes.Show 1
End Sub

Sub mnuExit_Click ()
    Butt3DExit_Click
End Sub

Sub mnuEye_Click ()
    FMeyesgn.Show 1
    'Load FMaddeyesgn
End Sub

Sub mnuEyesign_Click ()
FMPcodes.task.Text = "E"
FMPcodes.Show 1
End Sub

Sub mnugr_Click ()
FMRmaster.Show 1
End Sub

Sub mnuHHelp_Click ()
frmHelpBU.Show
End Sub

Sub mnuHOrder_Click ()
PrtOrder
End Sub

Sub mnuOwner_Click ()
    'FMowner.task.Text = "S"  'Allow Master Changes
    FMowner.task.Text = "N"  'Allow Reduced Changes
    FMowner.Show 1
End Sub

Sub mnuRaceBirds_Click ()
    Report1.ReportFileName = DBPath$ & "GenRace.rpt"
    
    ' Show printer setup dialog box if destination is printer.
    If Report1.Destination = 1 Then
	On Error GoTo PrinterErr1b
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

PrinterErr1b:
    MsgBox Error$, 16, "Report Printing Error"
    Exit Sub
    

End Sub

Sub mnuRacePoints_Click ()
    FMracepnt.Show 1
    'Load FMaddracepnt
End Sub

Sub mnuRatings_Click ()
FMLstRatings.Show 1
End Sub

Sub mnuRBird_Click ()
FMBRbird.Show 1
End Sub

Sub mnuRcol_Click ()
FMPcodes.task.Text = "C"
FMPcodes.Show 1
End Sub

Sub mnuRegSoftware_Click ()
    frmRegister.Show 1
    GetLoftDets
    SetupScreen
End Sub

Sub mnuRepair_Click ()
Dim reply As Integer
reply% = DB_Repair("REPAIR")
    If reply% = -1 Then
	MsgBox "Attempted repair appears successful", 16, "Repair completed"
    Else
	MsgBox "Repair not completed", 16, "Warning"
    End If
End Sub

Sub mnuRRace_Click ()
    Butt3DRacing_Click
End Sub

Sub mnuSaleCat_Click ()
frmRptSale.Show 1
End Sub

Sub mnuSC_Click ()
    FMspecial.Show 1 'bloodline
End Sub

Sub mnuSys_Click ()
msgPanel.Caption = "PLEASE WAIT while System details printed"
prtsys
msgPanel.Caption = ""
End Sub

Sub mnuTeam_Click ()
FMRteam.Show 1
End Sub

Sub mnuTeam1_Click ()
    FMteam.Show 1
End Sub

Sub MsgPanel_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Kieran Malone Copyright 1988-2001"
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = "Copyright K G Malone - All rights reserved"
End Sub

Sub Panel3D2_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
msgPanel.Caption = ""
End Sub

Sub Panel3D4_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xX As Integer
    Randomize   ' Seed random number generator.
    xX% = Int(6 * Rnd + 1)    ' Generate first die value.
    Select Case xX%
	Case 1
	    msgPanel.Caption = " The Software used by the Masters"
	Case 2
	    msgPanel.Caption = " Professional software to help you manage"
	Case 3
	    msgPanel.Caption = " Pedigrees at the 'Click' of a button"
	Case 4
	    msgPanel.Caption = " Winners are Grinners"
	Case 5
	    msgPanel.Caption = " RPMS assists you in your attempt to G R I N"
	Case 6
	    msgPanel.Caption = " Information is Power"
    End Select
End Sub

Sub SetupScreen ()
    Text1.Text = GLOwner$
    Text2.Text = GLAddress1$
    Text3.Text = GLAddress2$
    Text4.Text = GLCity$ & " " & GLState$ & " " & GLPostcode$
    Text5.Text = "Phone: " & GLPhoneno$
    Text6.Text = "Loft Name: " & GLLoft$
    Text7.Text = "Club/Fed Mark : " & GLMark$
    If GLOwner$ = "Demonstration Copy" Then
	Text8.Text = "Serial No. : " & GLSerialNo$
	FMMain.mnuRegSoftware.Enabled = True
	FMMain.mnuRegSoftware.Visible = True
	FMMain.f.Visible = True
    Else
	Text8.Text = "Serial No. : " & LicenseDisplay(GLSerialNo$)
	FMMain.mnuRegSoftware.Enabled = False
	FMMain.mnuRegSoftware.Visible = False
	FMMain.f.Visible = False
    End If
    text9.Text = "Loft Location:" & GLLoftLat & " Degs Latitude, " & GLLoftLng & " Degs Longitude"
End Sub

