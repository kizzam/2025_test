VERSION 2.00
Begin Form FMmain1 
   Caption         =   "The Racing Pigeon Management System"
   ClientHeight    =   6600
   ClientLeft      =   90
   ClientTop       =   660
   ClientWidth     =   9375
   ClipControls    =   0   'False
   Height          =   7290
   Left            =   30
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9375
   Top             =   30
   Width           =   9495
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9615
      Begin CommonDialog CMDialog1 
         Left            =   8520
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
         Left            =   8520
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
         Top             =   2400
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
         ForeColor       =   &H00008000&
         Height          =   2295
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   17
         Top             =   3720
         Width           =   8655
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
            Height          =   735
            Left            =   3120
            TabIndex        =   2
            Top             =   480
            Width           =   2535
         End
         Begin SSCommand Butt3DBreeding 
            Caption         =   "Br&eeding"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   1575
            Left            =   360
            TabIndex        =   1
            Top             =   480
            Width           =   2535
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
            Height          =   735
            Left            =   5880
            TabIndex        =   18
            Top             =   1320
            Width           =   2535
         End
         Begin SSCommand Butt3DReports 
            Caption         =   "Re&ports"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   735
            Left            =   5880
            TabIndex        =   4
            Top             =   480
            Width           =   2535
         End
         Begin SSCommand Butt3Contacts 
            Caption         =   "Contacts"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   3120
            TabIndex        =   3
            Top             =   1320
            Width           =   2535
         End
      End
      Begin SSPanel Panel3D4 
         BevelInner      =   2  'Raised
         Caption         =   "The Bird Stud"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Times New Roman"
         FontSize        =   22.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   720
         TabIndex        =   12
         Top             =   120
         Width           =   7815
      End
      Begin SSPanel Panel3D1 
         BevelInner      =   2  'Raised
         Caption         =   "Kieran Malone (Copyright 1990-2001) All Rights Reserved"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   960
         Width           =   5415
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Software Licensed to"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Times New Roman"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   1935
         Left            =   1320
         ShadowStyle     =   1  'Raised
         TabIndex        =   6
         Top             =   1560
         Width           =   6375
         Begin TextBox Text8 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            TabIndex        =   16
            Top             =   1440
            Width           =   2775
         End
         Begin TextBox Text7 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            TabIndex        =   15
            Top             =   1080
            Width           =   2775
         End
         Begin TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   2775
         End
         Begin TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2775
         End
         Begin TextBox Text3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   2775
         End
         Begin TextBox Text4 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   2775
         End
         Begin TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            TabIndex        =   8
            Top             =   360
            Width           =   2775
         End
         Begin TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            TabIndex        =   7
            Top             =   720
            Width           =   2775
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
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   9615
   End
   Begin Menu mnuLoft 
      Caption         =   "&File"
      Begin Menu mnuOwner 
         Caption         =   "&Owner"
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
   Begin Menu mnuReports 
      Caption         =   "R&eports"
      Begin Menu mnugr 
         Caption         =   "&General reports"
      End
      Begin Menu mnub 
         Caption         =   "-"
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
      Begin Menu mnuRPedg 
         Caption         =   "&Sale Catalogue"
         Enabled         =   0   'False
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
      Begin Menu mnuREPFlown 
         Caption         =   "Ra&ces flown by birds"
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
         Caption         =   "&Order Software"
      End
      Begin Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Option Explicit

Sub Butt3Contacts_Click ()
mnuContacts_Click
End Sub

Sub Butt3Contacts_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Click for Contact details"
End Sub

Sub Butt3DBreeding_Click ()
mnuBRMaint_Click
End Sub

Sub Butt3DBreeding_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Click to start Breeding module"
End Sub

Sub Butt3DExit_Click ()
Dim reply As Integer
Dim r As String

'Check of last backup over 21 days if so prompt
If Now > DateAdd("d", 21, GLLastBU) Then
    reply% = MsgBox("Do you wish to do backup before shutting down?", 36, "Your last backup was over 21 days ago")
    If reply% = 6 Then
	r$ = DB_Backup()
    End If
End If
Unload FMmain1
End
End Sub

Sub Butt3DExit_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Click to Shutdown or Exit the RPMS program"
End Sub

Sub Butt3DReports_Click ()
'mnuRBird_Click
FMRmaster.Show 1
End Sub

Sub Butt3DReports_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Click for Reporting functions"
End Sub

Sub CmdLstTeams_Click ()
FMlstTeam.Show
End Sub

Sub CmdLstTeams_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Click to enter TEAM module"
End Sub

Sub FileBackup_Click ()
    Dim yy As String
    yy$ = DB_Backup()
End Sub

Sub Form_Load ()
    Text1.Text = GLOwner$
    Text2.Text = GLAddress1$
    Text3.Text = GLAddress2$
    Text4.Text = GLCity$ & " " & GLState$ & " " & GLPostcode$
    Text5.Text = "Phone: " & GLPhoneno$
    Text6.Text = "Loft Name: " & GLLoft$
    Text7.Text = "Club/Fed Mark : " & GLMark$
    Text8.Text = "License No.: " & LicenseDisplay(GLSerialNo$)
End Sub

Sub Frame3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
Dim x1 As Integer
Randomize   ' Seed random number generator.
x1% = Int(4 * Rnd + 1)    ' Generate first die value.

Select Case x1%
    Case 1
	MsgPanel.Caption = "These are the details of the licenced user"
    Case 2
	MsgPanel.Caption = "Winners are Grinners!"
    Case 3
	MsgPanel.Caption = "Each copy of this software requires a separate licence"
    Case Else
	MsgPanel.Caption = GLOwner$ & " - Striving for the top"
End Select

End Sub

Sub Image4_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Townsville (1160 Klms) Winner 1990 - 1st WZIC Club, 7th North Section, 16th QPF Federation"
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

Sub mnuColour_Click ()
    FMcolor.Show 1
    'Load FMaddcolour
End Sub

Sub mnuCompress_Click ()
    Dim reply As Integer

    '+++ Notify user to backup prior to Compression
    reply% = MsgBox("Have you done a backup just prior to this operation?", 36, "BACKUP COMPLETED?")
    If reply% = 6 Then 'Yes
	reply% = MsgBox("Double Check - Have you done a backup just prior to this operation?", 36, "BACKUP COMPLETED?")
	If reply% = 6 Then 'Yes
	    screen.MousePointer = 11
	    'Open Database for exclusive use
	    Set DBFile = OpenDatabase(DBFullFileName$, True, False)
	    'Close Database
	    DBFile.Close
    
	    '+++ Check enough space
	    MsgPanel.Caption = "COMPRESSING DATABASE - Please Wait . . . . ."
	    CompactDatabase DBFullFileName$, DBPath$ & "COMPRESS.MDB", DB_LANG_GENERAL, DB_ENCRYPT
	    DoEvents
	    'MsgBox "Must do a backup prior to compression", 48, "OPERATION ABORTED BY SYSTEM"
	    Kill DBFullFileName$
	    DoEvents
	    Name DBPath$ & "COMPRESS.MDB" As DBFullFileName$
	    MsgPanel.Caption = ""
	Else
	    MsgBox "Must do a backup prior to compression", 48, "OPERATION ABORTED BY SYSTEM"
	End If
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

Sub mnuRepair_Click ()
Dim reply As Integer
reply% = DB_Repair("REPAIR")
    If reply% = -1 Then
	MsgBox "Attempted repair appears successful", 16, "Repair completed"
	Resume
    Else
	MsgBox "Repair not completed", 16, "Program Ending"
    End If
End Sub

Sub mnuRPedg_Click ()
    Load FMpedg
    FMpedg.Show 1
End Sub

Sub mnuSC_Click ()
    FMspecial.Show 1 'bloodline
End Sub

Sub mnuTeam_Click ()
FMRteam.Show 1
End Sub

Sub mnuTeam1_Click ()
    FMteam.Show 1
End Sub

Sub MsgPanel_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Kieran Malone Copyright 1988-2001"
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Copyright K G Malone - All rights reserved"
End Sub

Sub Panel3D2_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = ""
End Sub

Sub Panel3D4_MouseMove (Button As Integer, Shift As Integer, X As Single, y As Single)
MsgPanel.Caption = "Professional software"
End Sub

