VERSION 2.00
Begin Form formx 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   1800
   ClientTop       =   1515
   ClientWidth     =   7365
   Height          =   4425
   Left            =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1170
   Width           =   7485
   Begin CommandButton Command2 
      Caption         =   "Report"
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin CommandButton Command1 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin CommonDialog CMDialog1 
      Left            =   4920
      Top             =   600
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
      Left            =   3480
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
      Top             =   720
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
End
Option Explicit

Sub Command1_Click ()
Unload formx
End Sub

Sub Command2_Click ()
    
    Report1.ReportFileName = "c:\vbnew\conshort.rpt"
    
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

