VERSION 2.00
Begin Form FMPcodes 
   Caption         =   "Print Codes"
   ClientHeight    =   3720
   ClientLeft      =   1665
   ClientTop       =   2130
   ClientWidth     =   7365
   Height          =   4125
   Left            =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7365
   Top             =   1785
   Width           =   7485
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7335
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
         Height          =   615
         Left            =   3720
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin SSCommand ButtOK 
         Caption         =   "Do Report"
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
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2400
         Width           =   3615
      End
      Begin TextBox task 
         Height          =   285
         Left            =   240
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
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
         ReportSource    =   0  'Report File
         ReportTitle     =   ""
         SelectionFormula=   ""
         SessionHandle   =   0
         SQLQuery        =   ""
         Top             =   1200
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
         Top             =   360
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Output to"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   1575
         Left            =   1800
         ShadowStyle     =   1  'Raised
         TabIndex        =   6
         Top             =   720
         Width           =   3615
         Begin SSOption ToPrinter 
            Caption         =   "to Printer"
            Height          =   255
            Left            =   1200
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   960
            Width           =   1095
         End
         Begin SSOption ToScreen 
            Caption         =   "to Screen"
            Height          =   255
            Left            =   1200
            TabIndex        =   0
            Top             =   600
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
Unload FMPcodes
End Sub

Sub ButtOK_Click ()
    
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

Sub Form_Activate ()
    Select Case task.Text
        Case Is = "C"' colour
            Label1.Caption = "COLOUR CODES"
            Report1.ReportFileName = DBPATH$ & "colour.rpt"
        Case Is = "E" ' eyesign
            Label1.Caption = "EYESIGN CODES"
            Report1.ReportFileName = DBPATH$ & "eyesign.rpt"
        Case Is = "R" ' race points
            Label1.Caption = "RACE POINT CODES"
            Report1.ReportFileName = DBPATH$ & "point.rpt"
        Case Is = "A" ' Cages
            Label1.Caption = "CAGE CODES"
            Report1.ReportFileName = DBPATH$ & "cage.rpt"
        Case Is = "B" 'Bloodlines
            Label1.Caption = "BLOODLINE CODES"
            Report1.ReportFileName = DBPATH$ & "bloodlin.rpt"
        Case Is = "T" 'Team Codes
            Label1.Caption = "TEAM CODES"
            Report1.ReportFileName = DBPATH$ & "teamcode.rpt"
        Case Else  'Ratings
            Label1.Caption = "RATINGS CODES"
            Report1.ReportFileName = DBPATH$ & "ratings.rpt"
        End Select
        Text1.Text = Report1.ReportFileName
End Sub

Sub ToPrinter_Click (Value As Integer)
    Report1.Destination = 1
End Sub

Sub ToScreen_Click (Value As Integer)
    Report1.Destination = 0
End Sub

