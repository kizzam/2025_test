VERSION 2.00
Begin Form frmRptSale 
   Caption         =   "Print Sale Catalogue"
   ClientHeight    =   4545
   ClientLeft      =   1515
   ClientTop       =   1590
   ClientWidth     =   7305
   Height          =   4950
   Left            =   1455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7305
   Top             =   1245
   Width           =   7425
   Begin SSPanel Panel3D1 
      BevelInner      =   2  'Raised
      ForeColor       =   &H00FF0000&
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin CommonDialog CMDialog1 
         Left            =   840
         Top             =   3720
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
         Top             =   3720
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
      Begin SSCommand cmdPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2160
         TabIndex        =   7
         Top             =   3600
         Width           =   1575
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Select Section of Catalogue to print"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00008000&
         Height          =   1935
         Left            =   1200
         ShadowStyle     =   1  'Raised
         TabIndex        =   3
         Top             =   1440
         Width           =   5535
         Begin SSOption OptRB 
            Caption         =   "Section 3 - Race Birds"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4935
         End
         Begin SSOption OptSB 
            Caption         =   "Section 2 - Stock Birds"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   960
            Width           =   4935
         End
         Begin SSOption OptFP 
            Caption         =   "Section 1 - Front page && Reference birds"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Value           =   -1  'True
            Width           =   5055
         End
      End
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4320
         TabIndex        =   2
         Top             =   3600
         Width           =   1575
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Note: To produce the Full Catalogue you need to run the three sections of the Catalogue. Note: Any bird marked as a Reference bird are detailed in Section 1. Any Bird marked as Current and Stock bird is in 2nd Section and any bird marked as current, race and not stock is in 3rd section."
         Height          =   1095
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdPrint_Click ()
    Dim Formula As String
    Dim Cnt1 As Integer
    Dim s As String
    Dim mClass As String
    Dim mYear As Integer
    
    Cnt1% = 0

    If OptFP Then 'Front Page
        Report1.ReportFileName = DBPATH$ & "salefp.rpt"
    ElseIf OptSB Then
        Report1.ReportFileName = DBPATH$ & "salebp.rpt"
    Else
        Report1.ReportFileName = DBPATH$ & "salerp.rpt"
    End If
    
    If Report1.Destination = 1 Then
        On Error GoTo PrinterErr6z
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

PrinterErr6z:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

