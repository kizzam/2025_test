VERSION 2.00
Begin Form frmCopy 
   Caption         =   "Copy Files"
   ClientHeight    =   2610
   ClientLeft      =   1290
   ClientTop       =   2610
   ClientWidth     =   4875
   Height          =   3015
   Left            =   1230
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   4875
   Top             =   2265
   Width           =   4995
   Begin Timer Timer1 
      Left            =   720
      Top             =   1680
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Phase 1 - Copy Files....Please Wait...."
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Option Explicit

Sub Form_Load ()
    Timer1.Interval = 20000  ' Set Timer interval.
    Screen.MousePointer = 11
    FileCopy SourcePath$ & "\rpms.32z ", DestPath$ & "rpms.exe"
    DoEvents
End Sub

'
Private Sub Timer1_Timer ()
    'If Timer.Interval > 20000 Then
        Unload Me
    'End If
End Sub

