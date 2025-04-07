VERSION 2.00
Begin Form FMerror 
   Caption         =   "RPMS Error"
   ClientHeight    =   2865
   ClientLeft      =   1995
   ClientTop       =   2745
   ClientWidth     =   5775
   Height          =   3270
   Left            =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5775
   Top             =   2400
   Width           =   5895
   Begin SSPanel Panel3D1 
      BackColor       =   &H000000C0&
      BevelInner      =   1  'Inset
      Caption         =   "RPMS ERROR"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin SSCommand ButtContinue 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "&Continue"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3000
         RoundedCorners  =   0   'False
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin SSCommand ButtTerminate 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "&Terminate"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         RoundedCorners  =   0   'False
         TabIndex        =   5
         Top             =   2160
         Width           =   2175
      End
      Begin TextBox Text3 
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   1680
         Width           =   2295
      End
      Begin TextBox Text2 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1680
         Width           =   2175
      End
      Begin TextBox Text1 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin SSPanel Panel3D2 
         BevelInner      =   1  'Inset
         BevelOuter      =   1  'Inset
         Caption         =   "Error"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Option Explicit

Sub ButtContinue_Click ()
FMerror.Hide
End Sub

Sub ButtTerminate_Click ()
End
End Sub

Sub Form_Load ()
    GLLfmerror% = -1 'Flag form loaded
End Sub

Sub Form_Unload (Cancel As Integer)
    GLLfmerror% = 0 'Flag form unloaded
End Sub

