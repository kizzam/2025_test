VERSION 2.00
Begin Form FMConvMenu 
   Caption         =   "RP Conversions"
   ClientHeight    =   2745
   ClientLeft      =   2220
   ClientTop       =   3240
   ClientWidth     =   6315
   Height          =   3255
   Left            =   2160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   6315
   Top             =   2790
   Width           =   6435
   Begin SSCommand butt3UpdateRPExtra 
      Caption         =   "Update rpEXTRAMDB"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin SSCommand Command3D4 
      Caption         =   "Convert Dates"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin SSCommand Command3D3 
      Caption         =   "Test Team Birds?"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin SSCommand Command3D2 
      Caption         =   "dBase to Access"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin SSCommand Command3D1 
      Caption         =   "&Exit"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Option Explicit

Sub Command3D1_Click ()
Unload FMconvMenu
End Sub

Sub Command3D2_Click ()
CoForm1.Show 1
End Sub

Sub Command3D3_Click ()
test.Show 1
End Sub

Sub Command3D4_Click ()
CoDOB.Show 1

End Sub

