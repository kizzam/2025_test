VERSION 2.00
Begin Form frmAbtClub 
   Caption         =   "About rpSpeed Club Management Software"
   ClientHeight    =   4200
   ClientLeft      =   1260
   ClientTop       =   1830
   ClientWidth     =   7305
   Height          =   4605
   Left            =   1200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   7305
   Top             =   1485
   Width           =   7425
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin SSPanel Panel3D4 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "'rpClub is a separate software package that focuses on meeting Club and/or Federation requirements. FEATURES: Distance calculations, Member database, Race calculations, Certificate production from results etc."
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   6135
      End
      Begin SSPanel Panel3D3 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "About ""rpClub"" Club Management system"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   6135
      End
      Begin SSPanel Panel3D2 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "'rpClub' is a software system for managing Federations and Clubs. For more information contact and ask for more information on the Club management system."
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   600
         TabIndex        =   2
         Top             =   2160
         Width           =   6135
      End
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2880
         TabIndex        =   1
         Top             =   3360
         Width           =   1575
      End
   End
End
Option Explicit

Sub cmdExit_Click ()
Unload Me
End Sub

