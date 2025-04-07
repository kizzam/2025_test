VERSION 2.00
Begin Form FMRbird 
   Caption         =   "Report (Breeding Register)"
   ClientHeight    =   3810
   ClientLeft      =   1080
   ClientTop       =   1590
   ClientWidth     =   7815
   Height          =   4215
   Left            =   1020
   LinkTopic       =   "Form7"
   ScaleHeight     =   3810
   ScaleWidth      =   7815
   Top             =   1245
   Width           =   7935
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7815
      Begin SSCommand ButtPrint 
         Caption         =   "&Print"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   3120
         Width           =   1095
      End
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
         Height          =   495
         Left            =   4080
         TabIndex        =   10
         Top             =   3120
         Width           =   1095
      End
      Begin ComboBox Combo1 
         Height          =   300
         Left            =   3480
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   840
         Width           =   975
      End
      Begin SSFrame Frame3D3 
         Caption         =   "Select Birds to be included in Report by placing 'X' in box"
         Font3D          =   1  'Raised w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1455
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   12
         Top             =   1440
         Width           =   6375
         Begin SSCheck Check3D2 
            Caption         =   "Reference"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   1215
         End
         Begin SSCheck Check3D3 
            Caption         =   "Breeders Plate (A)"
            Height          =   195
            Left            =   2040
            TabIndex        =   5
            Top             =   360
            Width           =   1815
         End
         Begin SSCheck Check3D1 
            Caption         =   "Breeders Plate (B)"
            Height          =   195
            Left            =   2040
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin SSCheck Check3D4 
            Caption         =   "Special 1"
            Height          =   195
            Left            =   2040
            TabIndex        =   7
            Top             =   840
            Width           =   1095
         End
         Begin SSCheck Check3D5 
            Caption         =   "Special 2"
            Height          =   195
            Left            =   2040
            TabIndex        =   8
            Top             =   1080
            Width           =   1095
         End
         Begin SSCheck Check3D7 
            Caption         =   "Current"
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   855
         End
         Begin SSCheck Check3D9 
            Caption         =   "Race Bird"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1095
         End
         Begin SSCheck Check3D10 
            Caption         =   "Stock Bird"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1215
         End
      End
      Begin Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year : "
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECTION CRITERIA"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    Unload FMRbird
End Sub

