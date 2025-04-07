VERSION 2.00
Begin Form FMloftmgr 
   Caption         =   "Loft Manager"
   ClientHeight    =   5385
   ClientLeft      =   660
   ClientTop       =   1110
   ClientWidth     =   8400
   Height          =   5790
   Left            =   600
   LinkTopic       =   "Form2"
   ScaleHeight     =   5385
   ScaleWidth      =   8400
   Top             =   765
   Width           =   8520
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   8415
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin SSFrame Frame3D3 
         Caption         =   "Medication"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   1095
         Left            =   5640
         ShadowStyle     =   1  'Raised
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Owner"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   1335
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   8
         Top             =   720
         Width           =   2415
         Begin SSCommand Command3D1 
            Caption         =   "Command3D1"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Code Files"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   2295
         Left            =   840
         ShadowStyle     =   1  'Raised
         TabIndex        =   4
         Top             =   2280
         Width           =   2415
         Begin SSCommand Command3D6 
            Caption         =   "Command3D6"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   1800
            Width           =   1335
         End
         Begin SSCommand Command3D5 
            Caption         =   "Command3D5"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   480
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin SSCommand Command3D4 
            Caption         =   "Command3D4"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   480
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
      End
      Begin SSCommand Command3D3 
         Caption         =   "Command3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   3600
         TabIndex        =   3
         Top             =   3240
         Width           =   1575
      End
      Begin SSCommand Command3D2 
         Caption         =   "Command3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   3600
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
   End
End
Option Explicit

