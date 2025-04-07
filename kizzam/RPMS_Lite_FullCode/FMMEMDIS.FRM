VERSION 2.00
Begin Form FMmemdist 
   Caption         =   "Members Distances from Race Points"
   ClientHeight    =   4020
   ClientLeft      =   1770
   ClientTop       =   1425
   ClientWidth     =   7365
   Height          =   4425
   Left            =   1710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1080
   Width           =   7485
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'I~set
      BevelOuTer      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   7335
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      Font3D          =   0  'None
      ForeColor       =   &H00808080&
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "D:\DEV_VB\CLUB\CLUB.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "distance"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin CommandButton ButtUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   6120
         TabIndex        =   9
         Top             =   2640
         Width           =   855
      End
      Begin CommandButton ButtDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   2040
         Width           =   855
      End
      Begin CommandButton ButtAdd 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin CommandButton ButtCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin TrueGrid Table1 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         DataSource      =   "Data1"
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   2055
         HorzLines       =   0  'None
         Layout          =   FMMEMDIS.FRX:0000
         LayoutIndex     =   1
         Left            =   240
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   5
         Top             =   960
         UseBookmarks    =   -1  'True
         Width           =   5535
         WrapCellPointer =   0   'False
      End
      Begin TextBox lastname 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
      Begin TextBox code_mem 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Member : "
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
End
Option Explicit

Sub ButtCancel_Click (Index As Integer)
    Unload FMmemdist
End Sub

