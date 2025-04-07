VERSION 2.00
Begin Form FMcolour 
   Caption         =   "Bird Colour Maintenance/Enquiry"
   ClientHeight    =   4650
   ClientLeft      =   1905
   ClientTop       =   2010
   ClientWidth     =   7365
   Height          =   5055
   Left            =   1845
   LinkTopic       =   "Form4"
   ScaleHeight     =   4650
   ScaleWidth      =   7365
   Top             =   1665
   Width           =   7485
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   7815
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin TrueGrid FMcolour 
         AllowArrows     =   -1  'True
         AllowTabs       =   -1  'True
         Editable        =   -1  'True
         EditDropDown    =   -1  'True
         ExposeCellMode  =   0  'Expose upon selection
         FetchMode       =   0  'By cell
         HeadingHeight   =   1
         Height          =   2535
         HorzLines       =   0  'None
         Layout          =   FMCOLOUR.FRX:0000
         LayoutIndex     =   1
         Left            =   360
         LinesPerRow     =   1
         MarqueeUnique   =   -1  'True
         SplitPropsGlobal=   -1  'True
         SplitTabMode    =   0  'Don't tab across splits
         TabCapture      =   0   'False
         TabIndex        =   1
         Top             =   840
         UseBookmarks    =   -1  'True
         Width           =   5295
         WrapCellPointer =   0   'False
      End
      Begin CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   6120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin CommandButton Command3 
         Caption         =   "&Update"
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin CommandButton Command5 
         Caption         =   "&Add"
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   2880
         Width           =   975
      End
      Begin CommandButton Command6 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   3720
         Width           =   975
      End
   End
End
Option Explicit

Sub Command1_Click ()
    Unload FMcolour
End Sub

Sub Command2_Click ()
Unload FMcolour
End Sub

