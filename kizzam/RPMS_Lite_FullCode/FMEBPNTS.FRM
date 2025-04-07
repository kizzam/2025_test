VERSION 2.00
Begin Form FMEbpnts 
   Caption         =   "Report - Bird Points Flown"
   ClientHeight    =   5280
   ClientLeft      =   1500
   ClientTop       =   1395
   ClientWidth     =   6960
   Height          =   5685
   Left            =   1440
   LinkTopic       =   "Form4"
   ScaleHeight     =   5280
   ScaleWidth      =   6960
   Top             =   1050
   Width           =   7080
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4920
      Width           =   6975
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6975
      Begin SSCommand ButtDisplayReport 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "Display Report"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1200
         RoundedCorners  =   0   'False
         TabIndex        =   8
         Top             =   4320
         Width           =   2175
      End
      Begin SSCommand buttcancel 
         AutoSize        =   1  'Adjust Picture Size To Button
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3720
         RoundedCorners  =   0   'False
         TabIndex        =   0
         Top             =   4320
         Width           =   2175
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Select Birds to List for selection"
         Font3D          =   1  'Raised w/light shading
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   600
         ShadowStyle     =   1  'Raised
         TabIndex        =   12
         Top             =   120
         Width           =   4575
         Begin SSOption Option3D6 
            Caption         =   "Stock birds"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   840
            Width           =   1575
         End
         Begin SSOption Option3D4 
            Caption         =   "Race birds"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   1815
         End
         Begin SSOption Option3D5 
            Caption         =   "All birds on file"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
         Begin SSOption Option3D3 
            Caption         =   "Reference birds"
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            Top             =   840
            Width           =   1695
         End
         Begin SSOption Option3D2 
            Caption         =   "Current Stock birds"
            Height          =   255
            Left            =   2520
            TabIndex        =   5
            Top             =   600
            Width           =   1935
         End
         Begin SSOption Option3D1 
            Caption         =   "Current Race birds"
            Height          =   255
            Left            =   2520
            TabIndex        =   4
            Top             =   360
            Width           =   1935
         End
      End
      Begin CSVList CSVList1 
         AutoRecordCount =   -1  'True
         BorderStyle     =   1  'Fixed Single
         ColDelim        =   "	"
         ColWidth        =   FMEBPNTS.FRX:0000
         Height          =   2295
         HScroll         =   0   'False
         Left            =   240
         ListCount       =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         MultiSelect     =   0  'None
         ShadowColor     =   &H00808080&
         TabIndex        =   7
         Top             =   1800
         ValueCol        =   -1
         Width           =   6495
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Race birds"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1440
         Width           =   3135
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    Unload FMEbpnts
End Sub

Sub ButtDisplayReport_Click ()
    Load FMpfbb
    FMpfbb.Show 1
End Sub

