VERSION 2.00
Begin Form FMpedg 
   BackColor       =   &H00FFFF80&
   Caption         =   "Pedigree of Bird"
   ClientHeight    =   10125
   ClientLeft      =   795
   ClientTop       =   450
   ClientWidth     =   13875
   Height          =   10635
   Left            =   735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   13875
   Top             =   0
   Width           =   13995
   Begin SSPanel Panel3D5 
      BackColor       =   &H00E0E0E0&
      BevelOuter      =   0  'None
      Caption         =   "Pedgree produced by Racing Pigeon Management System RPMS Lite 2016 (Copyright Kieran Malone, All rights reserved Ph 0409495571)"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   159
      Top             =   11880
      Width           =   13935
   End
   Begin SSFrame FrmDetails 
      Caption         =   "Pedigree for"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Terminal"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   9975
      Left            =   3360
      TabIndex        =   101
      Top             =   600
      Visible         =   0   'False
      Width           =   6255
      Begin TextBox fBloodline 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   110
         Text            =   "Text2"
         Top             =   2520
         Width           =   4400
      End
      Begin TextBox fDOB 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   107
         Text            =   "Text1"
         Top             =   2040
         Width           =   4400
      End
      Begin TextBox fColour 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   840
         Width           =   4400
      End
      Begin TextBox fBredBy 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   105
         Text            =   "Text2"
         Top             =   1560
         Width           =   4400
      End
      Begin TextBox fName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   480
         Width           =   4400
      End
      Begin TextBox fDNotes 
         Enabled         =   0   'False
         Height          =   5775
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   3000
         Width           =   4935
      End
      Begin TextBox fSex 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   102
         Text            =   "fSex"
         Top             =   1200
         Width           =   4400
      End
      Begin Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bloodline:"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   2520
         Width           =   855
      End
      Begin Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "D.O.B.:"
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   2040
         Width           =   735
      End
      Begin Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Colour:"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   840
         Width           =   735
      End
      Begin Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bred by:"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   1560
         Width           =   735
      End
      Begin Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Notes:"
         Height          =   255
         Left            =   360
         TabIndex        =   87
         Top             =   3000
         Width           =   615
      End
      Begin Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   480
         Width           =   855
      End
      Begin Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   1200
         Width           =   735
      End
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      Height          =   11895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin SSCommand butt3PrintEnh_A 
         Caption         =   "Print Enhanced A"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   12000
         TabIndex        =   163
         Top             =   4440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   14
         Left            =   9720
         TabIndex        =   162
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   24
         Left            =   9720
         TabIndex        =   161
         Top             =   7560
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   23
         Left            =   9720
         TabIndex        =   160
         Top             =   6720
         Width           =   2055
      End
      Begin SSCommand butt3dPrintScrn 
         Caption         =   "Print Screen"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   12000
         TabIndex        =   158
         Top             =   5160
         Width           =   1695
      End
      Begin SSCommand butt3PrintEnh 
         Caption         =   "P&rint Enhanced"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   12000
         TabIndex        =   157
         Top             =   3720
         Width           =   1695
      End
      Begin SSCommand butt3Colours 
         Caption         =   "Show Colours"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   12240
         TabIndex        =   156
         Top             =   10440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand buttEditBird 
         Caption         =   "Edit Bird Details"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   9600
         TabIndex        =   155
         Top             =   120
         Width           =   1695
      End
      Begin CommandButton buttPedga 
         Caption         =   "Pedga"
         Height          =   855
         Left            =   12240
         TabIndex        =   154
         Top             =   9360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin TextBox fNotes 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1815
         Index           =   5
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   153
         Text            =   "Text2"
         Top             =   9960
         Width           =   2655
      End
      Begin TextBox fNotes 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1575
         Index           =   4
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   152
         Text            =   "Text2"
         Top             =   7320
         Width           =   2655
      End
      Begin TextBox fNotes 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1695
         Index           =   3
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   151
         Text            =   "Text2"
         Top             =   4440
         Width           =   2655
      End
      Begin TextBox fNotes 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1695
         Index           =   2
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   150
         Text            =   "Text2"
         Top             =   1680
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   6360
         TabIndex        =   149
         Top             =   11040
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   6360
         TabIndex        =   148
         Top             =   9720
         Width           =   2535
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   6360
         TabIndex        =   147
         Top             =   8520
         Width           =   2535
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   6360
         TabIndex        =   146
         Top             =   7080
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   6360
         TabIndex        =   145
         Top             =   5640
         Width           =   2775
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   6360
         TabIndex        =   144
         Top             =   4320
         Width           =   2775
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   6360
         TabIndex        =   143
         Top             =   2760
         Width           =   2775
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   6360
         TabIndex        =   142
         Top             =   1440
         Width           =   2775
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   29
         Left            =   9720
         TabIndex        =   141
         Text            =   "Text1"
         Top             =   11160
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   28
         Left            =   9720
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   10560
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   27
         Left            =   9720
         TabIndex        =   139
         Text            =   "Text1"
         Top             =   9840
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   26
         Left            =   9720
         TabIndex        =   138
         Text            =   "Text1"
         Top             =   9240
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   25
         Left            =   9720
         TabIndex        =   137
         Text            =   "Text1"
         Top             =   8520
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   24
         Left            =   9720
         TabIndex        =   136
         Text            =   "Text1"
         Top             =   7920
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   23
         Left            =   9720
         TabIndex        =   135
         Text            =   "Text1"
         Top             =   7080
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   22
         Left            =   9720
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   6480
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   21
         Left            =   9720
         TabIndex        =   133
         Text            =   "Text1"
         Top             =   5640
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   20
         Left            =   9720
         TabIndex        =   132
         Text            =   "Text1"
         Top             =   5040
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   19
         Left            =   9720
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   4200
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   18
         Left            =   9720
         TabIndex        =   130
         Text            =   "Text1"
         Top             =   3600
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   17
         Left            =   9720
         TabIndex        =   129
         Text            =   "Text1"
         Top             =   2760
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   16
         Left            =   9720
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   2160
         Width           =   2055
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   15
         Left            =   9720
         TabIndex        =   127
         Text            =   "Text1"
         Top             =   1440
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   28
         Left            =   9720
         TabIndex        =   126
         Top             =   10200
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   26
         Left            =   9720
         TabIndex        =   125
         Top             =   8880
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   22
         Left            =   9720
         TabIndex        =   124
         Top             =   6105
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   19
         Left            =   9720
         TabIndex        =   123
         Top             =   3840
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   29
         Left            =   9720
         TabIndex        =   121
         Top             =   10800
         Width           =   2055
         Begin SSPanel Panel3D2 
            BackColor       =   &H00E0E0E0&
            BevelOuter      =   0  'None
            Caption         =   "Panel3D5"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "MS Serif"
            FontSize        =   6
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   30
            Left            =   1080
            TabIndex        =   122
            Top             =   240
            Width           =   2055
         End
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   27
         Left            =   9720
         TabIndex        =   120
         Top             =   9480
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   25
         Left            =   9720
         TabIndex        =   119
         Top             =   8160
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   21
         Left            =   9720
         TabIndex        =   118
         Top             =   5280
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   20
         Left            =   9720
         TabIndex        =   117
         Top             =   4680
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   18
         Left            =   9720
         TabIndex        =   116
         Top             =   3240
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   17
         Left            =   9720
         TabIndex        =   115
         Top             =   2400
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   16
         Left            =   9720
         TabIndex        =   114
         Top             =   1800
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D4"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   15
         Left            =   9720
         TabIndex        =   113
         Top             =   1080
         Width           =   2055
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D4"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   150
         Index           =   14
         Left            =   9720
         TabIndex        =   112
         Top             =   480
         Width           =   2055
      End
      Begin CommandButton ButtPrintPedgN 
         Caption         =   "Print Pedigree New"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1215
         Left            =   12240
         TabIndex        =   111
         Top             =   7920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3360
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   9240
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   6360
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   2280
         Width           =   2775
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   6360
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   960
         Width           =   2775
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   6600
         Width           =   2655
      End
      Begin SSCommand ButtPrint 
         Caption         =   "Print (LS)"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   12000
         TabIndex        =   96
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin SSCommand ButtSire 
         Caption         =   "Command3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   95
         Top             =   960
         Width           =   2895
      End
      Begin SSCommand ButtDam 
         Caption         =   "Command3D1"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   94
         Top             =   6240
         Width           =   2895
      End
      Begin SSCommand cmdDetails 
         Caption         =   "&Details"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   12000
         TabIndex        =   93
         Top             =   1320
         Width           =   1695
      End
      Begin SSCommand ButtBack 
         Caption         =   "&Back"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   12000
         TabIndex        =   92
         Top             =   600
         Width           =   1695
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
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   12000
         TabIndex        =   91
         Top             =   6840
         Width           =   1695
      End
      Begin SSCommand ButtPrintN 
         Caption         =   "&Print (P)"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   12000
         TabIndex        =   90
         Top             =   3000
         Width           =   1695
      End
      Begin SSCommand Butt3Info 
         Caption         =   "&Info Off"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   12000
         TabIndex        =   83
         Top             =   2040
         Width           =   1695
      End
      Begin TextBox fNotes 
         Height          =   4455
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   7080
         Width           =   2895
      End
      Begin TextBox fNotes 
         Height          =   3735
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   1800
         Width           =   2895
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   80
         Top             =   9720
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   79
         Top             =   7080
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   78
         Top             =   4200
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   77
         Top             =   1440
         Width           =   2655
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Enabled         =   0   'False
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   76
         Top             =   6840
         Width           =   2895
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D3"
         Enabled         =   0   'False
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   75
         Top             =   1560
         Width           =   2895
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   7320
         TabIndex        =   74
         Top             =   10320
         Width           =   1695
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   7320
         TabIndex        =   73
         Top             =   9000
         Width           =   1575
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   7320
         TabIndex        =   72
         Top             =   7800
         Width           =   1935
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   7320
         TabIndex        =   71
         Top             =   6360
         Width           =   1695
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   7320
         TabIndex        =   70
         Top             =   4920
         Width           =   1815
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   7320
         TabIndex        =   69
         Top             =   3480
         Width           =   1815
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   7200
         TabIndex        =   68
         Top             =   2040
         Width           =   1695
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   7320
         TabIndex        =   67
         Top             =   720
         Width           =   1815
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   66
         Top             =   9000
         Width           =   1935
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4200
         TabIndex        =   65
         Top             =   6360
         Width           =   1815
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   64
         Top             =   3480
         Width           =   1935
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   63
         Top             =   720
         Width           =   1815
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   62
         Top             =   6000
         Width           =   2295
      End
      Begin SSPanel Panel3D2 
         BackColor       =   &H00E0E0E0&
         BevelOuter      =   0  'None
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   61
         Top             =   720
         Width           =   2535
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   13
         Left            =   6360
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   10800
         Width           =   2655
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   12
         Left            =   6360
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   9480
         Width           =   2535
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   11
         Left            =   6360
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   8280
         Width           =   2535
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   10
         Left            =   6360
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   6840
         Width           =   2655
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   9
         Left            =   6360
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   5400
         Width           =   2775
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   8
         Left            =   6360
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   4080
         Width           =   2775
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   7
         Left            =   6360
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   2520
         Width           =   2775
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   5
         Left            =   3360
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   9480
         Width           =   2655
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   6840
         Width           =   2655
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   3360
         TabIndex        =   51
         Text            =   "Text7"
         Top             =   3960
         Width           =   2655
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   6
         Left            =   6360
         TabIndex        =   50
         Text            =   "Text6"
         Top             =   1200
         Width           =   2775
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   3360
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   1200
         Width           =   2655
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   6600
         Width           =   2895
      End
      Begin TextBox Fbirdbb 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   1320
         Width           =   2895
      End
      Begin CommonDialog CMDialog1 
         Left            =   11640
         Top             =   11280
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
         Left            =   8640
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
         Top             =   11160
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
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   9720
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   10920
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   9720
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   10320
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   9720
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   9600
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   9720
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   9000
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   9720
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   8280
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   9720
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   7680
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   9720
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   6840
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   9720
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   6240
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   9720
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   5400
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   9720
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   4800
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   9720
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3960
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   9720
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3360
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   9720
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   2520
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   9720
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1920
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   9720
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1200
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   9720
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   600
         Width           =   2055
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   6360
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   10560
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   6360
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   9240
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   6360
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   8040
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   6360
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   6600
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   6360
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   5160
         Width           =   2775
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   6360
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   3720
         Width           =   2775
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   3720
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   960
         Width           =   2655
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   6240
         Width           =   2895
      End
      Begin TextBox Fbirdymr 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   960
         Width           =   2895
      End
      Begin TextBox Fyr 
         BackColor       =   &H00000080&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "Text"
         Top             =   120
         Width           =   855
      End
      Begin TextBox Fmark 
         BackColor       =   &H00000080&
         Enabled         =   0   'False
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   4560
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   120
         Width           =   1695
      End
      Begin TextBox Fring 
         BackColor       =   &H00000080&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   7440
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   120
         Width           =   2055
      End
      Begin Line Line2 
         Index           =   2
         X1              =   3360
         X2              =   3240
         Y1              =   3840
         Y2              =   3840
      End
      Begin Line Line13 
         X1              =   9120
         X2              =   9480
         Y1              =   6720
         Y2              =   6720
      End
      Begin Line Line11 
         X1              =   5880
         X2              =   6360
         Y1              =   9360
         Y2              =   9360
      End
      Begin Line Line5 
         Index           =   4
         X1              =   0
         X2              =   120
         Y1              =   0
         Y2              =   0
      End
      Begin Label fDetails 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedigree for:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   120
         Width           =   1695
      End
      Begin Line Line10 
         Index           =   1
         X1              =   3240
         X2              =   3240
         Y1              =   6720
         Y2              =   9360
      End
      Begin Line Line9 
         Index           =   1
         X1              =   3360
         X2              =   3240
         Y1              =   9360
         Y2              =   9360
      End
      Begin Line Line8 
         Index           =   1
         X1              =   3120
         X2              =   3360
         Y1              =   6720
         Y2              =   6720
      End
      Begin Line Line10 
         Index           =   0
         X1              =   3240
         X2              =   3240
         Y1              =   1080
         Y2              =   3840
      End
      Begin Line Line9 
         Index           =   0
         X1              =   3000
         X2              =   2880
         Y1              =   2880
         Y2              =   2880
      End
      Begin Line Line8 
         Index           =   0
         X1              =   3120
         X2              =   3360
         Y1              =   1080
         Y2              =   1080
      End
      Begin Line Line7 
         Index           =   3
         X1              =   9600
         X2              =   9480
         Y1              =   10440
         Y2              =   10440
      End
      Begin Line Line6 
         Index           =   3
         X1              =   6120
         X2              =   6120
         Y1              =   9360
         Y2              =   10680
      End
      Begin Line Line5 
         Index           =   3
         X1              =   6120
         X2              =   6360
         Y1              =   10680
         Y2              =   10680
      End
      Begin Line Line7 
         Index           =   2
         X1              =   6360
         X2              =   6120
         Y1              =   8160
         Y2              =   8160
      End
      Begin Line Line6 
         Index           =   2
         X1              =   6120
         X2              =   6120
         Y1              =   6720
         Y2              =   8160
      End
      Begin Line Line5 
         Index           =   2
         X1              =   6000
         X2              =   6360
         Y1              =   6720
         Y2              =   6720
      End
      Begin Line Line6 
         Index           =   1
         X1              =   6120
         X2              =   6120
         Y1              =   3840
         Y2              =   5280
      End
      Begin Line Line5 
         Index           =   1
         X1              =   6000
         X2              =   6360
         Y1              =   3840
         Y2              =   3840
      End
      Begin Line Line7 
         Index           =   0
         X1              =   6360
         X2              =   6120
         Y1              =   2400
         Y2              =   2400
      End
      Begin Line Line6 
         Index           =   0
         X1              =   6120
         X2              =   6120
         Y1              =   1080
         Y2              =   2400
      End
      Begin Line Line5 
         Index           =   0
         X1              =   6000
         X2              =   6360
         Y1              =   1080
         Y2              =   1080
      End
      Begin Line Line4 
         Index           =   9
         X1              =   9480
         X2              =   9600
         Y1              =   11040
         Y2              =   11040
      End
      Begin Line Line3 
         Index           =   9
         X1              =   7440
         X2              =   7455
         Y1              =   7320
         Y2              =   7335
      End
      Begin Line Line2 
         Index           =   9
         X1              =   9480
         X2              =   9120
         Y1              =   10680
         Y2              =   10680
      End
      Begin Line Line1 
         Index           =   9
         X1              =   9480
         X2              =   9480
         Y1              =   10440
         Y2              =   11040
      End
      Begin Line Line4 
         Index           =   8
         X1              =   9600
         X2              =   9480
         Y1              =   9720
         Y2              =   9720
      End
      Begin Line Line3 
         Index           =   8
         X1              =   9480
         X2              =   9600
         Y1              =   9120
         Y2              =   9120
      End
      Begin Line Line2 
         Index           =   8
         X1              =   9480
         X2              =   9120
         Y1              =   9360
         Y2              =   9360
      End
      Begin Line Line1 
         Index           =   8
         X1              =   9480
         X2              =   9480
         Y1              =   9120
         Y2              =   9720
      End
      Begin Line Line4 
         Index           =   7
         X1              =   9480
         X2              =   9600
         Y1              =   8520
         Y2              =   8520
      End
      Begin Line Line3 
         Index           =   7
         X1              =   9480
         X2              =   9600
         Y1              =   7800
         Y2              =   7800
      End
      Begin Line Line2 
         Index           =   7
         X1              =   9480
         X2              =   9120
         Y1              =   8160
         Y2              =   8160
      End
      Begin Line Line1 
         Index           =   7
         X1              =   9480
         X2              =   9480
         Y1              =   7800
         Y2              =   8520
      End
      Begin Line Line4 
         Index           =   6
         X1              =   9480
         X2              =   9600
         Y1              =   7080
         Y2              =   7080
      End
      Begin Line Line3 
         Index           =   6
         X1              =   9480
         X2              =   9600
         Y1              =   6360
         Y2              =   6360
      End
      Begin Line Line2 
         Index           =   6
         X1              =   6360
         X2              =   6120
         Y1              =   5280
         Y2              =   5280
      End
      Begin Line Line1 
         Index           =   6
         X1              =   9480
         X2              =   9480
         Y1              =   6360
         Y2              =   7080
      End
      Begin Line Line4 
         Index           =   5
         X1              =   9120
         X2              =   9480
         Y1              =   5280
         Y2              =   5280
      End
      Begin Line Line3 
         Index           =   5
         X1              =   9480
         X2              =   9600
         Y1              =   4920
         Y2              =   4920
      End
      Begin Line Line2 
         Index           =   5
         X1              =   9600
         X2              =   9480
         Y1              =   4200
         Y2              =   4200
      End
      Begin Line Line1 
         Index           =   5
         X1              =   9480
         X2              =   9480
         Y1              =   4920
         Y2              =   5640
      End
      Begin Line Line4 
         Index           =   4
         X1              =   9480
         X2              =   9600
         Y1              =   3480
         Y2              =   3480
      End
      Begin Line Line3 
         Index           =   4
         X1              =   9480
         X2              =   9600
         Y1              =   5640
         Y2              =   5640
      End
      Begin Line Line2 
         Index           =   4
         X1              =   9480
         X2              =   9120
         Y1              =   3840
         Y2              =   3840
      End
      Begin Line Line1 
         Index           =   4
         X1              =   9480
         X2              =   9480
         Y1              =   3480
         Y2              =   4200
      End
      Begin Line Line4 
         Index           =   1
         X1              =   9480
         X2              =   9600
         Y1              =   2640
         Y2              =   2640
      End
      Begin Line Line3 
         Index           =   1
         X1              =   9480
         X2              =   9600
         Y1              =   1440
         Y2              =   1440
      End
      Begin Line Line2 
         Index           =   1
         X1              =   9480
         X2              =   9120
         Y1              =   2400
         Y2              =   2400
      End
      Begin Line Line1 
         Index           =   1
         X1              =   9480
         X2              =   9480
         Y1              =   2040
         Y2              =   2640
      End
      Begin Line Line4 
         Index           =   0
         X1              =   9480
         X2              =   9600
         Y1              =   720
         Y2              =   720
      End
      Begin Line Line3 
         Index           =   0
         X1              =   9480
         X2              =   9600
         Y1              =   2040
         Y2              =   2040
      End
      Begin Line Line2 
         Index           =   0
         X1              =   9480
         X2              =   9120
         Y1              =   1080
         Y2              =   1080
      End
      Begin Line Line1 
         Index           =   0
         X1              =   9480
         X2              =   9480
         Y1              =   720
         Y2              =   1440
      End
      Begin Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Dam : "
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   10320
         Width           =   975
      End
      Begin Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Sire : "
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   9000
         Width           =   975
      End
      Begin Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Dam : "
         Height          =   255
         Left            =   6360
         TabIndex        =   18
         Top             =   7800
         Width           =   975
      End
      Begin Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Sire : "
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Top             =   6360
         Width           =   975
      End
      Begin Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Dam : "
         Height          =   255
         Left            =   6360
         TabIndex        =   16
         Top             =   4920
         Width           =   975
      End
      Begin Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Sire : "
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         Top             =   3480
         Width           =   975
      End
      Begin Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Dam : "
         Height          =   255
         Left            =   6360
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/G/Sire : "
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/Dam : "
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   9000
         Width           =   735
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/Sire : "
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   6360
         Width           =   735
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "G/Sire : "
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin Label Label6 
         BackColor       =   &H0000C0C0&
         Caption         =   "G/Dam : "
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   3480
         Width           =   735
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "Dam : "
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6000
         Width           =   615
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "Sire : "
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mark:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ring no.:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6360
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
End
Option Explicit

'Declare GetSysColor lib "user" (byval Color as Integer) as long

'In your form load event:
'me.backcolor=GetSysColor(4)

Dim MyColor(99)
'MyColor(1) = 37 'Blue
'Const MyColor = 37 'Blue



Dim mYr() As Integer
Dim mMark() As String
Dim mRing() As String
Dim mColour() As String
Dim Mcolcode() As String
Dim mEyeSign() As String
Dim mSex() As String
Dim mBredBy() As String
Dim mBredFor() As String
Dim Mbloodline() As String
Dim mName() As String
Dim mNotes() As String
Dim Mstatus() As Integer
Dim MPrtColour() As Integer
Dim MAnalysed() As Integer

Dim TblMaster As Table
Dim TblColour As Table
Dim TblBLine As Table
Dim DB As Database
Dim TblPedg As Table
Dim Mcnt As Integer
Dim Mcnt2 As Integer
Dim Mcnt3 As Integer
Dim Minfo As Integer
Dim MOyr As String
Dim MOmark As String
Dim MOring As String

Sub butt3Colours_Click ()
FRMColours.Show 1
End Sub

Sub butt3dPrintScrn_Click ()
      Const PORTRAIT = 1
      Const Landscape = 2
      Const GETSETPAPERORIENT = 30
      Dim x As Integer

      Dim Orient As OrientStructure

      '* Start the printer
      'Printer.Print ""

      '* Specify the orientation
      Orient.Orientation = Landscape

      '* Send escape sequence to change orientation
      'x% = Escape(Printer.hDC, GETSETPAPERORIENT, Len(Orient), Orient, Null)

      'Printer.Print "Should print in landscape mode"
      'Printer.Print x%
      fmPedg.PrintForm

      Printer.EndDoc

'reset orienation
      '* Specify the orientation
      'Orient.Orientation = PORTRAIT

      '* Send escape sequence to change orientation
      'X% = Escape(Printer.hDC, GETSETPAPERORIENT, Len(Orient), Orient, Null)
      '* The EndDoc will now re-initialize the printer
      'Printer.EndDoc


End Sub

Sub Butt3Info_Click ()
Dim x As Integer

If Minfo = -1 Then 'Info is currently off, turn on
    Butt3Info.Caption = "&Info On"
    'ButtDam.Visible = False
    'ButtSire.Visible = False
    Minfo = 0
    For x = 0 To 29 Step 1
        Panel3D2(x).Visible = False
        Fbirdbb(x).Visible = False
    Next x
    For x = 0 To 5 Step 1
        Panel3D3(x).Visible = False
        fNotes(x).Visible = False
    Next x
Else      'info is off
    Butt3Info.Caption = "&Info Off"
    'ButtDam.Visible = True
    'ButtSire.Visible = True
    Minfo = -1
    For x = 0 To 29 Step 1
        Panel3D2(x).Visible = True
        Fbirdbb(x).Visible = True
    Next x
    For x = 0 To 5 Step 1
        Panel3D3(x).Visible = True
        fNotes(x).Visible = True
    Next x
End If

End Sub

Sub butt3PrintEnh_A_Click ()
Dim Offset As Integer, Trigger As Integer
Dim Mnext_line As Integer
Dim Line_Count As Integer, Item_Count As Integer
Dim CounterR As Integer, CounterC As Integer
Dim Line_Desc As String
Dim Detail As String
Dim Mnotes_1 As String
Dim Mnotes_2 As String
Dim Mnotes_3 As String
Dim x As Integer
Dim Y As Integer

'
Dim Moset_p As Integer
Dim Moset_gp As Integer
Dim Moset_ggp As Integer
Dim Moset_gggp As Integer

Dim Mlmargin As Integer

'Print the Pedigree
Screen.MousePointer = 11

'Set Start of each of the columns
Moset_p% = 360
Moset_gp% = 2800
Moset_ggp% = 5500
Moset_gggp% = 8200

Mlmargin% = 150

Mnext_line% = 400

' Offset for left hand edge of duplicate.
Offset% = 108

' DO LINES

'vertical lines
Printer.Line (Moset_gp% - 20, 4170 - 20)-(Moset_gp% - 20, 14850 - 20)
Printer.Line (Moset_ggp% - 20, 4170 - 20)-(Moset_ggp% - 20, 14850 - 20)
Printer.Line (Moset_gggp% - 20, 4170 - 20)-(Moset_gggp% - 20, 14850 - 20)

'horizontal lines
Printer.Line (Mlmargin%, 9544 - 10)-(11200, 9544 - 10)      'above d
Printer.Line (Moset_gp% - 10, 6778 - 10)-(11200, 6778 - 10) 'above sd
Printer.Line (Moset_gp% - 10, 12150 - 10)-(11200, 12150 - 10) 'above ds

Printer.Line (Moset_ggp% - 10, 5474 - 10)-(11200, 5474 - 10) 'above 9
Printer.Line (Moset_ggp% - 10, 8161 - 10)-(11200, 8161 - 10) 'above 11
Printer.Line (Moset_ggp% - 10, 10847 - 10)-(11200, 10847 - 10) 'above 13
Printer.Line (Moset_ggp% - 10, 13454 - 10)-(11200, 13454 - 10) 'above 15

Printer.Line (Moset_gggp% - 10, 4822 - 10)-(11200, 4822 - 10) 'above 17
Printer.Line (Moset_gggp% - 10, 6126 - 10)-(11200, 6126 - 10) 'above 19
Printer.Line (Moset_gggp% - 10, 7469 - 10)-(11200, 7469 - 10) 'above 21
Printer.Line (Moset_gggp% - 10, 8852 - 10)-(11200, 8852 - 10) 'above 23
Printer.Line (Moset_gggp% - 10, 10195 - 10)-(11200, 10195 - 10) 'above 25
Printer.Line (Moset_gggp% - 10, 11498 - 10)-(11200, 11498 - 10) 'above 27
Printer.Line (Moset_gggp% - 10, 12802 - 10)-(11200, 12802 - 10) 'above 29
Printer.Line (Moset_gggp% - 10, 14106 - 10)-(11200, 14106 - 10) 'above 31

'Printer.Line (Mlmargin%, 14850 - 10)-(11200, 14850 - 10)

'DO HEADING
' Select font and point size.
Printer.FontName = "Arial"
Printer.ScaleMode = 1  ' twips

Printer.FontSize = 22
Printer.FontBold = True
Printer.CurrentX = 4500
Printer.CurrentY = 440
Printer.Print "PEDIGREE"
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = 5402
Printer.CurrentY = 890
Printer.Print "for"

Printer.FontSize = 12
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 1190
' centre maybe ? Mlen% = Len("Bird Life Ring - Year: " & Format$(Myr(1), "") & "  Mark: " & Trim(Mmark(1)) & "  Number: " & Trim(Mring(1)))
Printer.FontBold = False
Printer.Print "Bird Life Ring - Year: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Format$(mYr(1), "");
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Mark: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mMark(1));
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Number: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mRing(1))
Printer.CurrentX = Mlmargin%
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "Colour: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mColour(1);
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Sex: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mSex(1) ' & "  EyeSign: " & Meyesign(1)

Printer.FontItalic = True
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(1));
Printer.CurrentX = 5200
Printer.Print " Bloodline: " & Trim(Mbloodline(1))

Dim mNoteLine1 As String
Dim mNoteLine2 As String
Dim mNoteLine3 As String
Dim mNoteLine4 As String
Dim mNoteLine5 As String
Dim mNoteLine6 As String

mNoteLine1 = ReplaceChars(Left$(mNotes(1), 100) & "")
mNoteLine2 = ReplaceChars(Mid$(mNotes(1), 101, 100) & "")
mNoteLine3 = ReplaceChars(Mid$(mNotes(1), 201, 100) & "")
mNoteLine4 = ReplaceChars(Mid$(mNotes(1), 301, 100) & "")
mNoteLine5 = ReplaceChars(Mid$(mNotes(1), 401, 100) & "")
mNoteLine6 = ReplaceChars(Mid$(mNotes(1), 501, 100) & "")

Printer.FontItalic = True
Printer.FontSize = 9
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print "Notes: " & Trim$(mNoteLine1)
Printer.Print "         " & Trim$(mNoteLine2)
Printer.Print "         " & Trim$(mNoteLine3)
Printer.Print "         " & Trim$(mNoteLine4)
Printer.Print "         " & Trim$(mNoteLine5)
Printer.Print "         " & Trim$(mNoteLine6)

'Printer.CurrentY = 2878
Printer.Print
Printer.FontItalic = False
Printer.CurrentX = Mlmargin%
Printer.Print "Additional Notes :.........................................................................................................................................................."
Printer.Print
Printer.CurrentX = Mlmargin%
Printer.Print "_________________________________________________________________________________________________"

' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Sire  (2)
'
Printer.ForeColor = QBColor(MPrtColour(2))
'
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(2))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(2)) & "/" & Trim(mMark(2)) & "/" & Trim(mRing(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(2))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(2))

'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(2))
If Y% > 580 Then  'Max lines of notes that fit 18linesx28 chars
    Y% = 580
End If
For x% = 1 To Y% Step 29
    Printer.CurrentX = Mlmargin%
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(2), x%, 29)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Dam  (3)
'
Printer.ForeColor = QBColor(MPrtColour(3))
'
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(3))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(3)) & "/" & Trim(mMark(3)) & "/" & Trim(mRing(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(3))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(3))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(3))
If Y% > 580 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 580
End If
For x% = 1 To Y% Step 29
    Printer.CurrentX = Mlmargin%
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(3), x%, 29)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire (Sires Sire) (4)
'
Printer.ForeColor = QBColor(MPrtColour(4))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(4))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(4)) & "/" & Trim(mMark(4)) & "/" & Trim(mRing(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(4))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(4))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(4))
If Y% > 210 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 210
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(4), x%, 30)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Sires Dam) (5)
'
Printer.ForeColor = QBColor(MPrtColour(5))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 6778
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(5))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(5)) & "/" & Trim(mMark(5)) & "/" & Trim(mRing(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(5))
Printer.Print
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(5))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(5))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(5), x%, 30)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire  (Dams Sire) (6)
'
Printer.ForeColor = QBColor(MPrtColour(6))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(6))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(6)) & "/" & Trim(mMark(6)) & "/" & Trim(mRing(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(6))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(6))
'Print Notes
Printer.CurrentX = Moset_gp
Printer.FontSize = 8
Printer.FontBold = False
'Printer.Print "Notes:"
Y% = Len(mNotes(6))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(6), x%, 30)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print (Dams Dam)  (7)
'
Printer.ForeColor = QBColor(MPrtColour(7))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 12150
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(7))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(7))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(7)) & "/" & Trim(mMark(7)) & "/" & Trim(mRing(7)); ""
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(7))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(7))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(7))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(7), x%, 30)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True
    
'Print G/G Sire (Sires Sires Sire) (8)
'
Printer.ForeColor = QBColor(MPrtColour(8))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 4170
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(8))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(8)) & "/" & Trim(mMark(8)) & "/" & Trim(mRing(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(8))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(8))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(8))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(8), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Sires Sires Dam) (9)
'
Printer.ForeColor = QBColor(MPrtColour(9))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(9))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(9)) & "/" & Trim(mMark(9)) & "/" & Trim(mRing(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(9))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(9))
'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_gp

'Printer.Print "Notes:"
Y% = Len(mNotes(9))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
'Printer.CurrentY = 5474
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(9), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print Sires Dams Sire  (Sires Dams Sire) (10)
'
Printer.ForeColor = QBColor(MPrtColour(10))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(10))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(10)) & "/" & Trim(mMark(10)) & "/" & Trim(mRing(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(10))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(10))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(10))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(10), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print Sires Dams Dam  (Sires Dams Dam) (11)
'
Printer.ForeColor = QBColor(MPrtColour(11))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(11))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(11)) & "/" & Trim(mMark(11)) & "/" & Trim(mRing(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(11))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(11))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(11))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(11), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Dams Sires Sire) (12)
'
Printer.ForeColor = QBColor(MPrtColour(12))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 9544
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(12))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(12)) & "/" & Trim(mMark(12)) & "/" & Trim(mRing(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(12))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(12))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp

'Printer.Print "Notes:"
Y% = Len(mNotes(12))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(12), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Dams Sires Dam)
'
Printer.ForeColor = QBColor(MPrtColour(13))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(13))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(13)) & "/" & Trim(mMark(13)) & "/" & Trim(mRing(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(13))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(13))
'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp

'Printer.Print "Notes:"
Y% = Len(mNotes(13))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(13), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print G Dam  (Dams Dams Sire)  (14)
'
Printer.ForeColor = QBColor(MPrtColour(14))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(14))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(14))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(14)) & "/" & Trim(mMark(14)) & "/" & Trim(mRing(14)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(14))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(14))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(14))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(14), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Dams Dams Dam)  (15)
'
Printer.ForeColor = QBColor(MPrtColour(15))
'

Printer.CurrentX = Moset_ggp
Printer.CurrentY = 13454
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(15))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(15))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(15)) & "/" & Trim(mMark(15)) & "/" & Trim(mRing(15)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(15))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(15))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(15))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print ReplaceChars(Trim(Mid$(mNotes(15), x%, 42)))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'5th Generation limit to 3 lines per bird

'Print G/G/G Sire (Sires Sires Sires Sire)
'
Printer.ForeColor = QBColor(MPrtColour(16))
'
Printer.CurrentY = 4170
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(16))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(16))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(16)) & "/" & Trim(mMark(16)) & "/" & Trim(mRing(16))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(16))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(16))

'Print G/G Sire (Sires Sires Sires Dam) (17)
'
Printer.ForeColor = QBColor(MPrtColour(17))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 4822
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(17))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(17))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(17)) & "/" & Trim(mMark(17)) & "/" & Trim(mRing(17))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(17))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(17))

'Print (Sires Sires Dams Sire) (18)
'
Printer.ForeColor = QBColor(MPrtColour(18))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(18))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(18))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(18)) & "/" & Trim(mMark(18)) & "/" & Trim(mRing(18))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(18))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(18))

'Print (Sires Sires Dams Sire) (19)
'
Printer.ForeColor = QBColor(MPrtColour(19))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6126
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(19))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(19))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(19)) & "/" & Trim(mMark(19)) & "/" & Trim(mRing(19))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(19))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(19))

'Print Sires Dams Sires Sire  (Sires Dams Sires Sire) (20)
'
Printer.ForeColor = QBColor(MPrtColour(20))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(20))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(20))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(20)) & "/" & Trim(mMark(20)) & "/" & Trim(mRing(20))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(20))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(20))

'Print Sires Dams Dam  (Sires Dams Sires Dam) (21)
'
Printer.ForeColor = QBColor(MPrtColour(21))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 7469
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(21))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(21))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(21)) & "/" & Trim(mMark(21)) & "/" & Trim(mRing(21))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(21))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(21))

'Print Sires Dams Dams Sire  (Sires Dams Dams Sire) (22)
'
Printer.ForeColor = QBColor(MPrtColour(22))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(22))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(22))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(22)) & "/" & Trim(mMark(22)) & "/" & Trim(mRing(22))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(22))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(22))

'Print Sires Dams Dam  (Sires Dams Dams Dam) (23)
'
Printer.ForeColor = QBColor(MPrtColour(23))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8852
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(23))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(23))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(23)) & "/" & Trim(mMark(23)) & "/" & Trim(mRing(23))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(23))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(23))

'Print G/G/G Sire (Dams Sires Sires Sire)
'
Printer.ForeColor = QBColor(MPrtColour(24))
'
Printer.CurrentY = 9544
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(24))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(24))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(24)) & "/" & Trim(mMark(24)) & "/" & Trim(mRing(24))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(24))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(24))

'Print (Dams Sires Sires Dam)
'
Printer.ForeColor = QBColor(MPrtColour(25))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10195
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(25))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(25))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(25)) & "/" & Trim(mMark(25)) & "/" & Trim(mRing(25))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(25))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(25))

'Print (Dams Sires Dams Sire)
'
Printer.ForeColor = QBColor(MPrtColour(26))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(26))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(26))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(26)) & "/" & Trim(mMark(26)) & "/" & Trim(mRing(26))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(26))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(26))

'Print (Dams Sires Dams Dam)
'
Printer.ForeColor = QBColor(MPrtColour(27))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 11498
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(27))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(27))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(27)) & "/" & Trim(mMark(27)) & "/" & Trim(mRing(27))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(27))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(27))

'Print G Dam  (Dams Dams Sires Sire)  (28)
'
Printer.ForeColor = QBColor(MPrtColour(28))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(28))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(28))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(28)) & "/" & Trim(mMark(28)) & "/" & Trim(mRing(28)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(28))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(28))

'Print G Dam  (Dams Dams Sire)  (29)
'
Printer.ForeColor = QBColor(MPrtColour(29))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12802
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(29))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(29))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(29)) & "/" & Trim(mMark(29)) & "/" & Trim(mRing(29)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(29))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(29))

'Print (Dams Dams Dams Sire)  (30)
'
Printer.ForeColor = QBColor(MPrtColour(30))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 13454
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(30))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(30))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(30)) & "/" & Trim(mMark(30)) & "/" & Trim(mRing(30)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(30))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(30))

'Print (Dams Dams Dams Dam)  (31)
'
Printer.ForeColor = QBColor(MPrtColour(31))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 14106
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(31))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(31))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(31)) & "/" & Trim(mMark(31)) & "/" & Trim(mRing(31)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(31))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(31))

'Print Authorisation:
Printer.ForeColor = QBColor(MPrtColour(0))
Printer.FontItalic = False
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 14850
Printer.Print "Authorised by (Signature) : ";
Printer.FontBold = False
Printer.Print "................................................."
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print Trim(GLLoft) & "-" & GLOwner & " " & GLAddress1 & " " & GLAddress2 & " " & Trim(GLCity) & " " & GLState & " " & GLPostcode & " Phone Number: " & GLPhoneno

'Print Software Serial No.
Printer.FontBold = False
'Printer.FontSize = 8
'Printer.CurrentX = Mlmargin%
'Printer.Print "Software Licence No. " & GLSerialNo
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.FontSize = 6
Printer.Print "This pedigree was produced using The Racing Pigeon Management software RPMS (Copyright 1988-2014, All rights reserved, Kieran Malone Ph 0409495571 www.acepigeons.com.au )"
Printer.Print ""
Printer.CurrentX = Mlmargin% + 4000
Printer.FontSize = 6
Printer.Print "*** ACE Pigeon Auctions - www.acepigeonauctions.com.au ***"
Printer.NewPage
Printer.EndDoc

' Reset mouse pointer.
Screen.MousePointer = 0

End Sub

Sub butt3PrintEnh_Click ()
Dim Offset As Integer, Trigger As Integer
Dim Mnext_line As Integer
Dim Line_Count As Integer, Item_Count As Integer
Dim CounterR As Integer, CounterC As Integer
Dim Line_Desc As String
Dim Detail As String
Dim Mnotes_1 As String
Dim Mnotes_2 As String
Dim Mnotes_3 As String
Dim x As Integer
Dim Y As Integer

'
Dim Moset_p As Integer
Dim Moset_gp As Integer
Dim Moset_ggp As Integer
Dim Moset_gggp As Integer

Dim Mlmargin As Integer

'Print the Pedigree
Screen.MousePointer = 11

'Set Start of each of the columns
Moset_p% = 360
Moset_gp% = 2800
Moset_ggp% = 5500
Moset_gggp% = 8200

Mlmargin% = 150

Mnext_line% = 400

' Offset for left hand edge of duplicate.
Offset% = 108

' DO LINES

'vertical lines
Printer.Line (Moset_gp% - 20, 4170 - 20)-(Moset_gp% - 20, 14850 - 20)
Printer.Line (Moset_ggp% - 20, 4170 - 20)-(Moset_ggp% - 20, 14850 - 20)
Printer.Line (Moset_gggp% - 20, 4170 - 20)-(Moset_gggp% - 20, 14850 - 20)

'horizontal lines
Printer.Line (Mlmargin%, 9544 - 10)-(11200, 9544 - 10)      'above d
Printer.Line (Moset_gp% - 10, 6778 - 10)-(11200, 6778 - 10) 'above sd
Printer.Line (Moset_gp% - 10, 12150 - 10)-(11200, 12150 - 10) 'above ds

Printer.Line (Moset_ggp% - 10, 5474 - 10)-(11200, 5474 - 10) 'above 9
Printer.Line (Moset_ggp% - 10, 8161 - 10)-(11200, 8161 - 10) 'above 11
Printer.Line (Moset_ggp% - 10, 10847 - 10)-(11200, 10847 - 10) 'above 13
Printer.Line (Moset_ggp% - 10, 13454 - 10)-(11200, 13454 - 10) 'above 15

Printer.Line (Moset_gggp% - 10, 4822 - 10)-(11200, 4822 - 10) 'above 17
Printer.Line (Moset_gggp% - 10, 6126 - 10)-(11200, 6126 - 10) 'above 19
Printer.Line (Moset_gggp% - 10, 7469 - 10)-(11200, 7469 - 10) 'above 21
Printer.Line (Moset_gggp% - 10, 8852 - 10)-(11200, 8852 - 10) 'above 23
Printer.Line (Moset_gggp% - 10, 10195 - 10)-(11200, 10195 - 10) 'above 25
Printer.Line (Moset_gggp% - 10, 11498 - 10)-(11200, 11498 - 10) 'above 27
Printer.Line (Moset_gggp% - 10, 12802 - 10)-(11200, 12802 - 10) 'above 29
Printer.Line (Moset_gggp% - 10, 14106 - 10)-(11200, 14106 - 10) 'above 31

'Printer.Line (Mlmargin%, 14850 - 10)-(11200, 14850 - 10)

'DO HEADING
' Select font and point size.
Printer.FontName = "Arial"
Printer.ScaleMode = 1  ' twips

Printer.FontSize = 22
Printer.FontBold = True
Printer.CurrentX = 4500
Printer.CurrentY = 440
Printer.Print "PEDIGREE"
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = 5402
Printer.CurrentY = 890
Printer.Print "for"

Printer.FontSize = 12
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 1190
' centre maybe ? Mlen% = Len("Bird Life Ring - Year: " & Format$(Myr(1), "") & "  Mark: " & Trim(Mmark(1)) & "  Number: " & Trim(Mring(1)))
Printer.FontBold = False
Printer.Print "Bird Life Ring - Year: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Format$(mYr(1), "");
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Mark: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mMark(1));
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Number: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mRing(1))
Printer.CurrentX = Mlmargin%
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "Colour: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mColour(1);
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Sex: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mSex(1) ' & "  EyeSign: " & Meyesign(1)

Printer.FontItalic = True
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(1));
Printer.CurrentX = 5200
Printer.Print " Bloodline: " & Trim(Mbloodline(1))

Dim mNoteLine1 As String
Dim mNoteLine2 As String
Dim mNoteLine3 As String
Dim mNoteLine4 As String
Dim mNoteLine5 As String
Dim mNoteLine6 As String

mNoteLine1 = Left$(mNotes(1), 100) & ""
mNoteLine2 = Mid$(mNotes(1), 101, 100) & ""
mNoteLine3 = Mid$(mNotes(1), 201, 100) & ""
mNoteLine4 = Mid$(mNotes(1), 301, 100) & ""
mNoteLine5 = Mid$(mNotes(1), 401, 100) & ""
mNoteLine6 = Mid$(mNotes(1), 501, 100) & ""

Printer.FontItalic = True
Printer.FontSize = 9
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print "Notes: " & Trim$(mNoteLine1)
Printer.Print "         " & Trim$(mNoteLine2)
Printer.Print "         " & Trim$(mNoteLine3)
Printer.Print "         " & Trim$(mNoteLine4)
Printer.Print "         " & Trim$(mNoteLine5)
Printer.Print "         " & Trim$(mNoteLine6)

'Printer.CurrentY = 2878
Printer.Print
Printer.FontItalic = False
Printer.CurrentX = Mlmargin%
Printer.Print "Additional Notes :.........................................................................................................................................................."
Printer.Print
Printer.CurrentX = Mlmargin%
Printer.Print "_________________________________________________________________________________________________"

' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Sire  (2)
'
Printer.ForeColor = QBColor(MPrtColour(2))
'
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(2))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(2)) & "/" & Trim(mMark(2)) & "/" & Trim(mRing(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(2))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(2))

'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(2))
If Y% > 580 Then  'Max lines of notes that fit 18linesx28 chars
    Y% = 580
End If
For x% = 1 To Y% Step 29
    Printer.CurrentX = Mlmargin%
    Printer.Print Trim(Mid$(mNotes(2), x%, 29))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Dam  (3)
'
Printer.ForeColor = QBColor(MPrtColour(3))
'
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(3))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(3)) & "/" & Trim(mMark(3)) & "/" & Trim(mRing(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(3))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(3))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(3))
If Y% > 580 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 580
End If
For x% = 1 To Y% Step 29
    Printer.CurrentX = Mlmargin%
    Printer.Print Trim(Mid$(mNotes(3), x%, 29))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire (Sires Sire) (4)
'
Printer.ForeColor = QBColor(MPrtColour(4))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(4))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(4)) & "/" & Trim(mMark(4)) & "/" & Trim(mRing(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(4))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(4))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(4))
If Y% > 210 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 210
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(4), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Sires Dam) (5)
'
Printer.ForeColor = QBColor(MPrtColour(5))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 6778
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(5))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(5)) & "/" & Trim(mMark(5)) & "/" & Trim(mRing(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(5))
Printer.Print
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(5))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(5))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(5), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire  (Dams Sire) (6)
'
Printer.ForeColor = QBColor(MPrtColour(6))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(6))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(6)) & "/" & Trim(mMark(6)) & "/" & Trim(mRing(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(6))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(6))
'Print Notes
Printer.CurrentX = Moset_gp
Printer.FontSize = 8
Printer.FontBold = False
'Printer.Print "Notes:"
Y% = Len(mNotes(6))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(6), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print (Dams Dam)  (7)
'
Printer.ForeColor = QBColor(MPrtColour(7))
'
Printer.CurrentX = Moset_gp
Printer.CurrentY = 12150
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(7))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(7))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(7)) & "/" & Trim(mMark(7)) & "/" & Trim(mRing(7)); ""
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(7))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(7))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(7))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(7), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True
    
'Print G/G Sire (Sires Sires Sire) (8)
'
Printer.ForeColor = QBColor(MPrtColour(8))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 4170
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(8))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(8)) & "/" & Trim(mMark(8)) & "/" & Trim(mRing(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(8))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(8))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(8))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(8), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Sires Sires Dam) (9)
'
Printer.ForeColor = QBColor(MPrtColour(9))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(9))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(9)) & "/" & Trim(mMark(9)) & "/" & Trim(mRing(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(9))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(9))
'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_gp

'Printer.Print "Notes:"
Y% = Len(mNotes(9))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
'Printer.CurrentY = 5474
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(9), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print Sires Dams Sire  (Sires Dams Sire) (10)
'
Printer.ForeColor = QBColor(MPrtColour(10))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(10))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(10)) & "/" & Trim(mMark(10)) & "/" & Trim(mRing(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(10))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(10))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(10))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(10), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print Sires Dams Dam  (Sires Dams Dam) (11)
'
Printer.ForeColor = QBColor(MPrtColour(11))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(11))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(11)) & "/" & Trim(mMark(11)) & "/" & Trim(mRing(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(11))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(11))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(11))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(11), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Dams Sires Sire) (12)
'
Printer.ForeColor = QBColor(MPrtColour(12))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 9544
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(12))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(12)) & "/" & Trim(mMark(12)) & "/" & Trim(mRing(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(12))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(12))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp

'Printer.Print "Notes:"
Y% = Len(mNotes(12))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(12), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Dams Sires Dam)
'
Printer.ForeColor = QBColor(MPrtColour(13))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(13))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(13)) & "/" & Trim(mMark(13)) & "/" & Trim(mRing(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(13))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(13))
'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp

'Printer.Print "Notes:"
Y% = Len(mNotes(13))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(13), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print G Dam  (Dams Dams Sire)  (14)
'
Printer.ForeColor = QBColor(MPrtColour(14))
'
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(14))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(14))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(14)) & "/" & Trim(mMark(14)) & "/" & Trim(mRing(14)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(14))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(14))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(14))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(14), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Dams Dams Dam)  (15)
'
Printer.ForeColor = QBColor(MPrtColour(15))
'

Printer.CurrentX = Moset_ggp
Printer.CurrentY = 13454
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(15))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(15))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(15)) & "/" & Trim(mMark(15)) & "/" & Trim(mRing(15)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(15))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(15))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(15))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(15), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'5th Generation limit to 3 lines per bird

'Print G/G/G Sire (Sires Sires Sires Sire)
'
Printer.ForeColor = QBColor(MPrtColour(16))
'
Printer.CurrentY = 4170
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(16))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(16))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(16)) & "/" & Trim(mMark(16)) & "/" & Trim(mRing(16))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(16))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(16))

'Print G/G Sire (Sires Sires Sires Dam) (17)
'
Printer.ForeColor = QBColor(MPrtColour(17))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 4822
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(17))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(17))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(17)) & "/" & Trim(mMark(17)) & "/" & Trim(mRing(17))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(17))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(17))

'Print (Sires Sires Dams Sire) (18)
'
Printer.ForeColor = QBColor(MPrtColour(18))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(18))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(18))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(18)) & "/" & Trim(mMark(18)) & "/" & Trim(mRing(18))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(18))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(18))

'Print (Sires Sires Dams Sire) (19)
'
Printer.ForeColor = QBColor(MPrtColour(19))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6126
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(19))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(19))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(19)) & "/" & Trim(mMark(19)) & "/" & Trim(mRing(19))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(19))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(19))

'Print Sires Dams Sires Sire  (Sires Dams Sires Sire) (20)
'
Printer.ForeColor = QBColor(MPrtColour(20))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(20))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(20))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(20)) & "/" & Trim(mMark(20)) & "/" & Trim(mRing(20))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(20))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(20))

'Print Sires Dams Dam  (Sires Dams Sires Dam) (21)
'
Printer.ForeColor = QBColor(MPrtColour(21))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 7469
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(21))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(21))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(21)) & "/" & Trim(mMark(21)) & "/" & Trim(mRing(21))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(21))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(21))

'Print Sires Dams Dams Sire  (Sires Dams Dams Sire) (22)
'
Printer.ForeColor = QBColor(MPrtColour(22))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(22))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(22))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(22)) & "/" & Trim(mMark(22)) & "/" & Trim(mRing(22))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(22))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(22))

'Print Sires Dams Dam  (Sires Dams Dams Dam) (23)
'
Printer.ForeColor = QBColor(MPrtColour(23))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8852
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(23))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(23))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(23)) & "/" & Trim(mMark(23)) & "/" & Trim(mRing(23))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(23))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(23))

'Print G/G/G Sire (Dams Sires Sires Sire)
'
Printer.ForeColor = QBColor(MPrtColour(24))
'
Printer.CurrentY = 9544
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(24))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(24))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(24)) & "/" & Trim(mMark(24)) & "/" & Trim(mRing(24))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(24))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(24))

'Print (Dams Sires Sires Dam)
'
Printer.ForeColor = QBColor(MPrtColour(25))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10195
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(25))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(25))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(25)) & "/" & Trim(mMark(25)) & "/" & Trim(mRing(25))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(25))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(25))

'Print (Dams Sires Dams Sire)
'
Printer.ForeColor = QBColor(MPrtColour(26))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(26))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(26))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(26)) & "/" & Trim(mMark(26)) & "/" & Trim(mRing(26))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(26))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(26))

'Print (Dams Sires Dams Dam)
'
Printer.ForeColor = QBColor(MPrtColour(27))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 11498
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(27))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(27))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(27)) & "/" & Trim(mMark(27)) & "/" & Trim(mRing(27))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(27))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(27))

'Print G Dam  (Dams Dams Sires Sire)  (28)
'
Printer.ForeColor = QBColor(MPrtColour(28))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(28))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(28))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(28)) & "/" & Trim(mMark(28)) & "/" & Trim(mRing(28)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(28))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(28))

'Print G Dam  (Dams Dams Sire)  (29)
'
Printer.ForeColor = QBColor(MPrtColour(29))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12802
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(29))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(29))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(29)) & "/" & Trim(mMark(29)) & "/" & Trim(mRing(29)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(29))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(29))

'Print (Dams Dams Dams Sire)  (30)
'
Printer.ForeColor = QBColor(MPrtColour(30))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 13454
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(30))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(30))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(30)) & "/" & Trim(mMark(30)) & "/" & Trim(mRing(30)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(30))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(30))

'Print (Dams Dams Dams Dam)  (31)
'
Printer.ForeColor = QBColor(MPrtColour(31))
'
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 14106
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(31))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(31))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(31)) & "/" & Trim(mMark(31)) & "/" & Trim(mRing(31)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(31))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(31))

'Print Authorisation:
Printer.ForeColor = QBColor(MPrtColour(0))
Printer.FontItalic = False
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 14850
Printer.Print "Authorised by (Signature) : ";
Printer.FontBold = False
Printer.Print "................................................."
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print Trim(GLLoft) & "-" & GLOwner & " " & GLAddress1 & " " & GLAddress2 & " " & Trim(GLCity) & " " & GLState & " " & GLPostcode & " Phone Number: " & GLPhoneno

'Print Software Serial No.
Printer.FontBold = False
'Printer.FontSize = 8
'Printer.CurrentX = Mlmargin%
'Printer.Print "Software Licence No. " & GLSerialNo
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.FontSize = 6
Printer.Print "This pedigree was produced using The Racing Pigeon Management software RPMS (Copyright 1988-2014, All rights reserved, Kieran Malone Ph 0409495571 www.acepigeons.com.au )"
Printer.Print ""
Printer.CurrentX = Mlmargin% + 4000
Printer.FontSize = 6
Printer.Print "*** ACE Pigeon Auctions - www.acepigeonauctions.com.au ***"
Printer.NewPage
Printer.EndDoc

' Reset mouse pointer.
Screen.MousePointer = 0
End Sub

Sub ButtBack_Click ()
    Fyr.Text = MOyr
    Fmark.Text = MOmark
    Fring.Text = MOring
    ButtSire.Visible = True
    ButtDam.Visible = True
    Disp_PedgN
End Sub

Sub ButtBird_Click (Index As Integer)
Fyr = mYr(Index + 2)
Fmark = mMark(Index + 2)
Fring = mRing(Index + 2)
Disp_Pedg
End Sub

Sub ButtCancel_Click ()
    Unload fmPedg
End Sub

Sub ButtDam_Click ()
    Fyr.Text = mYr(3)
    Fmark.Text = mMark(3)
    Fring.Text = mRing(3)
    Disp_PedgN
End Sub

Sub buttEditBird_Click ()
On Error GoTo FMerr4ab
    Dim ii As Integer

    FMbird.Task.Text = "M"
    FMbird.Key_yr.Text = Fyr
    FMbird.Key_mark.Text = Fmark
    FMbird.Key_ring_no.Text = Fring
    FMbird.Show 1
    
    DoEvents
    
    'refresh data1 if records deleted
    'If RefreshData.Text = "Y" Then
        'data1.Refresh
        'DoEvents
        'RefreshData.Text = "N"
    'End If
    Exit Sub
FMerr4ab:
    Select Case Err
    Case 3021
        MsgBox "Database Low", 16, "Continue to add birds"
        Resume Next
        'Unload Me
        Exit Sub
    Case 3022
        MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
        MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
    End Select
    
    Resume Next


End Sub

Sub buttPedga_Click ()
Dim Mcnt As Integer
Dim DB1 As Database

Set DB1 = OpenDatabase(DBPath$ & "RPExtra.mdb")
Set TblPedg = DB1.OpenTable("pedg")
Do While Not TblPedg.EOF
    TblPedg.Delete
    TblPedg.MoveNext
Loop
TblPedg.AddNew
Mcnt% = 1
TblPedg("Colour1") = mColour$(1)
TblPedg("EyeSign1") = mEyeSign$(1)
If (mSex$(1) = "C") Then
    TblPedg("Sex1") = "Cock"
ElseIf (mSex$(1) = "H") Then
    TblPedg("Sex1") = "Hen"
Else
    TblPedg("Sex1") = "UnSexed"
End If

For Mcnt% = 1 To 31 Step 1
    TblPedg("yr" & Mcnt%) = mYr%(Mcnt%)
    TblPedg("mark" & Mcnt%) = mMark$(Mcnt%)
    TblPedg("ring" & Mcnt%) = mRing$(Mcnt%)
    TblPedg("bredby" & Mcnt%) = mBredBy$(Mcnt%)
    TblPedg("bredfor" & Mcnt%) = mBredFor$(Mcnt%)
    TblPedg("blood" & Mcnt%) = Mbloodline$(Mcnt%)
    TblPedg("note" & Mcnt%) = mNotes$(Mcnt%)
    TblPedg("name" & Mcnt%) = mName$(Mcnt%)
    TblPedg("colour" & Mcnt%) = mColour$(Mcnt%)
Next Mcnt%

TblPedg.Update
TblPedg.Close
DB1.Close
Set TblPedg = Nothing
Set DB1 = Nothing
Report1.ReportFileName = DBPath$ & "pedg4b.rpt"
Report1.Destination = 0
If Report1.Destination = 1 Then
    On Error GoTo PrinterErr4b
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

PrinterErr4b:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub


End Sub

Sub ButtPrint_Click ()
Dim Mcnt As Integer
Dim DB1 As Database

Set DB1 = OpenDatabase(DBPath$ & "RPExtra.mdb")
Set TblPedg = DB1.OpenTable("pedg")
Do While Not TblPedg.EOF
    TblPedg.Delete
    TblPedg.MoveNext
Loop
TblPedg.AddNew
Mcnt% = 1
TblPedg("Colour1") = mColour$(1)
TblPedg("EyeSign1") = mEyeSign$(1)
If (mSex$(1) = "C") Then
    TblPedg("Sex1") = "Cock"
ElseIf (mSex$(1) = "H") Then
    TblPedg("Sex1") = "Hen"
Else
    TblPedg("Sex1") = "UnSexed"
End If

For Mcnt% = 1 To 31 Step 1
    TblPedg("yr" & Mcnt%) = mYr%(Mcnt%)
    TblPedg("mark" & Mcnt%) = mMark$(Mcnt%)
    TblPedg("ring" & Mcnt%) = mRing$(Mcnt%)
    TblPedg("bredby" & Mcnt%) = mBredBy$(Mcnt%)
    TblPedg("blood" & Mcnt%) = Mbloodline$(Mcnt%)
    TblPedg("note" & Mcnt%) = mNotes$(Mcnt%)
    TblPedg("name" & Mcnt%) = mName$(Mcnt%)
Next Mcnt%

TblPedg.Update
TblPedg.Close
DB1.Close
Set TblPedg = Nothing
Set DB1 = Nothing
Report1.ReportFileName = DBPath$ & "pedg.rpt"
Report1.Destination = 0
If Report1.Destination = 1 Then
    On Error GoTo PrinterErr1
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

PrinterErr1:
     MsgBox Error$, 16, "Report Printing Error"

    Exit Sub

End Sub

Sub ButtPrintN_Click ()

Dim Offset As Integer, Trigger As Integer
Dim Mnext_line As Integer
Dim Line_Count As Integer, Item_Count As Integer
Dim CounterR As Integer, CounterC As Integer
Dim Line_Desc As String
Dim Detail As String
Dim Mnotes_1 As String
Dim Mnotes_2 As String
Dim Mnotes_3 As String
Dim x As Integer
Dim Y As Integer

'
Dim Moset_p As Integer
Dim Moset_gp As Integer
Dim Moset_ggp As Integer
Dim Moset_gggp As Integer

Dim Mlmargin As Integer

'Print the Pedigree
Screen.MousePointer = 11

'Set Start of each of the columns
Moset_p% = 360
Moset_gp% = 2800
Moset_ggp% = 5500
Moset_gggp% = 8200

Mlmargin% = 150

Mnext_line% = 400

' Offset for left hand edge of duplicate.
Offset% = 108

' DO LINES

'vertical lines
Printer.Line (Moset_gp% - 20, 4170 - 20)-(Moset_gp% - 20, 14850 - 20)
Printer.Line (Moset_ggp% - 20, 4170 - 20)-(Moset_ggp% - 20, 14850 - 20)
Printer.Line (Moset_gggp% - 20, 4170 - 20)-(Moset_gggp% - 20, 14850 - 20)

'horizontal lines
Printer.Line (Mlmargin%, 9544 - 10)-(11200, 9544 - 10)      'above d
Printer.Line (Moset_gp% - 10, 6778 - 10)-(11200, 6778 - 10) 'above sd
Printer.Line (Moset_gp% - 10, 12150 - 10)-(11200, 12150 - 10) 'above ds

Printer.Line (Moset_ggp% - 10, 5474 - 10)-(11200, 5474 - 10) 'above 9
Printer.Line (Moset_ggp% - 10, 8161 - 10)-(11200, 8161 - 10) 'above 11
Printer.Line (Moset_ggp% - 10, 10847 - 10)-(11200, 10847 - 10) 'above 13
Printer.Line (Moset_ggp% - 10, 13454 - 10)-(11200, 13454 - 10) 'above 15

Printer.Line (Moset_gggp% - 10, 4822 - 10)-(11200, 4822 - 10) 'above 17
Printer.Line (Moset_gggp% - 10, 6126 - 10)-(11200, 6126 - 10) 'above 19
Printer.Line (Moset_gggp% - 10, 7469 - 10)-(11200, 7469 - 10) 'above 21
Printer.Line (Moset_gggp% - 10, 8852 - 10)-(11200, 8852 - 10) 'above 23
Printer.Line (Moset_gggp% - 10, 10195 - 10)-(11200, 10195 - 10) 'above 25
Printer.Line (Moset_gggp% - 10, 11498 - 10)-(11200, 11498 - 10) 'above 27
Printer.Line (Moset_gggp% - 10, 12802 - 10)-(11200, 12802 - 10) 'above 29
Printer.Line (Moset_gggp% - 10, 14106 - 10)-(11200, 14106 - 10) 'above 31

'Printer.Line (Mlmargin%, 14850 - 10)-(11200, 14850 - 10)

'DO HEADING
' Select font and point size.
Printer.FontName = "Arial"
Printer.ScaleMode = 1  ' twips

Printer.FontSize = 22
Printer.FontBold = True
Printer.CurrentX = 4500
Printer.CurrentY = 440
Printer.Print "PEDIGREE"
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = 5402
Printer.CurrentY = 890
Printer.Print "for"

Printer.FontSize = 12
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 1190
' centre maybe ? Mlen% = Len("Bird Life Ring - Year: " & Format$(Myr(1), "") & "  Mark: " & Trim(Mmark(1)) & "  Number: " & Trim(Mring(1)))
Printer.FontBold = False
Printer.Print "Bird Life Ring - Year: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Format$(mYr(1), "");
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Mark: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mMark(1));
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Number: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mRing(1))
Printer.CurrentX = Mlmargin%
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "Colour: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mColour(1);
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Sex: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mSex(1) ' & "  EyeSign: " & Meyesign(1)

Printer.FontItalic = True
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(1));
Printer.CurrentX = 5200
Printer.Print " Bloodline: " & Trim(Mbloodline(1))

Dim mNoteLine1 As String
Dim mNoteLine2 As String
Dim mNoteLine3 As String
Dim mNoteLine4 As String
Dim mNoteLine5 As String
Dim mNoteLine6 As String

mNoteLine1 = Left$(mNotes(1), 100) & ""
mNoteLine2 = Mid$(mNotes(1), 101, 100) & ""
mNoteLine3 = Mid$(mNotes(1), 201, 100) & ""
mNoteLine4 = Mid$(mNotes(1), 301, 100) & ""
mNoteLine5 = Mid$(mNotes(1), 401, 100) & ""
mNoteLine6 = Mid$(mNotes(1), 501, 100) & ""

Printer.FontItalic = True
Printer.FontSize = 9
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print "Notes: " & Trim$(mNoteLine1)
Printer.Print "         " & Trim$(mNoteLine2)
Printer.Print "         " & Trim$(mNoteLine3)
Printer.Print "         " & Trim$(mNoteLine4)
Printer.Print "         " & Trim$(mNoteLine5)
Printer.Print "         " & Trim$(mNoteLine6)

'Printer.CurrentY = 2878
Printer.Print
Printer.FontItalic = False
Printer.CurrentX = Mlmargin%
Printer.Print "Additional Notes :.........................................................................................................................................................."
Printer.Print
Printer.CurrentX = Mlmargin%
Printer.Print "_________________________________________________________________________________________________"

' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Sire  (2)
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(2))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(2)) & "/" & Trim(mMark(2)) & "/" & Trim(mRing(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(2))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(2))

'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(2))
If Y% > 580 Then  'Max lines of notes that fit 18linesx28 chars
    Y% = 580
End If
For x% = 1 To Y% Step 29
    Printer.CurrentX = Mlmargin%
    Printer.Print Trim(Mid$(mNotes(2), x%, 29))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Dam  (3)
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(3))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(3)) & "/" & Trim(mMark(3)) & "/" & Trim(mRing(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(3))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(3))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(3))
If Y% > 580 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 580
End If
For x% = 1 To Y% Step 29
    Printer.CurrentX = Mlmargin%
    Printer.Print Trim(Mid$(mNotes(3), x%, 29))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire (Sires Sire) (4)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(4))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(4)) & "/" & Trim(mMark(4)) & "/" & Trim(mRing(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(4))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(4))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(4))
If Y% > 210 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 210
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(4), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Sires Dam) (5)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 6778
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(5))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(5)) & "/" & Trim(mMark(5)) & "/" & Trim(mRing(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(5))
Printer.Print
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(5))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(5))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(5), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire  (Dams Sire) (6)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(6))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(6)) & "/" & Trim(mMark(6)) & "/" & Trim(mRing(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(6))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(6))
'Print Notes
Printer.CurrentX = Moset_gp
Printer.FontSize = 8
Printer.FontBold = False
'Printer.Print "Notes:"
Y% = Len(mNotes(6))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(6), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print (Dams Dam)  (7)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 12150
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(7))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(7))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(7)) & "/" & Trim(mMark(7)) & "/" & Trim(mRing(7)); ""
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(7))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(7))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
'Printer.Print "Notes:"
Y% = Len(mNotes(7))
If Y% > 240 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 240
End If
For x% = 1 To Y% Step 30
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(7), x%, 30))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True
    
'Print G/G Sire (Sires Sires Sire) (8)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 4170
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(8))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(8)) & "/" & Trim(mMark(8)) & "/" & Trim(mRing(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(8))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(8))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(8))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(8), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Sires Sires Dam) (9)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(9))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(9)) & "/" & Trim(mMark(9)) & "/" & Trim(mRing(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(9))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(9))
'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_gp

'Printer.Print "Notes:"
Y% = Len(mNotes(9))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
'Printer.CurrentY = 5474
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(9), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print Sires Dams Sire  (Sires Dams Sire) (10)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(10))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(10)) & "/" & Trim(mMark(10)) & "/" & Trim(mRing(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(10))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(10))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(10))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(10), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print Sires Dams Dam  (Sires Dams Dam) (11)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(11))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(11)) & "/" & Trim(mMark(11)) & "/" & Trim(mRing(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(11))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(11))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(11))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(11), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Dams Sires Sire) (12)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 9544
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(12))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(12)) & "/" & Trim(mMark(12)) & "/" & Trim(mRing(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(12))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(12))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp

'Printer.Print "Notes:"
Y% = Len(mNotes(12))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(12), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print (Dams Sires Dam)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(13))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(13)) & "/" & Trim(mMark(13)) & "/" & Trim(mRing(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(13))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(13))
'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp

'Printer.Print "Notes:"
Y% = Len(mNotes(13))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(13), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True


'Print G Dam  (Dams Dams Sire)  (14)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(14))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(14))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(14)) & "/" & Trim(mMark(14)) & "/" & Trim(mRing(14)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(14))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(14))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(14))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(14), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Dams Dams Dam)  (15)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 13454
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(15))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(15))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(15)) & "/" & Trim(mMark(15)) & "/" & Trim(mRing(15)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(15))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(15))

'Print Notes
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = Moset_ggp
'Printer.Print "Notes:"
Y% = Len(mNotes(15))
If Y% > 84 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 84
End If
For x% = 1 To Y% Step 42
    Printer.CurrentX = Moset_ggp
    Printer.Print Trim(Mid$(mNotes(15), x%, 42))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'5th Generation limit to 3 lines per bird

'Print G/G/G Sire (Sires Sires Sires Sire)
Printer.CurrentY = 4170
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(16))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(16))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(16)) & "/" & Trim(mMark(16)) & "/" & Trim(mRing(16))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(16))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(16))

'Print G/G Sire (Sires Sires Sires Dam) (17)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 4822
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(17))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(17))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(17)) & "/" & Trim(mMark(17)) & "/" & Trim(mRing(17))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(17))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(17))

'Print (Sires Sires Dams Sire) (18)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(18))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(18))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(18)) & "/" & Trim(mMark(18)) & "/" & Trim(mRing(18))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(18))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(18))

'Print (Sires Sires Dams Sire) (19)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6126
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(19))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(19))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(19)) & "/" & Trim(mMark(19)) & "/" & Trim(mRing(19))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(19))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(19))

'Print Sires Dams Sires Sire  (Sires Dams Sires Sire) (20)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(20))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(20))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(20)) & "/" & Trim(mMark(20)) & "/" & Trim(mRing(20))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(20))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(20))

'Print Sires Dams Dam  (Sires Dams Sires Dam) (21)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 7469
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(21))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(21))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(21)) & "/" & Trim(mMark(21)) & "/" & Trim(mRing(21))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(21))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(21))

'Print Sires Dams Dams Sire  (Sires Dams Dams Sire) (22)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(22))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(22))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(22)) & "/" & Trim(mMark(22)) & "/" & Trim(mRing(22))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(22))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(22))

'Print Sires Dams Dam  (Sires Dams Dams Dam) (23)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8852
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(23))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(23))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(23)) & "/" & Trim(mMark(23)) & "/" & Trim(mRing(23))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(23))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(23))

'Print G/G/G Sire (Dams Sires Sires Sire)
Printer.CurrentY = 9544
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(24))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(24))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(24)) & "/" & Trim(mMark(24)) & "/" & Trim(mRing(24))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(24))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(24))

'Print (Dams Sires Sires Dam)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10195
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(25))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(25))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(25)) & "/" & Trim(mMark(25)) & "/" & Trim(mRing(25))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(25))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(25))

'Print (Dams Sires Dams Sire)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(26))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(26))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(26)) & "/" & Trim(mMark(26)) & "/" & Trim(mRing(26))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(26))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(26))

'Print (Dams Sires Dams Dam)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 11498
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(27))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(27))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(27)) & "/" & Trim(mMark(27)) & "/" & Trim(mRing(27))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(27))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(27))

'Print G Dam  (Dams Dams Sires Sire)  (28)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(28))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(28))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(28)) & "/" & Trim(mMark(28)) & "/" & Trim(mRing(28)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(28))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(28))

'Print G Dam  (Dams Dams Sire)  (29)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12802
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(29))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(29))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(29)) & "/" & Trim(mMark(29)) & "/" & Trim(mRing(29)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(29))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(29))

'Print (Dams Dams Dams Sire)  (30)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 13454
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(30))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(30))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(30)) & "/" & Trim(mMark(30)) & "/" & Trim(mRing(30)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(30))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(30))

'Print (Dams Dams Dams Dam)  (31)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 14106
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(31))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(31))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(31)) & "/" & Trim(mMark(31)) & "/" & Trim(mRing(31)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(31))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(31))

'Print Authorisation:
Printer.FontItalic = False
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 14850
Printer.Print "Authorised by (Signature) : ";
Printer.FontBold = False
Printer.Print "................................................."
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print Trim(GLLoft) & "-" & GLOwner & " " & GLAddress1 & " " & GLAddress2 & " " & Trim(GLCity) & " " & GLState & " " & GLPostcode & " Phone Number: " & GLPhoneno

'Print Software Serial No.
Printer.FontBold = False
'Printer.FontSize = 8
'Printer.CurrentX = Mlmargin%
'Printer.Print "Software Licence No. " & GLSerialNo
Printer.CurrentX = Mlmargin%
Printer.FontSize = 6
Printer.Print "This pedigree was produced using The Racing Pigeon Management software RPMS (Copyright 1988-2014, All rights reserved, Kieran Malone Ph 0409495571)"

Printer.NewPage
Printer.EndDoc

' Reset mouse pointer.
Screen.MousePointer = 0

End Sub

Sub ButtPrintPedgN_Click ()

Dim Offset As Integer, Trigger As Integer
Dim Mnext_line As Integer
Dim Line_Count As Integer, Item_Count As Integer
Dim CounterR As Integer, CounterC As Integer
Dim Line_Desc As String
Dim Detail As String
Dim Mnotes_1 As String
Dim Mnotes_2 As String
Dim Mnotes_3 As String
Dim x As Integer
Dim Y As Integer

'
Dim Moset_p As Integer
Dim Moset_gp As Integer
Dim Moset_ggp As Integer
Dim Moset_gggp As Integer

Dim Mlmargin As Integer

'Print the Pedigree
Screen.MousePointer = 11

'Set Start of each of the columns
Moset_p% = 360
Moset_gp% = 2800
Moset_ggp% = 5500
Moset_gggp% = 8200

Mlmargin% = 150

Mnext_line% = 400

' Offset for left hand edge of duplicate.
Offset% = 108

' DO LINES

'vertical lines
Printer.Line (Moset_gp% - 20, 4170 - 20)-(Moset_gp% - 20, 14850 - 20)
Printer.Line (Moset_ggp% - 20, 4170 - 20)-(Moset_ggp% - 20, 14850 - 20)
Printer.Line (Moset_gggp% - 20, 4170 - 20)-(Moset_gggp% - 20, 14850 - 20)

'horizontal lines
Printer.Line (Mlmargin%, 9544 - 10)-(11200, 9544 - 10)      'above d
Printer.Line (Moset_gp% - 10, 6778 - 10)-(11200, 6778 - 10) 'above sd
Printer.Line (Moset_gp% - 10, 12150 - 10)-(11200, 12150 - 10) 'above ds

Printer.Line (Moset_ggp% - 10, 5474 - 10)-(11200, 5474 - 10) 'above 9
Printer.Line (Moset_ggp% - 10, 8161 - 10)-(11200, 8161 - 10) 'above 11
Printer.Line (Moset_ggp% - 10, 10847 - 10)-(11200, 10847 - 10) 'above 13
Printer.Line (Moset_ggp% - 10, 13454 - 10)-(11200, 13454 - 10) 'above 15

Printer.Line (Moset_gggp% - 10, 4822 - 10)-(11200, 4822 - 10) 'above 17
Printer.Line (Moset_gggp% - 10, 6126 - 10)-(11200, 6126 - 10) 'above 19
Printer.Line (Moset_gggp% - 10, 7469 - 10)-(11200, 7469 - 10) 'above 21
Printer.Line (Moset_gggp% - 10, 8852 - 10)-(11200, 8852 - 10) 'above 23
Printer.Line (Moset_gggp% - 10, 10195 - 10)-(11200, 10195 - 10) 'above 25
Printer.Line (Moset_gggp% - 10, 11498 - 10)-(11200, 11498 - 10) 'above 27
Printer.Line (Moset_gggp% - 10, 12802 - 10)-(11200, 12802 - 10) 'above 29
Printer.Line (Moset_gggp% - 10, 14106 - 10)-(11200, 14106 - 10) 'above 31

'Printer.Line (Mlmargin%, 14850 - 10)-(11200, 14850 - 10)

'DO HEADING
' Select font and point size.
Printer.FontName = "Arial"
Printer.ScaleMode = 1  ' twips

Printer.FontSize = 22
Printer.FontBold = True
Printer.CurrentX = 4500
Printer.CurrentY = 440
Printer.Print "PEDIGREE"
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = 5402
Printer.CurrentY = 890
Printer.Print "for"

Printer.FontSize = 12
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 1190
' centre maybe ? Mlen% = Len("Bird Life Ring - Year: " & Format$(Myr(1), "") & "  Mark: " & Trim(Mmark(1)) & "  Number: " & Trim(Mring(1)))
Printer.FontBold = False
Printer.Print "Bird Life Ring - Year: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Format$(mYr(1), "");
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Mark: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mMark(1));
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Number: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print Trim(mRing(1))
Printer.CurrentX = Mlmargin%
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "Colour: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mColour(1);
Printer.FontBold = False
Printer.FontSize = 12
Printer.Print "  Sex: ";
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print mSex(1) ' & "  EyeSign: " & Meyesign(1)

Printer.FontSize = 12
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(1));
Printer.CurrentX = 5200
Printer.Print " Bloodline: " & Trim(Mbloodline(1))
If Len(Trim(mNotes(1))) > 0 Then
    If Len(mNotes(1)) < 71 Then
        Printer.CurrentX = Mlmargin%
        Printer.Print "Notes: " & Trim(mNotes(1))
    Else
        Printer.CurrentX = Mlmargin%
        Printer.Print "Notes: " & Left$(mNotes(1), 70)
        If Len(mNotes(1)) < 141 Then
            Printer.CurrentX = Mlmargin%
            Printer.Print "         " & Mid$(mNotes(1), 71, Len(mNotes(1)) - 70)
        Else
            Printer.CurrentX = Mlmargin%
            Printer.Print "         " & Mid$(mNotes(1), 71, 70)
        End If
    End If
End If
'Printer.CurrentY = 2878
Printer.Print
Printer.CurrentX = Mlmargin%
Printer.Print "Additional Notes :.........................................................................................................................................................."
Printer.Print
Printer.CurrentX = Mlmargin%
Printer.Print "............................................................................................................................................................................"
Printer.CurrentX = Mlmargin%
Printer.Print "_________________________________________________________________________________________________"

' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Sire  (2)
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(2))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(2)) & "/" & Trim(mMark(2)) & "/" & Trim(mRing(2))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(2))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(2))

'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(2))
If Y% > 400 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 400
End If
For x% = 1 To Y% Step 25
    Printer.CurrentX = Mlmargin%
    Printer.Print Trim(Mid$(mNotes(2), x%, 25))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print Dam  (3)
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(mColour(3))
Printer.CurrentX = Mlmargin%
Printer.Print Format$(mYr(3)) & "/" & Trim(mMark(3)) & "/" & Trim(mRing(3))
Printer.CurrentX = Mlmargin%
Printer.Print Trim(Mbloodline(3))
Printer.CurrentX = Mlmargin%
Printer.Print "Breeder: " & Trim(mBredBy(3))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Mlmargin%
Printer.Print ""
Printer.CurrentX = Mlmargin%
Printer.Print "Notes:"
Y% = Len(mNotes(3))
If Y% > 400 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 400
End If
For x% = 1 To Y% Step 25
    Printer.CurrentX = Mlmargin%
    Printer.Print Trim(Mid$(mNotes(3), x%, 25))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire (Sires Sire) (4)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 4170
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(4))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(4)) & "/" & Trim(mMark(4)) & "/" & Trim(mRing(4))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(4))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(4))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
Printer.Print "Notes:"
Y% = Len(mNotes(4))
If Y% > 150 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 150
End If
For x% = 1 To Y% Step 25
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(4), x%, 25))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Dam  (Sires Dam) (5)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 6778
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(5))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(5)) & "/" & Trim(mMark(5)) & "/" & Trim(mRing(5))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(5))
Printer.Print
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(5))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
Printer.Print "Notes:"
Y% = Len(mNotes(5))
If Y% > 150 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 150
End If
For x% = 1 To Y% Step 25
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(5), x%, 25))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print G Sire  (Dams Sire) (6)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 9544
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(6))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(6)) & "/" & Trim(mMark(6)) & "/" & Trim(mRing(6))
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(6))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(6))
'Print Notes
Printer.CurrentX = Moset_gp
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print "Notes:"
Y% = Len(mNotes(6))
If Y% > 150 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 150
End If
For x% = 1 To Y% Step 25
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(6), x%, 25))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True

'Print (Dams Dam)  (7)
Printer.CurrentX = Moset_gp
Printer.CurrentY = 12150
Printer.FontItalic = False
Printer.FontUnderline = True
Printer.Print "G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(7))
Printer.CurrentX = Moset_gp
Printer.Print Trim(mColour(7))
Printer.CurrentX = Moset_gp
Printer.Print Format$(mYr(7)) & "/" & Trim(mMark(7)) & "/" & Trim(mRing(7)); ""
Printer.CurrentX = Moset_gp
Printer.Print Trim(Mbloodline(7))
Printer.CurrentX = Moset_gp
Printer.Print "Breeder: " & Trim(mBredBy(7))
'Print Notes
Printer.FontSize = 8
Printer.FontBold = False
Printer.CurrentX = Moset_gp
Printer.Print "Notes:"
Y% = Len(mNotes(7))
If Y% > 150 Then  'Max lines of notes that fit 16linesx25chars
    Y% = 150
End If
For x% = 1 To Y% Step 25
    Printer.CurrentX = Moset_gp
    Printer.Print Trim(Mid$(mNotes(7), x%, 25))
Next x%
' Select font and point size.
Printer.FontName = "Arial"
Printer.FontSize = 9
Printer.FontBold = True
    
'Print G/G Sire (Sires Sires Sire) (8)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 4170
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(8))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(8)) & "/" & Trim(mMark(8)) & "/" & Trim(mRing(8))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(8))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(8))

'Print (Sires Sires Dam) (9)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(9))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(9)) & "/" & Trim(mMark(9)) & "/" & Trim(mRing(9))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(9))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(9))

'Print Sires Dams Sire  (Sires Dams Sire) (10)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(10))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(10)) & "/" & Trim(mMark(10)) & "/" & Trim(mRing(10))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(10))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(10))

'Print Sires Dams Dam  (Sires Dams Dam) (11)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(11))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(11)) & "/" & Trim(mMark(11)) & "/" & Trim(mRing(11))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(11))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(11))

'Print (Dams Sires Sire) (12)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 9544
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(12))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(12)) & "/" & Trim(mMark(12)) & "/" & Trim(mRing(12))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(12))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(12))

'Print (Dams Sires Dam)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(13))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(13)) & "/" & Trim(mMark(13)) & "/" & Trim(mRing(13))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(13))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(13))

'Print G Dam  (Dams Dams Sire)  (14)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(14))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(14))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(14)) & "/" & Trim(mMark(14)) & "/" & Trim(mRing(14)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(14))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(14))

'Print G Dam  (Dams Dams Dam)  (15)
Printer.CurrentX = Moset_ggp
Printer.CurrentY = 13454
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(15))
Printer.CurrentX = Moset_ggp
Printer.Print Trim(mColour(15))
Printer.CurrentX = Moset_ggp
Printer.Print Format$(mYr(15)) & "/" & Trim(mMark(15)) & "/" & Trim(mRing(15)); ""
Printer.CurrentX = Moset_ggp
Printer.Print Trim(Mbloodline(15))
Printer.CurrentX = Moset_ggp
Printer.Print "Breeder: " & Trim(mBredBy(15))

'5th Generation limit to 3 lines per bird

'Print G/G/G Sire (Sires Sires Sires Sire)
Printer.CurrentY = 4170
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(16))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(16))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(16)) & "/" & Trim(mMark(16)) & "/" & Trim(mRing(16))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(16))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(16))

'Print G/G Sire (Sires Sires Sires Dam) (17)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 4822
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(17))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(17))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(17)) & "/" & Trim(mMark(17)) & "/" & Trim(mRing(17))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(17))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(17))

'Print (Sires Sires Dams Sire) (18)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 5474
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(18))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(18))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(18)) & "/" & Trim(mMark(18)) & "/" & Trim(mRing(18))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(18))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(18))

'Print (Sires Sires Dams Sire) (19)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6126
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(19))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(19))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(19)) & "/" & Trim(mMark(19)) & "/" & Trim(mRing(19))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(19))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(19))

'Print Sires Dams Sires Sire  (Sires Dams Sires Sire) (20)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 6778
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(20))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(20))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(20)) & "/" & Trim(mMark(20)) & "/" & Trim(mRing(20))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(20))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(20))

'Print Sires Dams Dam  (Sires Dams Sires Dam) (21)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 7469
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(21))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(21))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(21)) & "/" & Trim(mMark(21)) & "/" & Trim(mRing(21))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(21))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(21))

'Print Sires Dams Dams Sire  (Sires Dams Dams Sire) (22)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8161
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(22))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(22))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(22)) & "/" & Trim(mMark(22)) & "/" & Trim(mRing(22))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(22))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(22))

'Print Sires Dams Dam  (Sires Dams Dams Dam) (23)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 8852
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(23))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(23))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(23)) & "/" & Trim(mMark(23)) & "/" & Trim(mRing(23))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(23))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(23))

'Print G/G/G Sire (Dams Sires Sires Sire)
Printer.CurrentY = 9544
Printer.CurrentX = Moset_gggp
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(24))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(24))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(24)) & "/" & Trim(mMark(24)) & "/" & Trim(mRing(24))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(24))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(24))

'Print (Dams Sires Sires Dam)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10195
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(25))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(25))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(25)) & "/" & Trim(mMark(25)) & "/" & Trim(mRing(25))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(25))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(25))

'Print (Dams Sires Dams Sire)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 10847
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(26))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(26))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(26)) & "/" & Trim(mMark(26)) & "/" & Trim(mRing(26))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(26))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(26))

'Print (Dams Sires Dams Dam)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 11498
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(27))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(27))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(27)) & "/" & Trim(mMark(27)) & "/" & Trim(mRing(27))
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(27))
'Printer.CurrentX = Moset_ggp
'Printer.Print "Breeder: " & Trim(Mbredby(27))

'Print G Dam  (Dams Dams Sires Sire)  (28)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12150
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(28))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(28))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(28)) & "/" & Trim(mMark(28)) & "/" & Trim(mRing(28)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(28))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(28))

'Print G Dam  (Dams Dams Sire)  (29)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 12802
Printer.FontItalic = False
'Printer.FontBold = False
Printer.FontUnderline = True
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(29))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(29))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(29)) & "/" & Trim(mMark(29)) & "/" & Trim(mRing(29)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(29))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(29))

'Print (Dams Dams Dams Sire)  (30)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 13454
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/SIRE: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(30))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(30))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(30)) & "/" & Trim(mMark(30)) & "/" & Trim(mRing(30)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(30))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(30))

'Print (Dams Dams Dams Dam)  (31)
Printer.CurrentX = Moset_gggp
Printer.CurrentY = 14106
Printer.FontUnderline = True
Printer.FontItalic = False
'Printer.FontBold = False
Printer.Print "G/G/G/DAM: ";
Printer.FontUnderline = False
Printer.FontItalic = True
Printer.Print Trim(mName(31))
Printer.CurrentX = Moset_gggp
Printer.Print Trim(mColour(31))
Printer.CurrentX = Moset_gggp
Printer.Print Format$(mYr(31)) & "/" & Trim(mMark(31)) & "/" & Trim(mRing(31)); ""
Printer.CurrentX = Moset_gggp
'Printer.Print Trim(Mbloodline(31))
'Printer.CurrentX = Moset_gggp
'Printer.Print "Breeder: " & Trim(Mbredby(31))

'Print Authorisation:
Printer.FontItalic = False
Printer.FontSize = 10
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.CurrentY = 14850
Printer.Print "Authorised by (Signature) : ";
Printer.FontBold = False
Printer.Print "................................................."
Printer.FontBold = True
Printer.CurrentX = Mlmargin%
Printer.Print Trim(GLLoft) & "-" & GLOwner & " " & GLAddress1 & " " & GLAddress2 & " " & Trim(GLCity) & " " & GLState & " " & GLPostcode & " Phone Number: " & GLPhoneno

'Print Software Serial No.
Printer.FontBold = False
'Printer.FontSize = 8
'Printer.CurrentX = Mlmargin%
'Printer.Print "Software Licence No. " & GLSerialNo
Printer.CurrentX = Mlmargin%
Printer.FontSize = 6
Printer.Print "This pedigree was produced using The Racing Pigeon Management software RPMS (Copyright 1988-2014, All rights reserved, Kieran Malone Ph 0409495571)"

Printer.NewPage
Printer.EndDoc

' Reset mouse pointer.
Screen.MousePointer = 0

End Sub

Sub ButtSire_Click ()
    Fyr = mYr(2)
    Fmark = mMark(2)
    Fring = mRing(2)
    Disp_PedgN
End Sub

Sub cmdDetails_Click ()
If frmDetails.Visible Then
    frmDetails.Visible = False
Else
    frmDetails.Visible = True
End If
End Sub

Sub Disp_Pedg ()

ReDim mYr(31)
ReDim mMark(31)
ReDim mRing(31)
ReDim mColour(31)
ReDim Mcolcode(31)
ReDim mEyeSign(31)
ReDim mSex(31)
ReDim mBredBy(31)
ReDim mBredFor(31)
ReDim Mbloodline(31)
ReDim mName(31)
ReDim mNotes(31)
ReDim Mstatus(31)

Minfo = -1 'Info is ON, Screen Cluttered

Mcnt% = 0
Do While Mcnt% < 31
    Mcnt% = Mcnt% + 1
    mYr%(Mcnt%) = 0
    mMark$(Mcnt%) = ""
    Mcolcode$(Mcnt%) = ""
    mColour$(Mcnt%) = ""
    mEyeSign$(Mcnt%) = ""
    mSex$(Mcnt%) = ""
    mRing$(Mcnt%) = ""
    mBredBy$(Mcnt%) = ""
    mBredFor$(Mcnt%) = ""
    Mbloodline$(Mcnt%) = ""
    mNotes$(Mcnt%) = ""
    mName$(Mcnt%) = ""
    Mstatus%(Mcnt%) = False
Loop

Mcnt% = 0
Do While Mcnt% < 30  'Bird + 30 ancestors Bird(0) to Bird(29)
    FbirdYMR(Mcnt%).Text = ""
    Mcnt% = Mcnt% + 1
Loop

For Mcnt% = 0 To 13 Step 1  'only Bird + 14 ancestors Bird(0) to Bird (13)
    Fbirdbb(Mcnt%).Text = ""
Next Mcnt%

For Mcnt% = 0 To 13 Step 1  'only Bird + 14 ancestors Bird(0) to Bird (13)
    Panel3D2(Mcnt%).Caption = ""
Next Mcnt%

For Mcnt% = 0 To 5 Step 1   'bloodline display
    Panel3D3(Mcnt%).Caption = ""
Next Mcnt%
fNotes(0) = ""
fNotes(1) = ""

TblMaster.Seek "=", Fyr.Text, Fmark.Text, Fring.Text 'Find the required record.
If TblMaster.NoMatch Then
    MsgBox "Bird NOT found", 4, "ERROR"
    MousePointer = 0
    Unload fmPedg
    Screen.MousePointer = 0
    Exit Sub
End If

'store info on bird u are getting pedigree for
Mstatus(1) = True
mYr(1) = TblMaster("yr")
mMark(1) = TblMaster("mark")
mRing(1) = TblMaster("ringno")
Mcolcode$(1) = TblMaster("codecolour")
If IsNull(TblMaster("codecolour")) Then
    mColour$(1) = "Colour unknown"
Else
    mColour$(1) = TblMaster("codecolour")
    TblColour.Seek "=", TblMaster("codecolour")
    If TblColour.NoMatch Then
    Else
        mColour$(1) = TblColour("desc")
    End If
End If
mEyeSign$(1) = TblMaster("codeeye") & ""

If TblMaster("codesex") = "H" Then
    mSex$(1) = "Hen"
ElseIf TblMaster("codesex") = "C" Then
    mSex(1) = "Cock"
Else
    mSex(1) = "Unsexed"
End If
mBredBy(1) = TblMaster("bredby") & ""
If IsNull(TblMaster("codebline")) Then
    If IsNull(TblMaster("bloodline")) Then
        Mbloodline(1) = ""
    Else
        Mbloodline(1) = TblMaster("bloodline")
    End If
Else
    Mbloodline(1) = TblMaster("codebline")
    TblBLine.Seek "=", TblMaster("codebline")
    If TblBLine.NoMatch Then
    Else
        If TblBLine("desc") = "Not Available" Or TblBLine("desc") = "NOT AVAILABLE" Then
            Mbloodline(1) = ""
        Else
            Mbloodline(1) = TblBLine("desc")
        End If
    End If
End If
mName(1) = TblMaster("birdname") & ""
mNotes(1) = TblMaster("notes") & ""

fDetails.Caption = mColour$(1) & " " & mSex$(1) & " " & mName$(1)
fColour.Text = mColour$(1)
fName.Text = mName$(1)
fBredBy.Text = mBredBy$(1)
fDNotes.Text = mNotes$(1)
fSex.Text = mSex$(1)
If IsNull(TblMaster("dDOB")) Then
    fDOB.Text = ""
Else
    fDOB.Text = Format(TblMaster("dDOB"), "ddddd")
End If
fBloodline.Text = Mbloodline(1)
frmDetails.Caption = "Pedigree for " & Str(mYr(1)) & "/" & mMark(1) & "/" & mRing(1)

'Store Sire in bird(2) Screen field (0)
If IsNull(TblMaster("sireyr")) Then
    mYr(2) = Val("")
Else
    mYr(2) = TblMaster("sireyr")
End If
If IsNull(TblMaster("siremark")) Then
    mMark(2) = ""
Else
    mMark(2) = TblMaster("siremark")
End If
If IsNull(TblMaster("sireringno")) Then
    mRing(2) = ""
Else
    mRing(2) = TblMaster("sireringno")
End If
FbirdYMR(0).Text = Trim(Mcolcode$(2)) & " " & Trim(Right$(Trim(Str$(mYr(2))), 2) & " " & Trim(mMark(2)) & " " & Trim(mRing(2)))
ButtSire.Caption = Trim(Mcolcode$(2)) & " " & Trim(Right$(Trim(Str$(mYr(2))), 2) & " " & Trim(mMark(2)) & " " & Trim(mRing(2)))

'Dam in bird(3) Screen field (1)
If IsNull(TblMaster("damyr")) Then
    mYr(3) = Val("")
Else
    mYr(3) = TblMaster("damyr")
End If
If IsNull(TblMaster("dammark")) Then
    mMark(3) = ""
Else
    mMark(3) = TblMaster("dammark")
End If
If IsNull(TblMaster("damringno")) Then
    mRing(3) = ""
Else
    mRing(3) = TblMaster("damringno")
End If
FbirdYMR(1).Text = Trim(Right$(Trim(Str$(mYr(3))), 2) & " " & Trim(mMark(3)) & " " & Trim(mRing(3)) & " " & Trim(Mcolcode(3)))
ButtDam.Caption = Trim$(Right$(Trim(Str$(mYr(3))), 2) & " " & Trim(mMark(3)) & " " & Trim(mRing(3)) & " " & Trim(Mcolcode$(3)) & " " & Trim(Mcolcode(3)))

Mcnt% = 3 'Record found
Mcnt2% = 1 'Parents of record found, start at 2
Mcnt3% = 1   'screen fields start at 2, already have sire(0) & dam(1)

'Get all breeding info & put in an array
Do While Mcnt% < 31 'was 31
    Mcnt2% = Mcnt2% + 1
    If Mstatus(Int(Mcnt2% / 2)) = True Then      '?((Mcnt2/2)-1)
        TblMaster.Seek "=", mYr(Mcnt2%), mMark(Mcnt2%), mRing(Mcnt2%) 'Find the required record.
        If TblMaster.NoMatch Then
            Mcnt% = Mcnt% + 2
            Mcnt3% = Mcnt3% + 2
        Else
            'Key YMR already stored as we used it to find record, from parent details
            Mstatus(Mcnt2%) = True
            Mcolcode(Mcnt2%) = TblMaster("codecolour")
            If IsNull(TblMaster("codecolour")) Then
                mColour(Mcnt2%) = "Colour unknown"
            Else
                mColour(Mcnt2%) = TblMaster("codecolour")
                TblColour.Seek "=", TblMaster("codecolour")
                If TblColour.NoMatch Then
                Else
                mColour(Mcnt2%) = TblColour("desc")
                End If
            End If
            If IsNull(TblMaster("codeeye")) Then
                mEyeSign(Mcnt2%) = ""
            Else
                mEyeSign(Mcnt2%) = TblMaster("codeeye")
            End If
            If IsNull(TblMaster("codesex")) Then
                mSex(Mcnt2%) = ""
            Else
                mSex(Mcnt2%) = TblMaster("codesex")
            End If
            If IsNull(TblMaster("bredby")) Then
                mBredBy(Mcnt2%) = ""
            Else
                mBredBy(Mcnt2%) = TblMaster("bredby")
            End If
            If IsNull(TblMaster("codebline")) Then
                If IsNull(TblMaster("bloodline")) Then
                    Mbloodline(Mcnt2%) = "" '"Bloodline - Not Available"
                Else
                    Mbloodline(Mcnt2%) = TblMaster("bloodline")
                End If
            Else
                Mbloodline(Mcnt2%) = TblMaster("codebline")
                TblBLine.Seek "=", TblMaster("codebline")
                If TblBLine.NoMatch Then
                Else
                    If TblBLine("desc") = "Not Available" Or TblBLine("desc") = "NOT AVAILABLE" Then
                        Mbloodline(Mcnt2%) = ""
                    Else
                        Mbloodline(Mcnt2%) = TblBLine("desc")
                    End If
                End If
            End If
            mName(Mcnt2%) = TblMaster("birdname") & ""
            mNotes(Mcnt2%) = TblMaster("notes") & ""
            
            'Record current records SIRE
            Mcnt% = Mcnt% + 1
            If IsNull(TblMaster("sireyr")) Then
                mYr(Mcnt%) = Val("")
            Else
                mYr(Mcnt%) = TblMaster("sireyr")
            End If
            mMark(Mcnt%) = TblMaster("siremark") & ""
            mRing(Mcnt%) = TblMaster("sireringno") & ""
            Mcnt3% = Mcnt3% + 1
            If Mcnt3% < 30 Then 'Fill only the 0 - 29 screen positions
                FbirdYMR(Mcnt3%).Text = Trim$(Right$(Trim(Str$(mYr(Mcnt%))), 2) & " " & Trim(mMark(Mcnt%)) & " " & Trim(mRing(Mcnt%)) & " " & Trim(Mcolcode(Mcnt%)))
                If Mcnt3% < 14 Then  'only Bird + 14 ancestors Bird(0) to Bird (14)
                    Fbirdbb(Mcnt3%).Text = Trim(mBredBy(Mcnt%))
                End If
            End If
            'Record current record DAM
            Mcnt% = Mcnt% + 1
            If IsNull(TblMaster("damyr")) Then
                mYr(Mcnt%) = Val("")
            Else
                mYr(Mcnt%) = TblMaster("damyr")
            End If
            mMark(Mcnt%) = TblMaster("dammark") & ""
            mRing(Mcnt%) = TblMaster("damringno") & ""
            Mcnt3% = Mcnt3% + 1
            If Mcnt3% < 30 Then 'Fill only the 0 - 29 screen positions
                FbirdYMR(Mcnt3%).Text = Trim(Mcolcode$(Mcnt%)) & " " & Right$(Trim(Str$(mYr(Mcnt%))), 2) & " " & Trim(mMark(Mcnt%)) & " " & Trim(mRing(Mcnt%))
            End If
        End If
    Else
        Mcnt% = Mcnt% + 2
        Mcnt3% = Mcnt3% + 2
    End If
Loop

'Display Bred By to Screen
Mcnt2% = 2
For Mcnt% = 0 To 13 Step 1
    Fbirdbb(Mcnt%).Text = Mbloodline(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Display Ring Nos to Screen
ButtSire.Caption = Trim$(Right$(Trim(Str$(mYr(2))), 2) & " " & Trim(mMark(2)) & " " & Trim(mRing(2)) & " " & Trim(Mcolcode(2)))
ButtDam.Caption = Trim$(Right$(Trim(Str$(mYr(3))), 2) & " " & Trim(mMark(3)) & " " & Trim(mRing(3)) & " " & Trim(Mcolcode(3)))

Mcnt2% = 2
For Mcnt% = 0 To 29 Step 1
    FbirdYMR(Mcnt%).Text = Trim$(Right$(Trim(Str$(mYr(Mcnt2%))), 2) & " " & Trim(mMark(Mcnt2%)) & " " & Trim(mRing(Mcnt2%)) & " " & Trim(Mcolcode(Mcnt2%)))
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Display Name to Screen
Mcnt2% = 2
For Mcnt% = 0 To 13 Step 1
    Panel3D2(Mcnt%).Caption = mName(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

Mcnt2% = 2
For Mcnt% = 0 To 5 Step 1
    Panel3D3(Mcnt%).Caption = mBredBy(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

Fyr.Text = Trim(mYr(1))
Fmark.Text = Trim(mMark(1))
Fring.Text = Trim(mRing(1))
fNotes(0) = mNotes(2)
fNotes(1) = mNotes(3)

'Determine if Sire & Dam record exist, if not remove button
TblMaster.Seek "=", mYr(2), mMark(2), mRing(2) 'Find the required record.
If TblMaster.NoMatch Then
    ButtSire.Visible = False
End If
TblMaster.Seek "=", mYr(3), mMark(3), mRing(3) 'Find the required record.
If TblMaster.NoMatch Then
    ButtDam.Visible = False
End If

End Sub

Sub Disp_PedgN ()

Dim ABIRD As String  'Yr/Mark/RingNo of bird being analysed - colours screen printing
Dim BBird As String  'Yr/Mark/RingNo of bird being analysed - colours screen printing
Dim CntColour As Integer

ReDim mYr(64)
ReDim mMark(64)
ReDim mRing(64)
ReDim mColour(64)
ReDim Mcolcode(64)
ReDim mEyeSign(64)
ReDim mSex(64)
ReDim mBredBy(64)
ReDim mBredFor(64)
ReDim Mbloodline(64)
ReDim mName(64)
ReDim mNotes(64)
ReDim Mstatus(64)
ReDim MPrtColour(64)
ReDim MAnalysed(64)


Minfo = -1 'Info is ON, Screen Cluttered

'Housekeeping clear memory variables
Mcnt% = 0
Do While Mcnt% < 65
    mYr%(Mcnt%) = 0
    mMark$(Mcnt%) = ""
    Mcolcode$(Mcnt%) = ""
    mColour$(Mcnt%) = ""
    mEyeSign$(Mcnt%) = ""
    mSex$(Mcnt%) = ""
    mRing$(Mcnt%) = ""
    mBredBy$(Mcnt%) = ""
    mBredFor$(Mcnt%) = ""
    Mbloodline$(Mcnt%) = ""
    mNotes$(Mcnt%) = ""
    mName$(Mcnt%) = ""
    MPrtColour(Mcnt%) = 0
    Mstatus%(Mcnt%) = False
    MPrtColour(Mcnt%) = 0
    MAnalysed(Mcnt%) = False
    Mcnt% = Mcnt% + 1
Loop

'Housekeeping clear Screen fields
Mcnt% = 0
Do While Mcnt% < 30  'Bird + 30 ancestors Bird(0) to Bird(30)
    FbirdYMR(Mcnt%).Text = ""
    Mcnt% = Mcnt% + 1
Loop

For Mcnt% = 0 To 29 Step 1  'only Bird + 14 ancestors Bird(0) to Bird (13)
    Fbirdbb(Mcnt%).Text = ""
Next Mcnt%

For Mcnt% = 0 To 30 Step 1  'NAME only Bird + 14 ancestors Bird(0) to Bird (13)
    Panel3D2(Mcnt%).Caption = ""
Next Mcnt%

For Mcnt% = 0 To 13 Step 1   'bloodline display
    Panel3D3(Mcnt%).Caption = ""
Next Mcnt%
fNotes(0) = ""
fNotes(1) = ""

TblMaster.Seek "=", Fyr.Text, Fmark.Text, Fring.Text 'Find the required record.
If TblMaster.NoMatch Then
    MsgBox "Bird NOT found", 4, "ERROR"
    MousePointer = 0
    Unload fmPedg
    Screen.MousePointer = 0
    Exit Sub
End If

'store info on bird u are getting pedigree for
Mstatus(1) = True
mYr(1) = TblMaster("yr")
mMark(1) = TblMaster("mark")
mRing(1) = TblMaster("ringno")
Mcolcode$(1) = TblMaster("codecolour")
If IsNull(TblMaster("codecolour")) Then
    mColour$(1) = "Colour unknown"
Else
    mColour$(1) = TblMaster("codecolour")
    TblColour.Seek "=", TblMaster("codecolour")
    If TblColour.NoMatch Then
    Else
        mColour$(1) = TblColour("desc")
    End If
End If
mEyeSign$(1) = TblMaster("codeeye") & ""

If TblMaster("codesex") = "H" Then
    mSex$(1) = "Hen"
ElseIf TblMaster("codesex") = "C" Then
    mSex(1) = "Cock"
Else
    mSex(1) = "Unsexed"
End If
mBredBy(1) = TblMaster("bredby") & ""
mBredFor(1) = TblMaster("bredfor") & ""
If IsNull(TblMaster("codebline")) Then
    If IsNull(TblMaster("bloodline")) Then
        Mbloodline(1) = ""
    Else
        Mbloodline(1) = TblMaster("bloodline")
    End If
Else
    Mbloodline(1) = TblMaster("codebline")
    TblBLine.Seek "=", TblMaster("codebline")
    If TblBLine.NoMatch Then
    Else
        If TblBLine("desc") = "Not Available" Or TblBLine("desc") = "NOT AVAILABLE" Then
            Mbloodline(1) = ""
        Else
            Mbloodline(1) = TblBLine("desc")
        End If
    End If
End If
mName(1) = TblMaster("birdname") & ""
mNotes(1) = TblMaster("notes") & ""

fDetails.Caption = mColour$(1) & " " & mSex$(1) & " " & mName$(1)
fColour.Text = mColour$(1)
fName.Text = mName$(1)
fBredBy.Text = mBredBy$(1)
fDNotes.Text = mNotes$(1)
fSex.Text = mSex$(1)
If IsNull(TblMaster("dDOB")) Then
    fDOB.Text = ""
Else
    fDOB.Text = Format(TblMaster("dDOB"), "ddddd")
End If
fBloodline.Text = Mbloodline(1)
frmDetails.Caption = "Pedigree for " & Str(mYr(1)) & "/" & mMark(1) & "/" & mRing(1)

'Store Sire in bird(2) Screen field (0)
If IsNull(TblMaster("sireyr")) Then
    mYr(2) = Val("")
Else
    mYr(2) = TblMaster("sireyr")
End If
If IsNull(TblMaster("siremark")) Then
    mMark(2) = ""
Else
    mMark(2) = TblMaster("siremark")
End If
If IsNull(TblMaster("sireringno")) Then
    mRing(2) = ""
Else
    mRing(2) = TblMaster("sireringno")
End If
FbirdYMR(0).Text = Trim(Mcolcode$(2)) & " " & Trim(Right$(Trim(Str$(mYr(2))), 2) & " " & Trim(mMark(2)) & " " & Trim(mRing(2)))
ButtSire.Caption = Trim(Mcolcode$(2)) & " " & Trim(Right$(Trim(Str$(mYr(2))), 2) & " " & Trim(mMark(2)) & " " & Trim(mRing(2)))

'Dam in bird(3) Screen field (1)
If IsNull(TblMaster("damyr")) Then
    mYr(3) = Val("")
Else
    mYr(3) = TblMaster("damyr")
End If
If IsNull(TblMaster("dammark")) Then
    mMark(3) = ""
Else
    mMark(3) = TblMaster("dammark")
End If
If IsNull(TblMaster("damringno")) Then
    mRing(3) = ""
Else
    mRing(3) = TblMaster("damringno")
End If
FbirdYMR(1).Text = Trim(Right$(Trim(Str$(mYr(3))), 2) & " " & Trim(mMark(3)) & " " & Trim(mRing(3)) & " " & Trim(Mcolcode(3)))
ButtDam.Caption = Trim$(Right$(Trim(Str$(mYr(3))), 2) & " " & Trim(mMark(3)) & " " & Trim(mRing(3)) & " " & Trim(Mcolcode$(3)) & " " & Trim(Mcolcode(3)))

Mcnt% = 3 'Record found
Mcnt2% = 1 'Parents of record found, start at 2
Mcnt3% = 1   'screen fields start at 2, already have sire(0) & dam(1)

'Get all breeding info & put in an array
Do While Mcnt% < 62 'was 31
    Mcnt2% = Mcnt2% + 1
    If Mstatus(Int(Mcnt2% / 2)) = True Then      '?((Mcnt2/2)-1)
        TblMaster.Seek "=", mYr(Mcnt2%), mMark(Mcnt2%), mRing(Mcnt2%) 'Find the required record.
        If TblMaster.NoMatch Then
            Mcnt% = Mcnt% + 2
            Mcnt3% = Mcnt3% + 2
        Else
            'Key YMR already stored as we used it to find record, from parent details
            Mstatus(Mcnt2%) = True
            Mcolcode(Mcnt2%) = TblMaster("codecolour")
            If IsNull(TblMaster("codecolour")) Then
                mColour(Mcnt2%) = "Colour unknown"
            Else
                mColour(Mcnt2%) = TblMaster("codecolour")
                TblColour.Seek "=", TblMaster("codecolour")
                If TblColour.NoMatch Then
                Else
                mColour(Mcnt2%) = TblColour("desc")
                End If
            End If
            If IsNull(TblMaster("codeeye")) Then
                mEyeSign(Mcnt2%) = ""
            Else
                mEyeSign(Mcnt2%) = TblMaster("codeeye")
            End If
            If IsNull(TblMaster("codesex")) Then
                mSex(Mcnt2%) = ""
            Else
                mSex(Mcnt2%) = TblMaster("codesex")
            End If
            If IsNull(TblMaster("bredby")) Then
                mBredBy(Mcnt2%) = ""
            Else
                mBredBy(Mcnt2%) = TblMaster("bredby")
            End If
            If IsNull(TblMaster("bredfor")) Then
                mBredFor(Mcnt2%) = ""
            Else
                mBredFor(Mcnt2%) = TblMaster("bredfor")
            End If
            If IsNull(TblMaster("codebline")) Then
                If IsNull(TblMaster("bloodline")) Then
                    Mbloodline(Mcnt2%) = "" '"Bloodline - Not Available"
                Else
                    Mbloodline(Mcnt2%) = TblMaster("bloodline")
                End If
            Else
                Mbloodline(Mcnt2%) = TblMaster("codebline")
                TblBLine.Seek "=", TblMaster("codebline")
                If TblBLine.NoMatch Then
                Else
                    If TblBLine("desc") = "Not Available" Or TblBLine("desc") = "NOT AVAILABLE" Then
                        Mbloodline(Mcnt2%) = ""
                    Else
                        Mbloodline(Mcnt2%) = TblBLine("desc")
                    End If
                End If
            End If
            mName(Mcnt2%) = TblMaster("birdname") & ""
            mNotes(Mcnt2%) = TblMaster("notes") & ""
            Print "Name:" + mName(Mcnt2%)
            'Record current records SIRE
            Mcnt% = Mcnt% + 1
            If IsNull(TblMaster("sireyr")) Then
                mYr(Mcnt%) = Val("")
            Else
                mYr(Mcnt%) = TblMaster("sireyr")
            End If
            mMark(Mcnt%) = TblMaster("siremark") & ""
            mRing(Mcnt%) = TblMaster("sireringno") & ""
            Mcnt3% = Mcnt3% + 1
            If Mcnt3% < 30 Then 'Fill only the 0 - 29 screen positions
                FbirdYMR(Mcnt3%).Text = Trim$(Right$(Trim(Str$(mYr(Mcnt%))), 2) & " " & Trim(mMark(Mcnt%)) & " " & Trim(mRing(Mcnt%)) & " " & Trim(Mcolcode(Mcnt%)))
                'OLD If Mcnt3% < 14 Then  'only Bird + 14 ancestors Bird(0) to Bird (14)
                If Mcnt3% < 30 Then  'only Bird + 14 ancestors Bird(0) to Bird (14)
                    Fbirdbb(Mcnt3%).Text = Trim(mBredBy(Mcnt%))
                End If
            End If
            'Record current record DAM
            Mcnt% = Mcnt% + 1
            If IsNull(TblMaster("damyr")) Then
                mYr(Mcnt%) = Val("")
            Else
                mYr(Mcnt%) = TblMaster("damyr")
            End If
            mMark(Mcnt%) = TblMaster("dammark") & ""
            mRing(Mcnt%) = TblMaster("damringno") & ""
            Mcnt3% = Mcnt3% + 1
            If Mcnt3% < 30 Then 'Fill only the 0 - 29 screen positions
                FbirdYMR(Mcnt3%).Text = Trim(Mcolcode$(Mcnt%)) & " " & Right$(Trim(Str$(mYr(Mcnt%))), 2) & " " & Trim(mMark(Mcnt%)) & " " & Trim(mRing(Mcnt%))
            End If
        End If
    Else
        Mcnt% = Mcnt% + 2
        Mcnt3% = Mcnt3% + 2
    End If
Loop

'Establish common ancestors and highlight with a common colour
'read each bird
'Read all other records for that same ring no, if any found give a colour other than white
Mcnt2% = 1
CntColour% = 0

'Make sure every bird is analysed
For Mcnt2% = 1 To 30
    Do While Not MAnalysed(Mcnt2%)
        MAnalysed(Mcnt2%) = True
        MPrtColour(Mcnt2%) = 0
        ABIRD = Trim(CStr(mYr(Mcnt2%))) & Trim(mMark(Mcnt2%) & "") & Trim(mRing(Mcnt2%) & "")
    
        For Mcnt% = 1 To 30
            BBird = Trim(CStr(mYr(Mcnt%))) & Trim(mMark(Mcnt%) & "") & Trim(mRing(Mcnt%) & "")
            If ABIRD = BBird Then
                 If Mcnt% <> Mcnt2% And Trim(ABIRD) <> "0" Then
                     If MPrtColour(Mcnt2%) = 0 Then
                        CntColour% = CntColour% + 1
                        MPrtColour(Mcnt2%) = CntColour%
                     End If
                     MPrtColour(Mcnt) = MPrtColour(Mcnt2%)
                     MAnalysed(Mcnt) = True
                 End If
            End If
        Next Mcnt%
    Loop
Next Mcnt2%

'Display Bred By to Screen
Mcnt2% = 2
For Mcnt% = 0 To 29 Step 1
    Fbirdbb(Mcnt%).Text = Mbloodline(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Display Ring Nos to Screen
ButtSire.Caption = Trim$(Right$(Trim(Str$(mYr(2))), 2) & " " & Trim(mMark(2)) & " " & Trim(mRing(2)) & " " & Trim(Mcolcode(2)))
ButtDam.Caption = Trim$(Right$(Trim(Str$(mYr(3))), 2) & " " & Trim(mMark(3)) & " " & Trim(mRing(3)) & " " & Trim(Mcolcode(3)))

Mcnt2% = 2
For Mcnt% = 0 To 29 Step 1
    FbirdYMR(Mcnt%).Text = Trim$(Right$(Trim(Str$(mYr(Mcnt2%))), 2) & " " & Trim(mMark(Mcnt2%)) & " " & Trim(mRing(Mcnt2%)) & " " & Trim(Mcolcode(Mcnt2%)))
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Display Name to Screen
Mcnt2% = 2
For Mcnt% = 0 To 29 Step 1
    Panel3D2(Mcnt%).Caption = mName(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

Mcnt2% = 2
For Mcnt% = 0 To 5 Step 1
    Panel3D3(Mcnt%).Caption = mBredBy(Mcnt2%)
    fNotes(Mcnt%).Text = mNotes(Mcnt2%)
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

'Do Background colour for common birds
Mcnt2% = 2
For Mcnt% = 0 To 29 Step 1
    If MPrtColour(Mcnt2%) > 0 Then
        Select Case MPrtColour(Mcnt2%)
              Case 1
                Panel3D2(Mcnt%).BackColor = vbYellow
                FbirdYMR(Mcnt%).BackColor = vbYellow
              Case 2
                Panel3D2(Mcnt%).BackColor = vbBlue
                FbirdYMR(Mcnt%).BackColor = vbBlue
              Case 3
                Panel3D2(Mcnt%).BackColor = vbRed
                FbirdYMR(Mcnt%).BackColor = vbRed
              Case 4
                Panel3D2(Mcnt%).BackColor = vbGreen
                FbirdYMR(Mcnt%).BackColor = vbGreen
              Case 5
                Panel3D2(Mcnt%).BackColor = vbMagenta
                FbirdYMR(Mcnt%).BackColor = vbMagenta
              Case 6
                Panel3D2(Mcnt%).BackColor = &HFFC0FF
                FbirdYMR(Mcnt%).BackColor = &HFFC0FF
              Case 7
                Panel3D2(Mcnt%).BackColor = &HFFFF80
                FbirdYMR(Mcnt%).BackColor = &HFFFF80
              Case 8
                Panel3D2(Mcnt%).BackColor = &H40C0&
                FbirdYMR(Mcnt%).BackColor = &H40C0&
              Case Else
                Panel3D2(Mcnt%).BackColor = vbWhite
                FbirdYMR(Mcnt%).BackColor = vbWhite
        End Select
    Else
                Panel3D2(Mcnt%).BackColor = &HE0E0E0
                FbirdYMR(Mcnt%).BackColor = &H80000005
    End If
    Mcnt2% = Mcnt2% + 1
Next Mcnt%

Fyr.Text = Trim(mYr(1))
Fmark.Text = Trim(mMark(1))
Fring.Text = Trim(mRing(1))
fNotes(0) = mNotes(2)
fNotes(1) = mNotes(3)

'Determine if Sire & Dam record exist, if not remove button
TblMaster.Seek "=", mYr(2), mMark(2), mRing(2) 'Find the required record.
If TblMaster.NoMatch Then
    ButtSire.Visible = False
End If
TblMaster.Seek "=", mYr(3), mMark(3), mRing(3) 'Find the required record.
If TblMaster.NoMatch Then
    ButtDam.Visible = False
End If
End Sub

Sub Form_Activate ()
    fNotes(0).Enabled = True
    fNotes(1).Enabled = True
    MOyr = Fyr.Text
    MOmark = Fmark.Text
    MOring = Fring.Text
    Disp_PedgN
    fNotes(0).Enabled = False
    fNotes(1).Enabled = False
End Sub

Sub Form_Load ()
Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
Set TblMaster = DB.OpenTable("master")     'Open Table
TblMaster.Index = "bird_ymr"              'Open Index
'Set TblPedg = DB.OpenTable("pedg")     'Open Table
'Set TblColour = DB.OpenTable("colour")     'Open Table
Set TblColour = DB.OpenTable("colour")     'Open Table
TblColour.Index = "colour"              'Open Index
Set TblBLine = DB.OpenTable("Bloodline")     'Open Table
TblBLine.Index = "Bloodline"              'Open Index
End Sub

Sub Form_Unload (Cancel As Integer)
    TblMaster.Close
    TblColour.Close
    TblBLine.Close
    DB.Close
    Set TblMaster = Nothing
    Set TblColour = Nothing
    Set TblBLine = Nothing
    Set DB = Nothing
End Sub

Sub FrmDetails_Click ()
frmDetails.Visible = False
End Sub

