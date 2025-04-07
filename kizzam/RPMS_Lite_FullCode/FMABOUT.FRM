VERSION 2.00
Begin Form FMabout 
   Caption         =   "About"
   ClientHeight    =   6570
   ClientLeft      =   1050
   ClientTop       =   1530
   ClientWidth     =   8415
   Height          =   7080
   Left            =   990
   LinkTopic       =   "Form4"
   ScaleHeight     =   6570
   ScaleWidth      =   8415
   Top             =   1080
   Width           =   8535
   Begin SSPanel Panel3D2 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6240
      Width           =   8415
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8415
      Begin SSCommand ButtPrtOrder 
         Caption         =   "Print &Order Form"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2160
         TabIndex        =   2
         Top             =   5520
         Width           =   1815
      End
      Begin SSCommand cmdPrtSys 
         Caption         =   "&Print System"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4080
         TabIndex        =   1
         Top             =   5520
         Width           =   1935
      End
      Begin SSCommand ButtOK 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6120
         TabIndex        =   0
         Top             =   5520
         Width           =   1935
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Demo Details"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   3375
         Left            =   600
         ShadowStyle     =   1  'Raised
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   6855
         Begin Label Label5 
            Caption         =   "Label5"
            Height          =   2295
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin Label Label4 
            Height          =   2295
            Left            =   360
            TabIndex        =   19
            Top             =   480
            Width           =   6135
         End
      End
      Begin SSCommand cmdDemo 
         Caption         =   "Demo Details"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   5520
         Width           =   1695
      End
      Begin SSPanel Panel3D10 
         Alignment       =   4  'Right Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Panel3D10"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   5040
         Width           =   6855
      End
      Begin SSPanel Panel3D9 
         Alignment       =   4  'Right Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Panel3D9"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   4680
         Width           =   6495
      End
      Begin SSPanel Panel3D8 
         Alignment       =   4  'Right Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Panel3D8"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   4320
         Width           =   6495
      End
      Begin SSPanel PnlDBVersion 
         Alignment       =   4  'Right Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Unknown"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   3480
         Width           =   2895
      End
      Begin SSPanel Panel3D7 
         Alignment       =   6  'Center - TOP
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Uses Access V2.0 Database"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   3000
         Width           =   2895
      End
      Begin SSPanel Panel3D6 
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Panel3D6"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   3720
         TabIndex        =   12
         Top             =   2520
         Width           =   3495
      End
      Begin SSPanel Panel3D5 
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "Panel3D5"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   2040
         Width           =   3495
      End
      Begin SSPanel Panel3D4 
         Alignment       =   4  'Right Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "This software is licensed to:"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   2520
         Width           =   2895
      End
      Begin SSPanel Panel3D3 
         Alignment       =   4  'Right Justify - MIDDLE
         BackColor       =   &H00E0E0E0&
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         Caption         =   "License No.:"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
      End
      Begin Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label6"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   4815
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   8175
         WordWrap        =   -1  'True
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Order from: Kieran Malone Ph +61 (0)409495571, email rpms2002@hotmail.com"
         ForeColor       =   &H00004080&
         Height          =   975
         Left            =   5760
         TabIndex        =   8
         Top             =   120
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copyright 1990-2014 by Kieran G Malone, All Rights Reserved"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1560
         Width           =   6375
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "The Racing Pigeon Management System Lite V3.3f May 2019"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   1320
         TabIndex        =   6
         Top             =   120
         Width           =   4215
      End
   End
End
Option Explicit

Sub ButtOK_Click ()
    Frame3d1.Visible = False
    Unload FMabout
End Sub

Sub ButtOK_KeyDown (KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_F12 Then    ' Display key combinations.
	label6.Visible = True
    End If
End Sub

Sub ButtOK_KeyUp (KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_F12 Then    ' Display key combinations.
	label6.Visible = False
    End If
End Sub

Sub ButtPrtOrder_Click ()
    PrtOrder
End Sub

Sub cmdDemo_Click ()
Dim a As String
If Frame3d1.Visible = True Then
    Frame3d1.Visible = False
Else
    Frame3d1.Visible = True
    a$ = "If the User is 'DEMO' then the system is running in a "
    a$ = a$ & "Demonstration or Trial mode. This special mode provides"
    a$ = a$ & " most functionality to give the user an idea of what he can get"
    a$ = a$ & " from the system. NOTE however that the software limits the user"
    a$ = a$ & " to 45 bird records."
    label4.Caption = a$
End If

End Sub

Sub cmdDemo_KeyDown (KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown, Txt

    Const KEY_F2 = &H71 ' Define constants.
    Dim a As String

    If KeyCode = KEY_F2 Then    ' Display key combinations.
      If label5.Visible = True Then
	label5.Visible = False
      Else
	a$ = "Version Details:" & Chr(10)
	a$ = a$ & "Ver 3.2z(a) - Race Sheet Style 12"
	label5.Caption = a$
	label5.Visible = True
      End If
    End If
End Sub

Sub cmdPrtSys_Click ()
prtsys
End Sub

Sub Form_Activate ()
    screen.MousePointer = 0
End Sub

Sub Form_Load ()
    
    Dim smOwner As String

    smOwner$ = GLOwner$ & Chr(10) & GLAddress1$ & Chr(10) & GLAddress2$ & Chr(10) & GLCity$
    smOwner$ = smOwner$ & "   " & GLState$ & "   " & GLPostcode$
    smOwner$ = smOwner$ & Chr(10) & "Ph: " & GLPhoneno$
    smOwner$ = smOwner$ & Chr(10) & "Club Mark: " & GLMark
    Panel3D5.Caption = LicenseDisplay(GLSerialNo$)
    Panel3D6.Caption = smOwner$
    PnlDBVersion.Caption = GLDBVersion$
    Panel3D8.Caption = "Database in use is " & DBFullFileName$
    Panel3D9.Caption = "Windows directory is " & WinPath$
    Panel3D10.Caption = ""
    Frame3d1.Visible = False
    label5 = "V3.3d - RPMS 2016 Lite"
    label5 = "V3.3f - RPMS 2019 Lite"
    label6.Caption = "May 2019 V3.3f - Update Dropdowns to include current years for selection" & Chr(10)
    label6.Caption = label6.Caption & "Dec 2015 - Remove CR LF TAB Pedigree Notes" & Chr(10)
    label6.Caption = label6.Caption & "Mar 2014 - Updated Pedg Screen, highlight common ancestors" & Chr(10)
    'label6.Caption = label6.Caption & "Feb 2014 - New Analysis reverse Pedigree from Bird Register screen" & Chr(10)
    label6.Caption = label6.Caption & "Feb 2014 - New Analysis reverse Pedigree from Bird Register screen" & Chr(10)
    label6.Caption = label6.Caption & "Feb 2014 - Major Enhancement to printed Pedigree to incl. more notes on birds" & Chr(10)
    label6.Caption = label6.Caption & "Jan 2014 - Major Enhancement to incl. more pedigree info on screen" & Chr(10)
    label6.Caption = label6.Caption & "Dec 2013 - Adjusted screen sizing, DropDown Years updated" & Chr(10)
    label6.Caption = label6.Caption & "Previous V3.2z(a) - Additional Race Entry Sheet Only - Style 12" & Chr(10)
    label6.Caption = label6.Caption & "31/7/2001 - Fixed error which allowed user to attempt to add birds to team when no team existed" & Chr(10)
    label6.Caption = label6.Caption & "31/7/2001 - Worked on adding race history (map)" & Chr(10)
    label6.Caption = label6.Caption & "31/7/2001 - Added race entry sheet for Central Highlands Pigeon combine" & Chr(10)
    label6.Caption = label6.Caption & "09/6/2002 - Fixed problem with adding Race Point, enhanced Lng/Lat entry validation." & Chr(10)
    label6.Visible = False
End Sub

Sub Panel3D1_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    label6.Visible = True
End Sub

Sub Panel3D1_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    label6.Visible = False
End Sub

