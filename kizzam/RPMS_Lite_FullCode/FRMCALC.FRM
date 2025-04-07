VERSION 2.00
Begin Form frmCalc 
   Caption         =   "Race Calculator"
   ClientHeight    =   5415
   ClientLeft      =   1185
   ClientTop       =   1065
   ClientWidth     =   6225
   Height          =   5820
   Left            =   1125
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   6225
   Top             =   720
   Width           =   6345
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin TextBox ft_sec 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   21
         Top             =   3720
         Width           =   495
      End
      Begin TextBox ft_Min 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   20
         Top             =   3720
         Width           =   375
      End
      Begin TextBox ft_hr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   3720
         Width           =   375
      End
      Begin TextBox arrival 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   4080
         Width           =   2295
      End
      Begin TextBox flying 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   3360
         Width           =   2295
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Enter Data here:"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00008000&
         Height          =   2175
         Left            =   360
         ShadowStyle     =   1  'Raised
         TabIndex        =   8
         Top             =   960
         Width           =   5415
         Begin CSComboBox CSCombo2 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FRMCALC.FRX:0000
            Contents        =   FRMCALC.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   315
            Left            =   2520
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   "pntname,DistCalcGEODyds,DistSpec"
            ListDataSource  =   "Data1"
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   24
            Text            =   "CSCombo2"
            TextDisplay     =   0  'first visible
            Top             =   360
            ValueCol        =   -1
            Visible         =   0   'False
            Width           =   2775
         End
         Begin CSComboBox CSCombo1 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FRMCALC.FRX:0042
            Contents        =   FRMCALC.FRX:0082
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   315
            Left            =   2520
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   "pntname,DistCalcGEOD,DistSpec"
            ListDataSource  =   "Data1"
            SearchString    =   ""
            StartSearch     =   0
            TabIndex        =   23
            Text            =   "CSCombo1"
            TextDisplay     =   0  'first visible
            Top             =   360
            ValueCol        =   -1
            Visible         =   0   'False
            Width           =   2775
         End
         Begin MaskEdBox release 
            Height          =   375
            Left            =   2520
            Mask            =   "##:##"
            MaxLength       =   5
            PromptChar      =   "_"
            TabIndex        =   3
            Top             =   1560
            Width           =   1815
         End
         Begin MaskEdBox Velocity 
            Format          =   "###0"
            Height          =   285
            Left            =   2520
            PromptChar      =   "_"
            TabIndex        =   2
            Top             =   1200
            Width           =   1815
         End
         Begin MaskEdBox Distance 
            Format          =   "########0.0000"
            Height          =   285
            Left            =   2520
            PromptChar      =   "_"
            TabIndex        =   1
            Top             =   840
            Width           =   1815
         End
         Begin Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Race Point:"
            Height          =   255
            Left            =   960
            TabIndex        =   10
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Release Time:"
            Height          =   255
            Left            =   720
            TabIndex        =   13
            Top             =   1560
            Width           =   1695
         End
         Begin Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Expected Velocity:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   2295
         End
         Begin Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Distance:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   2295
         End
      End
      Begin SSCommand cmdExit 
         Caption         =   "E&xit"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   4560
         Width           =   1575
      End
      Begin SSCommand cmdCalc 
         Caption         =   "&Calculate"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   4560
         Width           =   1575
      End
      Begin Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Expected Flying Time:"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   3720
         Width           =   2775
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Secs"
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   3720
         Width           =   495
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mins"
         Height          =   255
         Left            =   4560
         TabIndex        =   17
         Top             =   3720
         Width           =   495
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   3720
         Width           =   375
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Expected Arrival Time:"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   4080
         Width           =   2175
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Expected Flying Time (in Minutes):"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   3015
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enter the Distance and expected velocity and Release time and click 'Calculate' to get Arrival time."
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   855
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   4935
      End
   End
End
Option Explicit

Sub cmdCalc_Click ()
Dim dRelTime As Variant
Dim dClkTime As Variant
Dim dFlyingTime As Variant

If distance = "" Then
    MsgBox "Must enter Numeric Distance", 16, "Please Reenter"
    distance.SetFocus
    Exit Sub
End If
If Not IsNumeric(distance) Then
    MsgBox "Must enter Numeric Distance", 16, "Please Reenter"
    distance.SetFocus
    Exit Sub
End If
If velocity = "" Then
    MsgBox "Must enter Numeric Velocity", 16, "Please Reenter"
    velocity.SetFocus
    Exit Sub
End If
If Not IsNumeric(Mid(release, 1, 2)) Then
    MsgBox "Must enter Release Time", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
If Not IsNumeric(Mid(release, 4, 2)) Then
    MsgBox "Must enter Release Time", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
If Val(Mid(release, 1, 2)) > 23 Then
    MsgBox "Hour must be less then 24", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
If Val(Mid(release, 3, 2)) > 59 Then
    MsgBox "Minute must be less than 60", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
If Val(Mid(release, 1, 2)) < 0 Then
    MsgBox "Hour must be 0 to 24", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
If Val(Mid(release, 3, 2)) < 0 Then
    MsgBox "Minute must be 0 to 59", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
If Format(release, "ttttt") = "__:__" Then
    MsgBox "Please provide release time", 16, "Please Reenter"
    release.SetFocus
    Exit Sub
End If
dRelTime = Format(release, "ttttt")
If GLDistType = "Y" Then
    dFlyingTime = distance * 1760 / velocity
Else
    dFlyingTime = distance * 1000 / velocity
End If
dClkTime = DateAdd("n", dFlyingTime, dRelTime)
ft_hr.Text = Int(dFlyingTime / 60)
ft_min = Int(dFlyingTime - (Int(dFlyingTime / 60) * 60))
ft_sec = (dFlyingTime - Int(dFlyingTime)) * 60
Flying = dFlyingTime
arrival = Format(dClkTime, "hh:mm")
End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub CSCombo1_Change ()
    distance = CSCombo1.Value
End Sub

Sub CSCombo2_Change ()
    distance = CSCombo2.Value
End Sub

Sub Form_Load ()
If GLDistType = "Y" Then
    Label2.Caption = "Distance in Miles"
    Label4.Caption = "Velocity (YPM)"
'    CSCombo1.Visible = False
'    CSCombo2.Visible = True
Else
    Label2.Caption = "Distance in Kilometres"
    Label4.Caption = "Velocity (MPM)"
'    CSCombo2.Visible = False
'    CSCombo1.Visible = True
End If
End Sub

