VERSION 2.00
Begin Form FMrace 
   Caption         =   "Race Entry"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9330
   Height          =   7125
   Left            =   0
   LinkTopic       =   "Form4"
   ScaleHeight     =   6720
   ScaleWidth      =   9330
   Top             =   15
   Width           =   9450
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   6240
      Width           =   9375
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9375
      Begin SSCommand ButtBirdList 
         Caption         =   "&Nominate Birds"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   7440
         TabIndex        =   26
         Top             =   2040
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
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   7440
         TabIndex        =   0
         Top             =   120
         Width           =   1695
      End
      Begin SSCommand ButtOK 
         Caption         =   "&Save"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   7440
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
      End
      Begin SSCommand CmdPrtRaceSheet 
         Caption         =   "&Print Race Sheet"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7440
         TabIndex        =   28
         Top             =   3720
         Width           =   1695
      End
      Begin TextBox scrWinVel 
         Height          =   285
         Left            =   6240
         MaxLength       =   7
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin SSFrame Frame3D3 
         Caption         =   "Clock Details"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   56
         Top             =   5400
         Width           =   8775
         Begin TextBox scrVelocity 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin TextBox scrFT 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3960
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   360
            Width           =   2055
         End
         Begin TextBox scrCB 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   360
            Width           =   2175
         End
         Begin Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Velocity:"
            Height          =   255
            Left            =   6000
            TabIndex        =   59
            Top             =   360
            Width           =   855
         End
         Begin Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flying Time:"
            Height          =   255
            Left            =   2880
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
         Begin Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bird:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   495
         End
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FMRACE.FRX:0000
         Contents        =   FMRACE.FRX:0040
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   4680
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         TabIndex        =   4
         Text            =   "CSCombo1"
         TextDisplay     =   0  'first visible
         Top             =   1560
         ValueCol        =   1
         Width           =   1095
      End
      Begin SSFrame Frame3D2 
         Caption         =   "No Competing:"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   45
         Top             =   2640
         Width           =   7095
         Begin TextBox Text14 
            Height          =   285
            Left            =   6000
            TabIndex        =   22
            Text            =   "Text14"
            Top             =   840
            Width           =   735
         End
         Begin TextBox Text13 
            Height          =   285
            Left            =   6000
            TabIndex        =   21
            Text            =   "Text13"
            Top             =   480
            Width           =   735
         End
         Begin TextBox Text12 
            Height          =   285
            Left            =   5160
            MaxLength       =   6
            TabIndex        =   20
            Text            =   "Text12"
            Top             =   840
            Width           =   735
         End
         Begin TextBox Text11 
            Height          =   285
            Left            =   5160
            MaxLength       =   6
            TabIndex        =   19
            Text            =   "Text11"
            Top             =   480
            Width           =   735
         End
         Begin TextBox Text10 
            Height          =   285
            Left            =   4320
            MaxLength       =   6
            TabIndex        =   18
            Text            =   "Text10"
            Top             =   840
            Width           =   735
         End
         Begin TextBox Text9 
            Height          =   285
            Left            =   4320
            MaxLength       =   6
            TabIndex        =   17
            Text            =   "Text9"
            Top             =   480
            Width           =   735
         End
         Begin TextBox Text8 
            Height          =   285
            Left            =   3480
            MaxLength       =   5
            TabIndex        =   16
            Text            =   "Text8"
            Top             =   840
            Width           =   735
         End
         Begin TextBox Text7 
            Height          =   285
            Left            =   3480
            MaxLength       =   5
            TabIndex        =   15
            Text            =   "Text7"
            Top             =   480
            Width           =   735
         End
         Begin TextBox Text6 
            Height          =   285
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   14
            Text            =   "Text6"
            Top             =   840
            Width           =   615
         End
         Begin TextBox Text5 
            Height          =   285
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   13
            Text            =   "Text5"
            Top             =   480
            Width           =   615
         End
         Begin TextBox Text4 
            Height          =   285
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "Text4"
            Top             =   840
            Width           =   615
         End
         Begin TextBox Text3 
            Height          =   285
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "Text3"
            Top             =   480
            Width           =   615
         End
         Begin TextBox scrNoBirds 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   480
            Width           =   495
         End
         Begin Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Combine"
            Height          =   255
            Left            =   2640
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
         Begin Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Flyers:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   1215
         End
         Begin Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Birds:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   1215
         End
         Begin Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Loft"
            Height          =   255
            Left            =   1320
            TabIndex        =   51
            Top             =   240
            Width           =   495
         End
         Begin Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Fed."
            Height          =   255
            Left            =   3480
            TabIndex        =   50
            Top             =   240
            Width           =   495
         End
         Begin Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Int."
            Height          =   255
            Left            =   6000
            TabIndex        =   49
            Top             =   240
            Width           =   495
         End
         Begin Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Club"
            Height          =   255
            Left            =   1920
            TabIndex        =   48
            Top             =   240
            Width           =   495
         End
         Begin Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Nat."
            Height          =   255
            Left            =   5160
            TabIndex        =   47
            Top             =   240
            Width           =   495
         End
         Begin Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sect."
            Height          =   255
            Left            =   4320
            TabIndex        =   46
            Top             =   240
            Width           =   495
         End
      End
      Begin TextBox mRaceCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin SSCommand CmdResults 
         Caption         =   "Enter &Results"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   7440
         TabIndex        =   27
         Top             =   2880
         Width           =   1695
      End
      Begin TextBox Key_reldate 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   42
         Text            =   "Text7"
         Top             =   4200
         Visible         =   0   'False
         Width           =   495
      End
      Begin TextBox Key_coderace 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   41
         Text            =   "Text7"
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Race Release Time:"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   40
         Top             =   1920
         Width           =   4215
         Begin TextBox Mrelmins 
            Height          =   285
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "Text6"
            Top             =   240
            Width           =   495
         End
         Begin TextBox Mrelhrs 
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "Text2"
            Top             =   240
            Width           =   495
         End
         Begin SSOption Mindam 
            Caption         =   "AM"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2400
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin SSOption Mindpm 
            Caption         =   "PM"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3360
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Mins"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Hrs"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   255
            Left            =   720
            TabIndex        =   32
            Top             =   240
            Width           =   495
         End
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7560
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "points"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1140
      End
      Begin CSComboBox CSMcodepnt 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FMRACE.FRX:0054
         Contents        =   FMRACE.FRX:0094
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   2040
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "pntname,code_pnt"
         ListDataSource  =   "Data1"
         SearchString    =   ""
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TextDisplay     =   1  'all visible
         Top             =   1080
         ValueCol        =   1
         Width           =   4815
      End
      Begin TextBox Mcoderace 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   1560
         Width           =   1095
      End
      Begin TextBox task 
         Height          =   285
         Left            =   120
         MaxLength       =   1
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
      End
      Begin TextBox Key_codepnt 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   4560
         Visible         =   0   'False
         Width           =   495
      End
      Begin TextBox Mnotes 
         Height          =   975
         Left            =   4560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "Text4"
         Top             =   4320
         Width           =   4455
      End
      Begin TextBox Mweather 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "Text3"
         Top             =   4320
         Width           =   3975
      End
      Begin TextBox Mreldate 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1815
      End
      Begin Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Winning Velocity:"
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   2160
         Width           =   1575
      End
      Begin Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race/Toss:"
         Height          =   255
         Left            =   3480
         TabIndex        =   55
         Top             =   1560
         Width           =   1095
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Code:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notes: "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   37
         Top             =   3960
         Width           =   735
      End
      Begin Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Weather: "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   3960
         Width           =   975
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Type:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date :"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   1575
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Race Point :"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
   End
End
Option Explicit
Dim TblMaster As Table
Dim TblLoft As Table
Dim DStrans As Dynaset
Dim DB As Database
Dim GLindChange As Integer
Dim FgCntTrans As Long
Dim FlagAct As Integer 'To flag Activate run

Sub ButtBirdList_Click ()
    If GLindChange% = -1 Then    'should update record?
    End If
    FMracelist.mRaceDate.Text = mRelDate.Text
    FMracelist.Mcodepnt.Text = CSMcodepnt.Value
    FMracelist.mPntName.Text = CSMcodepnt.Text
    FMracelist.Mcoderace.Text = Mcoderace.Text
    FMracelist.mRaceCode.Text = mRaceCode.Text
    FMracelist.mRaceStatus.Text = Trim(CSCombo1.Text)
    FMracelist.Show 1
End Sub

Sub ButtCancel_Click ()
    Dim reply As Integer, zz As Integer

    If FMRace.task.Text = "A" Then
	'If RecordChangedInAdd() Then
	If GLindChange% Then ' has any fields changed
	    reply = MsgBox("Do you want to discard new record?", 4, "Warning?")
	    If reply <> 6 Then
		mRelDate.SetFocus
		Exit Sub
	    End If
	Else
	    MsgBox "Add Record Cancelled", 16, "Notification Only"
	End If
    End If
    
    If FMRace.task.Text = "M" Then
	If GLindChange% Then ' has any fields changed
	    reply = MsgBox("Do you want to cancel changes?", 4, "Warning?")
	    If reply <> 6 Then
		mRelHrs.SetFocus
		Exit Sub
	    End If
	End If
    End If

    'FMRace.Hide
    Unload Me
End Sub

Sub ButtOK_Click ()
    Dim NL As String
    NL$ = Chr$(10)

    On Error GoTo RaceErr

    'Validate Data entry
    If FMRace.task.Text = "A" Then
	If Trim(mRelDate.Text = "") Then
	    MsgBox "Release Date cannot be empty", 16, "Error: Cannot add record"
	    mRelDate.SetFocus
	    Exit Sub
	End If
	'Fix potential leap year problem otherwise won't pass isdate() check
	If mRelDate.Text = "29/2/00" Or mRelDate.Text = "29/02/00" Then
	    mRelDate.Text = DateAdd("yyyy", 100, mRelDate.Text)
	End If
	If Not IsDate(mRelDate.Text) Then
	    MsgBox "Release Date not valid", 16, "Error: Cannot add record"
	    mRelDate.SetFocus
	    Exit Sub
	End If
	'*Fix VB3 century default - Y2k
	'If Year(mRelDate.Text) < 1929 Then
	    'mRelDate.Text = DateAdd("yyyy", 100, mRelDate.Text)
	'End If
	mRelDate.Text = DateFix(mRelDate.Text)
	If Trim(CSMcodepnt.Value) = "" Or Null Then
	    MsgBox "Race Point not valid", 16, "Error: Cannot add record"
	    CSMcodepnt.SetFocus
	    Exit Sub
	End If
	If Trim(Mcoderace.Text) = "" Or Null Then
	    MsgBox "Race Code cannot be empty", 16, "Error: Cannot add record"
	    Mcoderace.SetFocus
	    Exit Sub
	End If
	If Trim(mRelHrs.Text = "") Then
	    MsgBox "Release Hours cannot be empty", 16, "Error: Cannot add record"
	    mRelHrs.SetFocus
	    Exit Sub
	End If
	If Trim(mRelMins.Text = "") Then
	    MsgBox "Release Minutes cannot be empty", 16, "Error: Cannot add record"
	    mRelMins.SetFocus
	    Exit Sub
	End If
	'etc add in all validation for Add, maybe set up in Add Val func
    End If
    If scrNoBirds.Text = "" Then
	scrNoBirds.Text = ""
    End If
    If IsNull(scrWinVel.Text) Then
	scrWinVel.Text = "0"
    End If
    If scrWinVel.Text = "" Then
	scrWinVel.Text = "0"
    End If
    If IsNumeric(scrWinVel.Text) Then
    Else
	MsgBox "Winning Velocity MUST be Numeric", 32, "Invalid Winning Velocity"
	scrWinVel.SetFocus
	Exit Sub
    End If
    If Val(scrWinVel.Text) > 4000 Then
	MsgBox "Winning Velocity over 4000, you MUST be joking!", 32, "Invalid Winning Velocity"
	scrWinVel.SetFocus
	Exit Sub
    End If

    If IsNull(Text3.Text) Then
	Text3.Text = ""
    End If
    If IsNull(Text3.Text) Then
	Text3.Text = ""
    End If
    If IsNull(Text3.Text) Then
	Text3.Text = ""
    End If
    If IsNull(Text3.Text) Then
	Text3.Text = ""
    End If
    If IsNull(Text4.Text) Then
	Text4.Text = ""
    End If
    If IsNull(Text5.Text) Then
	Text5.Text = ""
    End If
    If IsNull(Text6.Text) Then
	Text6.Text = ""
    End If
    If IsNull(Text7.Text) Then
	Text7.Text = ""
    End If
    If IsNull(Text8.Text) Then
	Text8.Text = ""
    End If
    If IsNull(Text9.Text) Then
	Text9.Text = ""
    End If
    If IsNull(Text10.Text) Then
	Text10.Text = ""
    End If
    If IsNull(Text11.Text) Then
	Text11.Text = ""
    End If
    If IsNull(Text12.Text) Then
	Text12.Text = ""
    End If
    If IsNull(Text13.Text) Then
	Text13.Text = ""
    End If
    If IsNull(Text14.Text) Then
	Text14.Text = ""
    End If
    If FMRace.task.Text = "M" Or task.Text = "A" Then
	If Trim(mRelHrs.Text = "") Then
	    MsgBox "Release Hrs MUST NOT be empty", 16, "Error: Cannot update record"
	    mRelHrs.SetFocus
	    Exit Sub
	End If
	If Trim(mRelMins.Text = "") Then
	    MsgBox "Release Mins MUST NOT be empty", 16, "Error: Cannot update record"
	    mRelMins.SetFocus
	    Exit Sub
	End If
    End If

    'Add record
    If FMRace.task.Text = "A" Then
	'TblMaster.Seek "=", Mreldate.Text, CSMcodepnt.Value, Mcoderace.Text  'Ensure unique index
	mRaceCode.Text = CLng(TblLoft("RaceCode") + 1)
	TblMaster.Seek "=", CLng(mRaceCode.Text) 'Ensure unique index
	If TblMaster.NoMatch Then
	    'Update Control to keep last Race Code used
	    TblLoft.Edit
	    TblLoft("RaceCode") = TblLoft("RaceCode") + 1
	    TblLoft.Update
	    TblMaster.AddNew
	    ScrToTbl
	    MsgBox "Record Added" & NL$ & "Click on Bird List to select birds" & NL$ & "to participate", 16, "Notification Only"
	    ButtBirdList.Enabled = True
	    FMRace.task.Text = "M"
	    FMRace.Key_reldate.Text = mRelDate.Text
	    FMRace.Key_coderace.Text = Mcoderace.Text
	    FMRace.Key_codepnt.Text = CSMcodepnt.Value
	    Exit Sub
	Else
	    MsgBox "Cannot add record, Code already used", 16, "Duplicate Key"
	    Mcoderace.SetFocus
	    Mcoderace.SelStart = 0
	    Mcoderace.SelLength = Len(Mcoderace.Text)
	    Exit Sub
	End If
    End If

    ' Update changes
    If FMRace.task.Text = "M" Then
	If GLindChange% Then
	    TblMaster.Edit
	    ScrToTbl
	    MsgBox "Record Changed", 16, "Notification Only"
	    GLindChange% = 0
	Else
	    MsgBox "No Changes effected", 16, "Notification Only"
	End If
    End If
    'FMRace.Hide
    ButtOk.Enabled = False
    ButtBirdList.Enabled = True
    CmdResults.Enabled = True
    Exit Sub

RaceErr:
    Load FMerror
    FMerror.Text1.Text = "FMrace"
    FMerror.Text2.Text = Str(Err)
    FMerror.Text3.Text = Error$(Err)
    FMerror.Show 1
    Resume Next
End Sub

Sub CmdPrtRaceSheet_Click ()
Dim vCodePnt As String
FMPrtRaceSheet.mRaceCode.Text = mRaceCode.Text
FMPrtRaceSheet.mRaceDate.Text = mRelDate.Text
vCodePnt = CSMcodepnt.Value
Data1.Recordset.FindFirst "code_pnt = '" & vCodePnt & "'"
FMPrtRaceSheet.mRacePoint.Text = Data1.Recordset("pntname")
MsgPanel.Caption = ""
FMPrtRaceSheet.Show 1
End Sub

Sub CmdResults_Click ()
FMResults.mRaceCode.Text = mRaceCode.Text
FMResults.Show 1
End Sub

Sub CSCombo1_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
    If CSCombo1.Value = "T" Then
	Text3.Enabled = False
	Text4.Enabled = False
	Text5.Enabled = False
	Text6.Enabled = False
	Text7.Enabled = False
	Text8.Enabled = False
	Text9.Enabled = False
	Text10.Enabled = False
	Text11.Enabled = False
	Text12.Enabled = False
	Text13.Enabled = False
	Text14.Enabled = False
    Else
	Text3.Enabled = True
	Text4.Enabled = True
	Text5.Enabled = True
	Text6.Enabled = True
	Text7.Enabled = True
	Text8.Enabled = True
	Text9.Enabled = True
	Text10.Enabled = True
	Text11.Enabled = True
	Text12.Enabled = True
	Text13.Enabled = True
	Text14.Enabled = True
    End If
End Sub

Sub CSMcodepnt_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub DefaultFields ()
    mRelDate.Text = Format$(Now, "ddddd")
    CSMcodepnt.Text = ""
    CSMcodepnt.Value = ""
    Mcoderace = "FLOCK"
    mRelHrs.Text = "07"
    mRelMins = "00"
    Mindam.Value = -1
    Mweather.Text = ""
    Mnotes.Text = ""
    mRaceCode.Text = ""
    GLindChange% = 0
    
    scrNoBirds.Text = 0
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    CSCombo1.Text = "Race"
    CSCombo1.Value = "R"
    scrWinVel.Text = ""
    scrCB.Text = ""
    scrFT.Text = ""
    scrVelocity.Text = ""
    End Sub

Sub Form_Activate ()

    Dim strCriterion As String
    Dim SQLstmt As String

    On Error GoTo FMerror

    Data1.Refresh
    DoEvents
    If FMRace.task.Text = "A" Then
	CmdPrtRaceSheet.Enabled = False
	CSCombo1.Enabled = True
	mRelDate.Enabled = True
	CSMcodepnt.Enabled = True
	Mcoderace.Enabled = True
	ButtBirdList.Enabled = False
	FMRace.Key_reldate.Text = ""
	FMRace.Key_codepnt.Text = ""
	FMRace.Key_coderace.Text = ""
	TblLoft.MoveFirst
	FMRace.mRaceCode.Text = TblLoft("RaceCode") + 1
	DefaultFields   'Setup Default Field Values
	CmdResults.Enabled = False
	ButtBirdList.Enabled = False
	mRelDate.SetFocus
	Exit Sub
    Else
	FlagAct = -1
	CmdPrtRaceSheet.Enabled = True
	CSCombo1.Enabled = False
	mRelDate.Enabled = False
	CSMcodepnt.Enabled = False
	Mcoderace.Enabled = False
	TblMaster.Seek "=", CLng(mRaceCode.Text)
	'Get no entries etc.
	'DB is opened in 'LOAD'
	SQLstmt$ = "SELECT * FROM racetrans WHERE RaceCode = " & CLng(mRaceCode.Text) & " ORDER BY ringno"
	' Create the new Dynaset.
	Set DStrans = DB.CreateDynaset(SQLstmt$, 4)
	DStrans.MoveLast
	FgCntTrans& = DStrans.RecordCount
	scrNoBirds.Text = FgCntTrans&
	If FgCntTrans& < 1 Then
	    CmdResults.Enabled = False
	Else
	    CmdResults.Enabled = True
	End If
	
	'Get ClockBird info
	DStrans.MoveFirst
	DStrans.FindFirst "PosLoft = 1"
	If DStrans.NoMatch Then
	    scrCB.Text = "No Results"
	Else
	    scrCB.Text = DStrans("yr") & "/" & DStrans("mark") & "/" & DStrans("ringno")
	    scrFT.Text = DStrans("FThr") & " hrs " & DStrans("FTmin") & " mins " & DStrans("FTsec") & " secs "
	    If IsNull(DStrans("velocity")) Then
		If IsNull(DStrans("VelocityCalc")) Then
		    scrVelocity.Text = "None Calculated"
		Else
		    scrVelocity.Text = DStrans("VelocityCalc")
		End If
	    Else
		scrVelocity.Text = DStrans("Velocity")
	    End If
	    
	    TblMaster.Edit
	    TblMaster("ringno") = DStrans("ringno")
	    TblMaster("yr") = DStrans("yr")
	    TblMaster("mark") = DStrans("mark")
	    TblMaster("ClubVel") = DStrans("velocity")
	    TblMaster("BiLoft") = FgCntTrans& 'set no of birds in race from loft
	    TblMaster("Distance") = DStrans("Distance")
	    TblMaster("ClkDate") = DStrans("clkDate")
	    TblMaster("ClkHrs") = DStrans("clkHr")
	    TblMaster("ClkMins") = DStrans("clkMin")
	    TblMaster("ClkSecs") = DStrans("clkSec")
	    TblMaster("FlyTime") = DStrans("FtHr") + (DStrans("FtMin") / 60) + (DStrans("FtSec") / 3600)
	    TblMaster("ClubPos") = DStrans("PosClub")
	    TblMaster("FedPos") = DStrans("PosFed")
	    TblMaster("NatPos") = DStrans("PosNat")
	    TblMaster("IntNatPos") = DStrans("PosINat")
	    
	    TblMaster.Update
	    
	End If

	TblToScr          'Move fields from table to screen
	mRelHrs.SetFocus
    End If

    ButtOk.Enabled = False
    ButtBirdList.Enabled = True
    
    If Val(scrNoBirds.Text) < 1 Then
	CmdPrtRaceSheet.Enabled = False
	CmdResults.Enabled = False
    Else
	CmdPrtRaceSheet.Enabled = True
	CmdResults.Enabled = True
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub

FMerror:
    Select Case Err
    Case 380
	Resume Next
    Case 3021 ' No Current record
	Resume Next
    Case Else
	MsgBox "Error " & Err & " " & Error$ & " found.", 16, "Program Error"
	Exit Sub
    End Select
End Sub

Sub Form_Load ()
Dim NL As String
NL$ = Chr$(10)

On Error GoTo FMErr:
    
FlagAct = 0

'Set screen
scrNoBirds.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
CSCombo1.Value = "R"

' Reset field change indicator
GLindChange% = 0

'Open Database
Set DB = OpenDatabase(DBFullFileName$, False, False)  'Open Database
Set TblMaster = DB.OpenTable("race")     'Open Table
TblMaster.Index = "raceCode"              'Open Index

'Open loft for next available Race Code
Set TblLoft = DB.OpenTable("loft")     'Open Table

Data1.DatabaseName = DBFullFileName$
Data1.RecordSource = "points"
Data1.Refresh

Exit Sub

FMErr:
    Select Case Err
    Case 3022
	MsgBox "Record already exists", 16, "Add record Error"
    Case 3026
	MsgBox "Disk is full no more additions possible.", 16, "Add Owner Error"
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add Owner Error"
    Case Else
	MsgBox "   Error - " & Err & NL$ & Error$(Err), 16, "Error in FMrace.frm"
    End Select
    
    Resume Next
 

End Sub

Sub Form_Unload (Cancel As Integer)
TblMaster.Close
TblLoft.Close
If FlagAct = -1 Then DStrans.Close
Set DStrans = Nothing
DB.Close
End Sub

Sub Mcoderace_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mindam_Click (Value As Integer)
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mindpm_Click (Value As Integer)
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mnotes_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mreldate_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mrelhrs_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mrelhrs_GotFocus ()
mRelHrs.SelStart = 0
mRelHrs.SelLength = Len(mRelHrs.Text)
End Sub

Sub Mrelmins_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Mrelmins_GotFocus ()
mRelMins.SelStart = 0
mRelMins.SelLength = Len(mRelMins.Text)
End Sub

Sub Mweather_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub nullfields ()
    mRelDate.Text = "" 'Date field?
    CSMcodepnt.Text = ""
    CSMcodepnt.Value = ""
    Mcoderace.Text = ""
    mRelHrs.Text = ""
    mRelMins.Text = ""
    Mindam.Value = -1
    Mindpm.Value = 0
    Mweather.Text = ""
    Mnotes.Text = ""
    mRaceCode.Text = ""
    
    scrNoBirds.Text = 0
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    CSCombo1.Text = ""
    CSCombo1.Value = ""
    'Let Status = ""
    scrWinVel.Text = ""
    scrCB.Text = ""
    scrFT.Text = ""
    scrVelocity.Text = ""
    GLindChange% = 0

End Sub

Sub scrNoBirds_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False

End Sub

Sub ScrToTbl ()
    ' doesnt allow for bred for stock or race
    If task.Text = "A" Then 'Only allow key change if new record
	TblMaster("reldate") = CVDate(mRelDate.Text)
	TblMaster("code_pnt") = Trim(CSMcodepnt.Value)
	TblMaster("code_race") = Trim(Mcoderace.Text)
	TblMaster("RaceCode") = CLng(mRaceCode.Text)
    End If
    TblMaster("WinVel") = Val(scrWinVel.Text)
    TblMaster("RaceStatus") = CSCombo1.Value
    TblMaster("relhrs") = Trim(mRelHrs.Text)
    TblMaster("relmins") = Trim(mRelMins.Text)
    TblMaster("race_weather") = Trim(Mweather.Text)
    TblMaster("race_notes") = Trim(Mnotes.Text)
    If Mindpm.Value = -1 Then
	TblMaster("indampm") = 0
    Else
	TblMaster("indampm") = -1
    End If
    If Trim(scrNoBirds.Text) = "" Then
	TblMaster("BiLoft") = Null
    Else
	TblMaster("BiLoft") = CLng(Trim(scrNoBirds.Text))
    End If
    If Trim(Text3.Text) = "" Then
	TblMaster("BiClub") = Null
    Else
	TblMaster("BiClub") = CLng(Trim(Text3.Text))
    End If
    If Trim(Text5.Text) = "" Then
	TblMaster("BiComb") = Null
    Else
	TblMaster("BiComb") = CLng(Trim(Text5.Text))
    End If
    If Trim(Text7.Text) = "" Then
	TblMaster("BiFed") = Null
    Else
	TblMaster("BiFed") = CLng(Trim(Text7.Text))
    End If
    If Trim(Text9.Text) = "" Then
	TblMaster("BiSec") = Null
    Else
	TblMaster("BiSec") = CLng(Trim(Text9.Text))
    End If
    If Trim(Text11.Text) = "" Then
	TblMaster("BiNat") = Null
    Else
	TblMaster("BiNat") = CLng(Trim(Text11.Text))
    End If
    If Trim(Text13.Text) = "" Then
	TblMaster("BiINat") = Null
    Else
	TblMaster("BiINat") = CLng(Trim(Text13.Text))
    End If
    'No Member Flyers
    If Trim(Text4.Text) = "" Then
	TblMaster("OOClub") = Null
    Else
	TblMaster("OOClub") = CLng(Trim(Text4.Text))
    End If
    If Trim(Text6.Text) = "" Then
	TblMaster("OOComb") = Null
    Else
	TblMaster("OOComb") = CLng(Trim(Text6.Text))
    End If
    If Trim(Text8.Text) = "" Then
	TblMaster("OOFed") = Null
    Else
	TblMaster("OOFed") = CLng(Trim(Text8.Text))
    End If
    If Trim(Text10.Text) = "" Then
	TblMaster("OOSec") = Null
    Else
	TblMaster("OOSec") = CLng(Trim(Text10.Text))
    End If
    If Trim(Text12.Text) = "" Then
	TblMaster("OONat") = Null
    Else
	TblMaster("OONat") = CLng(Trim(Text12.Text))
    End If
    If Trim(Text14.Text) = "" Then
	TblMaster("OOINat") = Null
    Else
	TblMaster("OOINat") = CLng(Trim(Text14.Text))
    End If
    TblMaster.Update
    TblMaster.Bookmark = TblMaster.LastModified
    ButtOk.Enabled = False
    GLindChange% = 0
End Sub

Sub scrWinVel_Change ()
    GLindChange% = -1
    ButtOk.Enabled = True
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False

End Sub

Sub scrWinVel_GotFocus ()
scrWinVel.SelStart = 0
scrWinVel.SelLength = Len(scrWinVel.Text)
End Sub

Sub TblToScr ()
    If IsNull(TblMaster("reldate")) Then
	mRelDate.Text = "  /  /    " '0
    Else
	mRelDate.Text = Format$(TblMaster("reldate"), "dd/mm/yyyy")
    End If
    'If IsNull(TblMaster("racedatet")) Then
	'Mreldate.Text = "  /  /  " '0
    'Else
	'Mreldate.Text = TblMaster("racedatet")
    'End If
    If IsNull(TblMaster("code_pnt")) Then
	CSMcodepnt.Value = ""
    Else
	CSMcodepnt.Value = TblMaster("code_pnt")
    End If

    If IsNull(TblMaster("RaceCode")) Then
	mRaceCode.Text = ""
    Else
	mRaceCode.Text = TblMaster("RaceCode")
    End If
    
    If IsNull(TblMaster("code_race")) Then
	Mcoderace.Text = ""
    Else
	Mcoderace.Text = TblMaster("code_race")
    End If
    If IsNull(TblMaster("RaceStatus")) Then
	CSCombo1.Value = ""
    Else
	CSCombo1.Value = TblMaster("RaceStatus")
    End If
    
    If IsNull(TblMaster("WinVel")) Then
	scrWinVel = ""
    Else
	scrWinVel = TblMaster("WinVel")
    End If

    If IsNull(TblMaster("relhrs")) Then
	mRelHrs.Text = ""
    Else
	mRelHrs.Text = TblMaster("relhrs")
    End If
    If IsNull(TblMaster("relmins")) Then
	mRelMins.Text = ""
    Else
	mRelMins.Text = TblMaster("relmins")
    End If
    If IsNull(TblMaster("indampm")) Then
	Mindam.Value = 0
	Mindpm.Value = -1
    Else
	Mindam.Value = -1
	Mindpm.Value = 0
    End If

'woek out how to get time from rel date/time mayybe

    'If IsNull(TblMaster("relhrs")) Then
    '    Mring.Text = ""
    'Else
    '    Mring.Text = TblMaster("ringno")
    'End If
    'set up Mrelhrs, Mrelmins, Mindam/pm
    
    If IsNull(TblMaster("race_weather")) Then
	Mweather.Text = ""
    Else
	Mweather.Text = TblMaster("race_weather")
    End If
    If IsNull(TblMaster("race_notes")) Then
	Mnotes.Text = ""
    Else
	Mnotes.Text = TblMaster("race_notes")
    End If
    'If IsNull(TblMaster("BiLoft")) Then
	'scrNoBirds.Text = ""
    'Else
	'scrNoBirds.Text = TblMaster("BiLoft")
    'End If
    If IsNull(TblMaster("BiClub")) Then
	Text3.Text = ""
    Else
	Text3.Text = TblMaster("BiClub")
    End If
    If IsNull(TblMaster("OOClub")) Then
	Text4.Text = ""
    Else
	Text4.Text = TblMaster("OOClub")
    End If
    If IsNull(TblMaster("BiComb")) Then
	Text5.Text = ""
    Else
	Text5.Text = TblMaster("BiComb")
    End If
    If IsNull(TblMaster("OOComb")) Then
	Text6.Text = ""
    Else
	Text6.Text = TblMaster("OOComb")
    End If
    If IsNull(TblMaster("BiFed")) Then
	Text7.Text = ""
    Else
	Text7.Text = TblMaster("BiFed")
    End If
    If IsNull(TblMaster("OOFed")) Then
	Text8.Text = ""
    Else
	Text8.Text = TblMaster("OOFed")
    End If
    If IsNull(TblMaster("BiSec")) Then
	Text9.Text = ""
    Else
	Text9.Text = TblMaster("BiSec")
    End If
    If IsNull(TblMaster("OOSec")) Then
	Text10.Text = ""
    Else
	Text10.Text = TblMaster("OOSec")
    End If
    If IsNull(TblMaster("BiNat")) Then
	Text11.Text = ""
    Else
	Text11.Text = TblMaster("BiNat")
    End If
    If IsNull(TblMaster("OONat")) Then
	Text12.Text = ""
    Else
	Text12.Text = TblMaster("OONat")
    End If
    If IsNull(TblMaster("BiINat")) Then
	Text13.Text = ""
    Else
	Text13.Text = TblMaster("BiINat")
    End If
    If IsNull(TblMaster("OOINat")) Then
	Text14.Text = ""
    Else
	Text14.Text = TblMaster("OOINat")
    End If
    GLindChange% = 0
    'Leave calculated figure here
    'should never be null
    'Actual Calculated no of birds will be in scrNoBirds.Text
    'if they are different to whats on file then let GLindChange indicate update required.
    'If TblMaster("BiLoft") = Val(scrNoBirds.Text) Then
    'Else
	'If IsNull(TblMaster("BiLoft")) Then
	    'scrNoBirds.Text = "0"
	    'GLindChange% = -1
	'Else
	    'scrNoBirds.Text = TblMaster("BiLoft")
	    'GLindChange% = -1
	'End If
    'End If
End Sub

Sub Text10_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text11_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text11_GotFocus ()
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)

End Sub

Sub Text12_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text13_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text13_GotFocus ()
Text13.SelStart = 0
Text13.SelLength = Len(Text13.Text)

End Sub

Sub Text14_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text3_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text3_GotFocus ()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Sub Text4_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text5_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text5_GotFocus ()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Sub Text6_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text7_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text7_GotFocus ()
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)

End Sub

Sub Text8_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text9_Change ()
    ButtOk.Enabled = True
    GLindChange% = -1
    ButtBirdList.Enabled = False
    CmdResults.Enabled = False
End Sub

Sub Text9_GotFocus ()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)

End Sub

