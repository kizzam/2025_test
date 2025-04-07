VERSION 2.00
Begin Form FMowner 
   Caption         =   "Registered Owner Details"
   ClientHeight    =   6795
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   9480
   Height          =   7305
   Left            =   15
   LinkTopic       =   "Form3"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Top             =   -120
   Width           =   9600
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   6480
      Width           =   9495
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      Begin SSFrame Frame3D2 
         Caption         =   "Measurement"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5640
         ShadowStyle     =   1  'Raised
         TabIndex        =   51
         Top             =   4080
         Width           =   2415
         Begin SSOption optYards 
            Caption         =   "Yards"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1320
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin SSOption optMetres 
            Caption         =   "Metres"
            Font3D          =   0  'None
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
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
         Height          =   615
         Left            =   7680
         TabIndex        =   0
         Top             =   1200
         Width           =   1335
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
         Height          =   615
         Left            =   7680
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Race Sheet - Select Style to suit your Club/Fed also Lines per page"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   240
         ShadowStyle     =   1  'Raised
         TabIndex        =   46
         Top             =   5520
         Width           =   9015
         Begin SSCommand cmdHelpRS 
            Caption         =   "&Help"
            Font3D          =   3  'Inset w/light shading
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8040
            TabIndex        =   49
            Top             =   360
            Width           =   855
         End
         Begin CSComboBox CSCombo2 
            AutoSearch      =   0  'No Search
            BorderEffect    =   0  'None
            ColDelim        =   ";"
            ColWidth        =   FMOWNER.FRX:0000
            Contents        =   FMOWNER.FRX:0040
            Dropped         =   0   'False
            FoundIndex      =   0
            Height          =   405
            Left            =   2280
            ListBoxHeight   =   0
            ListBoxWidth    =   0
            ListDataFields  =   ""
            ListDataSource  =   ""
            SearchString    =   ""
            StartSearch     =   0
            Style           =   2  'Dropdown List
            TabIndex        =   47
            TextDisplay     =   0  'first visible
            Top             =   360
            ValueCol        =   1
            Width           =   5655
         End
         Begin Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select Race Sheet Style:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   2175
         End
      End
      Begin SSPanel Panel3D5 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Club Name :"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   45
         Top             =   5160
         Width           =   1215
      End
      Begin TextBox Text18 
         Height          =   285
         Left            =   3840
         MaxLength       =   36
         TabIndex        =   20
         Text            =   "Text18"
         Top             =   5160
         Width           =   5415
      End
      Begin SSPanel Panel3D4 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "State :"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   44
         Top             =   2400
         Width           =   615
      End
      Begin TextBox Text17 
         Height          =   285
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "Text17"
         Top             =   2400
         Width           =   735
      End
      Begin SSPanel Panel3D3 
         Alignment       =   3  'Right Justify - TOP
         BevelOuter      =   0  'None
         Caption         =   "Federation Name :"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   43
         Top             =   4800
         Width           =   1695
      End
      Begin TextBox Text16 
         DataField       =   "ownerfed"
         Height          =   285
         Left            =   3840
         MaxLength       =   45
         TabIndex        =   19
         Text            =   "Text16"
         Top             =   4800
         Width           =   5415
      End
      Begin TextBox Text15 
         DataField       =   "license"
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "Text15"
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FMOWNER.FRX:02E6
         Contents        =   FMOWNER.FRX:0326
         DataField       =   "bscmonth"
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   405
         Left            =   3840
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TextDisplay     =   0  'first visible
         Top             =   4080
         ValueCol        =   1
         Width           =   1695
      End
      Begin TextBox task 
         Height          =   285
         Left            =   8400
         MaxLength       =   1
         TabIndex        =   41
         Text            =   "a"
         Top             =   5400
         Visible         =   0   'False
         Width           =   855
      End
      Begin TextBox Text14 
         DataField       =   "ownerph"
         Height          =   285
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   10
         Text            =   "Text14"
         Top             =   2760
         Width           =   1935
      End
      Begin TextBox Text2 
         DataField       =   "name"
         Height          =   285
         Left            =   2760
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1320
         Width           =   3495
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "C:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   8160
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "loft"
         Top             =   6120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "ownername"
         Height          =   285
         Left            =   2760
         MaxLength       =   24
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   3615
      End
      Begin TextBox Text3 
         DataField       =   "ownermark"
         Height          =   285
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   4440
         Width           =   1215
      End
      Begin TextBox Text4 
         DataField       =   "owneradd1"
         Height          =   285
         Left            =   2760
         MaxLength       =   24
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   1680
         Width           =   3495
      End
      Begin TextBox Text5 
         DataField       =   "owneradd2"
         Height          =   285
         Left            =   2760
         MaxLength       =   24
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   2040
         Width           =   3495
      End
      Begin TextBox Text6 
         DataField       =   "ownercity"
         Height          =   285
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   2400
         Width           =   1575
      End
      Begin TextBox Text7 
         DataField       =   "ownerpc"
         Height          =   285
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "Text7"
         Top             =   2400
         Width           =   1215
      End
      Begin TextBox Text8 
         DataField       =   "lat_deg"
         Height          =   285
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "Text8"
         Top             =   3240
         Width           =   495
      End
      Begin TextBox Text9 
         DataField       =   "lat_min"
         Height          =   285
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "Text9"
         Top             =   3240
         Width           =   615
      End
      Begin TextBox Text10 
         DataField       =   "lat_sec"
         Height          =   285
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   13
         Text            =   "Text10"
         Top             =   3240
         Width           =   735
      End
      Begin TextBox Text11 
         DataField       =   "lng_deg"
         Height          =   285
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "Text11"
         Top             =   3600
         Width           =   495
      End
      Begin TextBox Text12 
         DataField       =   "lng_min"
         Height          =   285
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "Text12"
         Top             =   3600
         Width           =   615
      End
      Begin TextBox Text13 
         DataField       =   "lng_sec"
         Height          =   285
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "Text13"
         Top             =   3600
         Width           =   735
      End
      Begin Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Software Licence No. :"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   600
         Width           =   2055
      End
      Begin Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Breeding Season Commences (Month) :"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   4080
         Width           =   3495
      End
      Begin Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No. :"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   2760
         Width           =   1095
      End
      Begin Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Default Ring Mark :"
         Height          =   255
         Left            =   1920
         TabIndex        =   38
         Top             =   4440
         Width           =   1815
      End
      Begin Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds :"
         Height          =   255
         Left            =   5160
         TabIndex        =   37
         Top             =   3240
         Width           =   975
      End
      Begin Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         Height          =   255
         Left            =   2040
         TabIndex        =   36
         Top             =   960
         Width           =   615
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loft : "
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   1320
         Width           =   495
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   1680
         Width           =   855
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   2040
         Width           =   855
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "City/Town :"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   2400
         Width           =   1095
      End
      Begin Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Post Code : "
         Height          =   255
         Left            =   6000
         TabIndex        =   31
         Top             =   2400
         Width           =   1095
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude - "
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   3600
         Width           =   1215
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees :"
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   3240
         Width           =   975
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes :"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   3240
         Width           =   855
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude - "
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   3240
         Width           =   975
      End
      Begin Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees :"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   3600
         Width           =   975
      End
      Begin Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes :"
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   3600
         Width           =   855
      End
      Begin Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds :"
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   3600
         Width           =   975
      End
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "System Details"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   120
         Width           =   2175
      End
   End
End
Option Explicit
Dim GLindchange As Integer

Sub ButtCancel_Click ()

Dim reply As Integer

If GLindchange% Then ' has any fields changed
    reply = MsgBox("Do you want to cancel changes?", 4, "Warning?")
    If reply <> 6 Then
	Text2.SetFocus
	Exit Sub
    Else
	MsgBox "No Changes effected", 16, "Notification Only"
    End If
End If

Unload FMowner

End Sub

Sub ButtOK_Click ()
    Dim U As String
    Dim T As String
    Dim C1 As Integer
    Dim Z As Integer
    Dim X As Integer
    Dim L As String

    If Trim(Text2.Text = "") Then
	MsgBox "Loft cannot be empty", 16, "Error: Cannot update record"
	Text2.SetFocus
	Exit Sub
    End If
    If Trim(Text3.Text) = "" Then
	MsgBox "Default Ring Mark cannot be empty", 16, "Error: Cannot update record"
	Text3.SetFocus
	Exit Sub
    End If
    If Trim(Text4.Text) = "" Then
	MsgBox "Address1 cannot be empty", 16, "Error: Cannot update record"
	Text4.SetFocus
	Exit Sub
    End If
    If Trim(Text14.Text) = "" Then
	MsgBox "Phone cannot be empty", 16, "Error: Cannot update record"
	Text14.SetFocus
	Exit Sub
    End If
    If Trim(Text6.Text) = "" Then
	MsgBox "City/Town cannot be empty", 16, "Error: Cannot update record"
	Text6.SetFocus
	Exit Sub
    End If
    If Trim(Text17.Text) = "" Then
	MsgBox "State cannot be empty", 16, "Error: Cannot update record"
	Text17.SetFocus
	Exit Sub
    End If
    If Trim(Text7.Text) = "" Then
	MsgBox "PostCode cannot be empty", 16, "Error: Cannot update record"
	Text7.SetFocus
	Exit Sub
    End If
    If Trim(Text8.Text) <> "" Then
	If IsNumeric(Trim(Text8.Text = "")) Then
	    If Val(Trim(Text8.Text)) > 179.99999 Then
		MsgBox "Latitude Degrees must be number between 0 and 179.99999", 16, "Error: Cannot update record"
		Text8.SetFocus
		Text8.SelStart = 0
		Text8.SelLength = Len(Text8.Text)
		Exit Sub
	    ElseIf Val(Trim(Text8.Text)) < 0 Then
		MsgBox "Latitude Degrees must be number between 0 and 179.99999", 16, "Error: Cannot update record"
		Text8.SetFocus
		Text8.SelStart = 0
		Text8.SelLength = Len(Text8.Text)
		Exit Sub
	    End If
	Else
	    MsgBox "Latitude Degrees must be numeric - number between 1 and 179.99999", 16, "Error: Cannot update record"
	    Text8.SetFocus
	    Text8.SelStart = 0
	    Text8.SelLength = Len(Text8.Text)
	    Exit Sub
	End If
    End If
    If Trim(Text9.Text) <> "" Then
	If IsNumeric(Trim(Text9.Text = "")) Then
	    If Val(Trim(Text9.Text)) > 59.99999 Then
		MsgBox "Latitude Minutes must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text9.SetFocus
		Text9.SelStart = 0
		Text9.SelLength = Len(Text9.Text)
		Exit Sub
	    ElseIf Val(Trim(Text9.Text)) < 0 Then
		MsgBox "Latitude Minutes must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text9.SetFocus
		Text9.SelStart = 0
		Text9.SelLength = Len(Text9.Text)
		Exit Sub
	    End If
	Else
	    MsgBox "Latitude Minutes must be numeric - number between 0 and 59.99999", 16, "Error: Cannot update record"
	    Text9.SetFocus
	    Text9.SelStart = 0
	    Text9.SelLength = Len(Text9.Text)
	    Exit Sub
	End If
    End If
    If Trim(Text10.Text) <> "" Then
	If IsNumeric(Trim(Text10.Text = "")) Then
	    If Val(Trim(Text10.Text)) > 59.99999 Then
		MsgBox "Latitude Seconds must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text10.SetFocus
		Text10.SelStart = 0
		Text10.SelLength = Len(Text10.Text)
		Exit Sub
	    ElseIf Val(Trim(Text10.Text)) < 0 Then
		MsgBox "Latitude Seconds must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text10.SetFocus
		Text10.SelStart = 0
		Text10.SelLength = Len(Text10.Text)
		Exit Sub
	    End If
	Else
	    MsgBox "Latitude Seconds must be numeric - number between 0 and 59.99999", 16, "Error: Cannot update record"
	    Text10.SetFocus
	    Text10.SelStart = 0
	    Text10.SelLength = Len(Text10.Text)
	    Exit Sub
	End If
    End If
    If Trim(Text11.Text) <> "" Then
	If IsNumeric(Trim(Text11.Text = "")) Then
	    If Val(Trim(Text11.Text)) > 179.99999 Then
		MsgBox "Longitude Degrees must be number between 0 and 179.99999", 16, "Error: Cannot update record"
		Text11.SetFocus
		Text11.SelStart = 0
		Text11.SelLength = Len(Text11.Text)
		Exit Sub
	    ElseIf Val(Trim(Text11.Text)) < 0 Then
		MsgBox "Longitude Degrees must be number between 0 and 179.99999", 16, "Error: Cannot update record"
		Text11.SetFocus
		Text11.SelStart = 0
		Text11.SelLength = Len(Text11.Text)
		Exit Sub
	    End If
	Else
	    MsgBox "Longitude Degrees must be numeric - number between 1 and 179.99999", 16, "Error: Cannot update record"
	    Text11.SetFocus
	    Text11.SelStart = 0
	    Text11.SelLength = Len(Text11.Text)
	    Exit Sub
	End If
    End If
    If Trim(Text12.Text) <> "" Then
	If IsNumeric(Trim(Text12.Text = "")) Then
	    If Val(Trim(Text12.Text)) > 59.99999 Then
		MsgBox "Longitude Minutes must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text12.SetFocus
		Text12.SelStart = 0
		Text12.SelLength = Len(Text12.Text)
		Exit Sub
	    ElseIf Val(Trim(Text12.Text)) < 0 Then
		MsgBox "Longitude Minutes must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text12.SetFocus
		Text12.SelStart = 0
		Text12.SelLength = Len(Text12.Text)
		Exit Sub
	    End If
	Else
	    MsgBox "Longitude Minutes must be numeric - number between 0 and 59.99999", 16, "Error: Cannot update record"
	    Text12.SetFocus
	    Text12.SelStart = 0
	    Text12.SelLength = Len(Text12.Text)
	    Exit Sub
	End If
    End If
    If Trim(Text13.Text) <> "" Then
	If IsNumeric(Trim(Text13.Text = "")) Then
	    If Val(Trim(Text13.Text)) > 59.99999 Then
		MsgBox "Longitude Seconds must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text13.SetFocus
		Text13.SelStart = 0
		Text13.SelLength = Len(Text13.Text)
		Exit Sub
	    ElseIf Val(Trim(Text13.Text)) < 0 Then
		MsgBox "Longitude Seconds must be number between 0 and 59.99999", 16, "Error: Cannot update record"
		Text13.SetFocus
		Text13.SelStart = 0
		Text13.SelLength = Len(Text13.Text)
		Exit Sub
	    End If
	Else
	    MsgBox "Longitude Seconds must be numeric - number between 0 and 59.99999", 16, "Error: Cannot update record"
	    Text13.SetFocus
	    Text13.SelStart = 0
	    Text13.SelLength = Len(Text13.Text)
	    Exit Sub
	End If
    End If
    
    If CSCombo1.Value < 1 Or CSCombo1.Value > 12 Then
	MsgBox "Breeding season commence month (ie MAY) MUST be entered", 16, "Error: Cannot update record"
	CSCombo1.SetFocus
	Exit Sub
    End If
    If GLindchange% Then
	    Data1.Recordset.Edit
	    Data1.Recordset("name") = "" & Trim(Text2.Text)
	    Data1.Recordset("ownermark") = "" & Trim(Text3.Text)
	    Data1.Recordset("owneradd1") = "" & Trim(Text4.Text)
	    Data1.Recordset("owneradd2") = "" & Trim(Text5.Text)
	    Data1.Recordset("ownercity") = "" & Trim(Text6.Text)
	    Data1.Recordset("ownerpc") = "" & Trim(Text7.Text)
	    Data1.Recordset("ownerstate") = "" & Trim(Text17.Text)
	    Data1.Recordset("ownerph") = "" & Trim(Text14.Text)
	    'data1.Recordset("license") = "" & Trim(text15.Text)
	    Data1.Recordset("lat_deg") = "" & Trim(Text8.Text)
	    Data1.Recordset("lat_min") = "" & Trim(Text9.Text)
	    Data1.Recordset("lat_sec") = "" & Trim(Text10.Text)
	    Data1.Recordset("lng_deg") = "" & Trim(Text11.Text)
	    Data1.Recordset("lng_min") = "" & Trim(Text12.Text)
	    Data1.Recordset("lng_sec") = "" & Trim(Text13.Text)
	    Data1.Recordset("bscmonth") = "" & CSCombo1.Value
	    Data1.Recordset("ownerfed") = "" & Trim(Text16.Text)
	    Data1.Recordset("OwnerClub") = "" & Trim(Text18.Text)
	    If OptYards Then
		Data1.Recordset("DistanceType") = "Y"
	    Else
		Data1.Recordset("DistanceType") = "M"
	    End If
	    'Setup Racesheet type & Lines per page
	    Select Case CSCombo2.Value
		Case 1
		Data1.Recordset("RSLPP") = 30
		Data1.Recordset("RSRaceFormType") = "1"
		Case 2
		Data1.Recordset("RSLPP") = 28
		Data1.Recordset("RSRaceFormType") = "2"
		Case 3
		Data1.Recordset("RSLPP") = 30
		Data1.Recordset("RSRaceFormType") = "3"
		Case 4
		Data1.Recordset("RSLPP") = 28
		Data1.Recordset("RSRaceFormType") = "4"
		Case 5
		Data1.Recordset("RSLPP") = 30
		Data1.Recordset("RSRaceFormType") = "5"
		Case 6
		Data1.Recordset("RSLPP") = 28
		Data1.Recordset("RSRaceFormType") = "6"
		Case 7
		Data1.Recordset("RSLPP") = 30
		Data1.Recordset("RSRaceFormType") = "7"
		Case 8
		Data1.Recordset("RSLPP") = 28
		Data1.Recordset("RSRaceFormType") = "8"
		Case 9
		Data1.Recordset("RSLPP") = 30
		Data1.Recordset("RSRaceFormType") = "9"
		Case 10
		Data1.Recordset("RSLPP") = 28
		Data1.Recordset("RSRaceFormType") = "10"
		Case 11
		Data1.Recordset("RSLPP") = 25
		Data1.Recordset("RSRaceFormType") = "11"
		Case 12
		Data1.Recordset("RSLPP") = 25
		Data1.Recordset("RSRaceFormType") = "12"
		Case 13
		Data1.Recordset("RSLPP") = 30
		Data1.Recordset("RSRaceFormType") = "13"
		Case 14
		Data1.Recordset("RSLPP") = 28
		Data1.Recordset("RSRaceFormType") = "14"
		Case 15
		Data1.Recordset("RSLPP") = 25
		Data1.Recordset("RSRaceFormType") = "15"
	    End Select
	    
	    Data1.Recordset.Update
	    MsgBox "Record Changed", 16, "Notification Only"
	    GLindchange% = 0
	    GetLoftDets
    Else
	    MsgBox "No Changes effected", 16, "Notification Only"
    End If
    GLindchange% = 0
    Unload FMowner
End Sub

Sub cmdHelpRS_Click ()
frmHelpRS.Show 1
End Sub

Sub CSCombo1_Click ()
   GLindchange% = -1
End Sub

Sub CSCombo2_Change ()
   GLindchange% = -1
End Sub

Sub CSCombo2_Click ()
GLindchange% = -1
End Sub

Sub Form_Activate ()
    Text1.Text = Data1.Recordset("ownername")
    If IsNull(Data1.Recordset("name")) Then
	Text2.Text = ""
    Else
	Text2.Text = Data1.Recordset("name")
    End If
    If IsNull(Data1.Recordset("ownermark")) Then
	Text3.Text = ""
    Else
	Text3.Text = Data1.Recordset("ownermark")
    End If
    If IsNull(Data1.Recordset("owneradd1")) Then
	Text4.Text = ""
    Else
	Text4.Text = Data1.Recordset("owneradd1")
    End If
    If IsNull(Data1.Recordset("owneradd2")) Then
	Text5.Text = ""
    Else
	Text5.Text = Data1.Recordset("owneradd2")
    End If
    If IsNull(Data1.Recordset("ownercity")) Then
	Text6.Text = ""
    Else
	Text6.Text = Data1.Recordset("ownercity")
    End If
    If IsNull(Data1.Recordset("ownerpc")) Then
	Text7.Text = ""
    Else
	Text7.Text = Data1.Recordset("ownerpc")
    End If
    If IsNull(Data1.Recordset("ownerstate")) Then
	Text17.Text = ""
    Else
	Text17.Text = Data1.Recordset("ownerstate")
    End If
    If IsNull(Data1.Recordset("ownerph")) Then
	Text14.Text = ""
    Else
	Text14.Text = Data1.Recordset("ownerph")
    End If

    text15.Text = LicenseDisplay(Data1.Recordset("license"))

    If IsNull(Data1.Recordset("lat_deg")) Then
	Text8.Text = ""
    Else
	Text8.Text = Data1.Recordset("lat_deg")
    End If
    If IsNull(Data1.Recordset("lat_min")) Then
	Text9.Text = ""
    Else
	Text9.Text = Data1.Recordset("lat_min")
    End If
    If IsNull(Data1.Recordset("lat_sec")) Then
	Text10.Text = ""
    Else
	Text10.Text = Data1.Recordset("lat_sec")
    End If
    If IsNull(Data1.Recordset("lng_deg")) Then
	Text11.Text = ""
    Else
	Text11.Text = Data1.Recordset("lng_deg")
    End If
    If IsNull(Data1.Recordset("lng_min")) Then
	Text12.Text = ""
    Else
	Text12.Text = Data1.Recordset("lng_min")
    End If
    If IsNull(Data1.Recordset("lng_sec")) Then
	Text13.Text = ""
    Else
	Text13.Text = Data1.Recordset("lng_sec")
    End If

    CSCombo1.Value = Data1.Recordset("bscmonth")
    
    If IsNull(Data1.Recordset("ownerfed")) Then
	Text16.Text = ""
    Else
	Text16.Text = Data1.Recordset("ownerfed")
    End If

    If IsNull(Data1.Recordset("OwnerClub")) Then
	Text18.Text = ""
    Else
	Text18.Text = Data1.Recordset("OwnerClub")
    End If
    
    If IsNull(Data1.Recordset("DistanceType")) Then
	OptMetres.Value = -1
    ElseIf Data1.Recordset("DistanceType") = "Y" Then
	OptYards.Value = -1
    Else
	OptMetres.Value = -1
    End If

    If IsNull(Data1.Recordset("RSLPP")) Or IsNull(Data1.Recordset("RSRaceFormType")) Then
	CSCombo2.Value = 1
    Else
	CSCombo2.Value = Data1.Recordset("RSRaceFormType")
    End If

    GLindchange% = 0

    If task.Text = "S" Then
	Text1.Enabled = True
	Text2.Enabled = True
	Text3.Enabled = True
	Text4.Enabled = True
	Text5.Enabled = True
	Text6.Enabled = True
	Text14.Enabled = True
    ' Show to System manager, Licence & No
	Label19.Visible = True
	text15.Enabled = True
	text15.Visible = True
	Text16.Enabled = True
	Text18.Enabled = True
    End If

    If task.Text = "N" Then
	Text1.Enabled = False
	Text2.Enabled = True
	Text3.Enabled = True
	Text4.Enabled = True
	Text5.Enabled = True
	Text6.Enabled = True
	Text7.Enabled = True
	Text14.Enabled = True
	Text16.Enabled = True
	Text18.Enabled = True
	text15.Enabled = False
	text15.Visible = True
	Label19.Visible = True
    End If

End Sub

Sub Form_Load ()
   Data1.DatabaseName = DBFullFileName$
   Data1.RecordSource = "loft"
   Data1.Refresh
   Data1.Recordset.MoveFirst
End Sub

Sub optMetres_Click (Value As Integer)
   GLindchange% = -1
End Sub

Sub optYards_Click (Value As Integer)
   GLindchange% = -1
End Sub

Sub Text1_Change ()
GLindchange% = -1
End Sub

Sub Text10_Change ()
   GLindchange% = -1
End Sub

Sub Text11_Change ()
      GLindchange% = -1
End Sub

Sub Text12_Change ()
   GLindchange% = -1
End Sub

Sub Text13_Change ()
   GLindchange% = -1
End Sub

Sub Text14_Change ()
   GLindchange% = -1
End Sub

Sub Text15_Change ()
GLindchange% = -1
End Sub

Sub Text16_Change ()
GLindchange% = -1
End Sub

Sub Text17_Change ()
   GLindchange% = -1
End Sub

Sub Text18_Change ()
GLindchange% = -1
End Sub

Sub Text2_Change ()
   GLindchange% = -1
End Sub

Sub Text3_Change ()
   GLindchange% = -1
End Sub

Sub Text4_Change ()
   GLindchange% = -1
End Sub

Sub Text5_Change ()
   GLindchange% = -1
End Sub

Sub Text6_Change ()
   GLindchange% = -1
End Sub

Sub Text7_Change ()
   GLindchange% = -1
End Sub

Sub Text8_Change ()
   GLindchange% = -1
End Sub

Sub Text9_Change ()
   GLindchange% = -1
End Sub

