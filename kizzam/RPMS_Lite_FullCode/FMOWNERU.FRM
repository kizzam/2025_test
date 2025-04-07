VERSION 2.00
Begin Form FMowner 
   Caption         =   "Registered Owner Details"
   ClientHeight    =   6480
   ClientLeft      =   375
   ClientTop       =   405
   ClientWidth     =   8985
   Height          =   6885
   Left            =   315
   LinkTopic       =   "Form3"
   ScaleHeight     =   6480
   ScaleWidth      =   8985
   Top             =   60
   Width           =   9105
   Begin SSPanel Panel3D2 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   9015
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin TextBox Text15 
         DataField       =   "license"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2760
         TabIndex        =   37
         Text            =   "Text15"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   0  'No Search
         BorderEffect    =   0  'None
         ColDelim        =   ";"
         ColWidth        =   FMOWNERU.FRX:0000
         Contents        =   FMOWNERU.FRX:0040
         DataField       =   "bscmonth"
         DataSource      =   "Data1"
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   4200
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   ""
         ListDataSource  =   ""
         SearchString    =   ""
         StartSearch     =   0
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TextDisplay     =   0  'first visible
         Top             =   4800
         ValueCol        =   1
         Width           =   1695
      End
      Begin TextBox task 
         Height          =   285
         Left            =   8040
         MaxLength       =   1
         TabIndex        =   36
         Text            =   "a"
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin CommandButton ButtOK 
         Caption         =   "&OK"
         Height          =   735
         Left            =   7440
         TabIndex        =   35
         Top             =   600
         Width           =   1335
      End
      Begin TextBox Text14 
         DataField       =   "ownerph"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   33
         Text            =   "Text14"
         Top             =   3360
         Width           =   1815
      End
      Begin TextBox Text2 
         DataField       =   "name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   1920
         Width           =   2295
      End
      Begin CommandButton ButtCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   7440
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "F:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   7800
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "loft"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "ownername"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1560
         Width           =   3615
      End
      Begin TextBox Text3 
         DataField       =   "ownermark"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4200
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   5280
         Width           =   1215
      End
      Begin TextBox Text4 
         DataField       =   "oaddress1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   2280
         Width           =   3495
      End
      Begin TextBox Text5 
         DataField       =   "oaddress2"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin TextBox Text6 
         DataField       =   "ownercity"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "Text6"
         Top             =   3000
         Width           =   1215
      End
      Begin TextBox Text7 
         DataField       =   "ownerpc"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "Text7"
         Top             =   3000
         Width           =   1215
      End
      Begin TextBox Text8 
         DataField       =   "lat_deg"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Text            =   "Text8"
         Top             =   3840
         Width           =   495
      End
      Begin TextBox Text9 
         DataField       =   "lat_min"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4440
         TabIndex        =   8
         Text            =   "Text9"
         Top             =   3840
         Width           =   615
      End
      Begin TextBox Text10 
         DataField       =   "lat_sec"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Text            =   "Text10"
         Top             =   3840
         Width           =   735
      End
      Begin TextBox Text11 
         DataField       =   "lng_deg"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Text            =   "Text11"
         Top             =   4200
         Width           =   495
      End
      Begin TextBox Text12 
         DataField       =   "lng_min"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4440
         TabIndex        =   5
         Text            =   "Text12"
         Top             =   4200
         Width           =   615
      End
      Begin TextBox Text13 
         DataField       =   "lng_sec"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   6240
         TabIndex        =   4
         Text            =   "Text13"
         Top             =   4200
         Width           =   735
      End
      Begin Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Licence No. :"
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   1080
         Width           =   2055
      End
      Begin Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Breeding Season Commences (Month) :"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   4800
         Width           =   3375
      End
      Begin Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No. : "
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   3360
         Width           =   1095
      End
      Begin Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Ring Mark : "
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   5280
         Width           =   1695
      End
      Begin Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds : "
         Height          =   255
         Left            =   5280
         TabIndex        =   30
         Top             =   3840
         Width           =   975
      End
      Begin Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   1560
         Width           =   615
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loft : "
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   1920
         Width           =   495
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   2280
         Width           =   855
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address : "
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   2640
         Width           =   855
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "City/Town : "
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   3000
         Width           =   1095
      End
      Begin Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Post Code : "
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   3000
         Width           =   1095
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude - "
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   4200
         Width           =   1215
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees : "
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   3840
         Width           =   975
      End
      Begin Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes :"
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   3840
         Width           =   855
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude - "
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   3840
         Width           =   975
      End
      Begin Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees : "
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   4200
         Width           =   975
      End
      Begin Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes : "
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   4200
         Width           =   855
      End
      Begin Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds : "
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   4200
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
         TabIndex        =   2
         Top             =   240
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
        Text1.SetFocus
        Exit Sub
    Else
        Data1.Recordset.Edit
        MsgBox "No Changes effected", 16, "Notification Only"
    End If
End If

Unload FMowner

End Sub

Sub ButtOK_Click ()
    If Trim(Text1.Text = "") Then
        MsgBox "Name cannot be empty", 16, "Error: Cannot update record"
        Text1.SetFocus
        Exit Sub
    End If
    If Trim(Text2.Text = "") Then
        MsgBox "Loft cannot be empty", 16, "Error: Cannot update record"
        Text2.SetFocus
        Exit Sub
    End If
    If Trim(Text3.Text = "") Then
        MsgBox "Default Ring Mark cannot be empty", 16, "Error: Cannot update record"
        Text3.SetFocus
        Exit Sub
    End If
    If Trim(Text4.Text = "") Then
        MsgBox "Address1 cannot be empty", 16, "Error: Cannot update record"
        Text4.SetFocus
        Exit Sub
    End If
    If Trim(Text5.Text = "") Then
        MsgBox "Address2 cannot be empty", 16, "Error: Cannot update record"
        Text5.SetFocus
        Exit Sub
        End If
    If Trim(Text14.Text = "") Then
        MsgBox "Phone cannot be empty", 16, "Error: Cannot update record"
        Text14.SetFocus
        Exit Sub
    End If
    If CSCombo1.Value < 1 Or CSCombo1.Value > 12 Then
        MsgBox "Breeding season commence month (ie MAY) MUST be entered", 16, "Error: Cannot update record"
        CSCombo1.SetFocus
        Exit Sub
    End If
    If GLindchange% Then
            Data1.Recordset.Update
            MsgBox "Record Changed", 16, "Notification Only"
            GLindchange% = 0
            GLOwner$ = Data1.Recordset("ownername")
            GLAddress1$ = Data1.Recordset("owneradd1")
            GLAddress2$ = Data1.Recordset("owneradd2")
            GLCity$ = Data1.Recordset("ownercity")
            GLState$ = Data1.Recordset("ownerstate")
            GLPostcode$ = Data1.Recordset("ownerpc")
            GLState$ = Data1.Recordset("ownerstate")
            GLPhoneno$ = Data1.Recordset("ownerph")
            GLSerialNo$ = Data1.Recordset("license")
            GLMark$ = Data1.Recordset("ownermark")
            GLbscmonth% = Data1.Recordset("bscmonth")
    Else
            MsgBox "No Changes effected", 16, "Notification Only"
        End If
    GLindchange% = 0
    Unload FMowner
End Sub

Sub CSCombo1_Click ()
   GLindchange% = -1
End Sub

Sub Form_Activate ()
GLindchange% = 0
If task.Text = "S" Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text14.Enabled = True
    Text15.Enabled = True
    Text15.Visible = True
    Label19.Visible = True
End If
Data1.Recordset.Edit

End Sub

Sub Form_Load ()
   Data1.DatabaseName = DBFullFileName$
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   Text4.Enabled = False
   Text5.Enabled = False
   Text6.Enabled = False
   Text14.Enabled = False
   Text15.Enabled = False
   Text15.Visible = False
   Label19.Visible = False
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

