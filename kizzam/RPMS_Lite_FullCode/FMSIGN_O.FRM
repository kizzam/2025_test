VERSION 2.00
Begin Form FMsign_on 
   Caption         =   "Sign In"
   ClientHeight    =   2700
   ClientLeft      =   2370
   ClientTop       =   2295
   ClientWidth     =   4920
   Height          =   3105
   Left            =   2310
   LinkTopic       =   "Form4"
   ScaleHeight     =   2700
   ScaleWidth      =   4920
   Top             =   1950
   Width           =   5040
   Begin SSPanel Panel3D2 
      BackColor       =   &H0080FF80&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   4935
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H0080FF80&
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "H:\DEV_VB\RP\RP.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   4
         ReadOnly        =   -1  'True
         RecordSource    =   "pass"
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin CommandButton ButtCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin CommandButton ButtOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin CSComboBox CSCombo1 
         AutoSearch      =   2  'Incremental Search
         BorderEffect    =   0  'None
         ColDelim        =   "	"
         ColWidth        =   FMSIGN_O.FRX:0000
         Contents        =   FMSIGN_O.FRX:0040
         DataField       =   "username"
         DataSource      =   "Data1"
         Dropped         =   0   'False
         FoundIndex      =   0
         Height          =   315
         Left            =   1320
         ListBoxHeight   =   0
         ListBoxWidth    =   0
         ListDataFields  =   "username"
         ListDataSource  =   "Data1"
         SearchString    =   ""
         Sorted          =   -1  'True
         StartSearch     =   0
         TabIndex        =   1
         TextDisplay     =   1  'all visible
         Top             =   480
         ValueCol        =   -1
         Width           =   2175
      End
      Begin Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Max 6 Characters"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password : "
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
End
Option Explicit
    Const MB_ICONSTOP = 16          ' Define Icons.
    Const MB_ICONEXCLAMATION = 48   ' Define Icons.
    Const MB_OK = 0

    Dim Title As String             ' Declare variables.
    Dim Msg As String               ' Declare variables.
    Dim DgDef As Integer
    Dim tries As Integer

Sub ButtCancel_Click ()
    Unload FMSign_On
    End
End Sub

Sub ButtOK_Click ()
    Dim Key As String
    
    Title$ = "Security Violation"   ' Assign Values
    Msg$ = "Invalid Password"
    DgDef% = MB_ICONSTOP + MB_OK
    
    ' Do nothing if no password entered.
    If IsEmpty(Text1.Text) Then
	Exit Sub
    End If

    'Validate Password
    'allow max of 3 tries before unloading program
    ' Too many invalid signon attempts.
    tries% = tries% + 1
    If tries% >= 4 Then
	Msg$ = "Too many retries, Program aborting..."
	DgDef% = MB_OK + MB_ICONEXCLAMATION
	MsgBox Msg$, DgDef%, Title$
	Unload FMSign_On
	End
	Exit Sub
    End If

    'Validate user password
    'Data1.Refresh
    
    Screen.MousePointer = 11

    Data1.Recordset.FindFirst "username = '" & CSCombo1 & "'"
	
    If Trim(Data1.Recordset("password")) = Trim(Text1) Then
	'password validated
	Unload FMSign_On
	Screen.MousePointer = 0
	Exit Sub
    Else
	'password NOT validated
	Beep
	Text1.SelStart = 0
	Text1.SelLength = Len(Trim(Text1.Text))
	MsgBox Msg$, DgDef%, Title$
	Text1.SetFocus
	'select password
    End If

    ' Check for the user in the file.
    'Key$ = Encrypt(Text1.Text)
    'TBLic.Seek "=", Key$
    'If TBLic.NoMatch Then
       ' Beep
       ' PWord.SelStart = 0
	'PWord.SelLength = Len(PWord.Text)
    'Else

	'Unload SignOn
    'End If
    Screen.MousePointer = 0
End Sub

Sub Form_Load ()
    Data1.DatabaseName = DBFullFileName
    Data1.RecordSource = "Pass"
    Data1.Refresh
    tries% = 0
    'CSCombo1.SetFocus
End Sub

