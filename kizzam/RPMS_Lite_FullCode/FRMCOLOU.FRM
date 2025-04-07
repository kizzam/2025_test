VERSION 2.00
Begin Form frmColours 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Colours Sample"
   ClientHeight    =   10065
   ClientLeft      =   2010
   ClientTop       =   2445
   ClientWidth     =   9735
   FillStyle       =   4  'Upward Diagonal
   Height          =   10575
   Left            =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   9735
   Top             =   1995
   Width           =   9855
   Begin CommandButton ButYellow 
      Caption         =   "Yellow"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin CommandButton butBlue 
      Caption         =   "Blue"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2400
      Width           =   2775
   End
   Begin CSComboBox CSCombo1 
      AutoSearch      =   0  'No Search
      BorderEffect    =   0  'None
      ColDelim        =   "|"
      ColWidth        =   FRMCOLOU.FRX:0000
      Contents        =   FRMCOLOU.FRX:0040
      Dropped         =   0   'False
      FoundIndex      =   0
      Height          =   405
      Index           =   3
      Left            =   480
      ListBoxHeight   =   0
      ListBoxWidth    =   0
      ListDataFields  =   ""
      ListDataSource  =   ""
      SearchString    =   ""
      StartSearch     =   0
      TabIndex        =   3
      TextDisplay     =   1  'all visible
      Top             =   840
      ValueCol        =   -1
      Width           =   2895
   End
   Begin TextBox Text1 
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   3960
      Width           =   4935
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00C0FFC0&
      BevelOuter      =   0  'None
      Caption         =   "Panel3D1"
      FloodColor      =   &H00E0E0E0&
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin SSCommand Butt3Exit 
      Caption         =   "E&xit"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   8880
      Width           =   1575
   End
   Begin Label Label1 
      Caption         =   "Label1"
      Height          =   2175
      Left            =   2040
      TabIndex        =   5
      Top             =   5640
      Width           =   7095
   End
End
'Declare GetSysColor lib "user" (byval Color as Integer) as long

'In your form load event:
'me.backcolor=GetSysColor(4)

Dim MyColor(99)
'MyColor(1) = 37 'Blue
'Const MyColor = 37 'Blue

Sub butBlue_Click ()
Me.BackColor = vbBlue
End Sub

Sub Butt3Exit_Click ()
Unload Me
End Sub

Sub ButYellow_Click ()
Me.BackColor = vbYELLOW
End Sub

Sub Form_Activate ()

On Error Resume Next

'frmcolours.FillColor = vbBLUE
Me.BackColor = vbRED
Me.BackColor = "vb" & "BLUE"
text1.Text = frmcolours.BackColor
label1.Caption = cscombo1(3).Value


End Sub

