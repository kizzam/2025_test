VERSION 2.00
Begin Form FMaddteam 
   Caption         =   "Add Team Code"
   ClientHeight    =   2280
   ClientLeft      =   870
   ClientTop       =   1665
   ClientWidth     =   7560
   Height          =   2685
   Left            =   810
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   7560
   Top             =   1320
   Width           =   7680
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   7575
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7575
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
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
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
         Height          =   495
         Left            =   6120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin TextBox code 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   0
         Top             =   600
         Width           =   2055
      End
      Begin TextBox description 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1080
         Width           =   3975
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code : "
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    Dim reply As Integer
    If code.Text = "" And description.Text = "" Then
	MsgBox "Add record cancelled", 16, "Notification only"
	FMaddteam.Hide
	Exit Sub
    End If
    reply% = MsgBox("Please confirm cancellation of add record", 36, "Record Add Cancellation")
    If reply% = 6 Then 'ie YES
	code.Text = ""
	description.Text = ""
	FMaddteam.Hide
    Else
	Exit Sub
    End If

End Sub

Sub ButtCancel_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Click Cancel to tell system to cancel/ignor data entered/typed"
End Sub

Sub ButtOK_Click ()
    
    On Error GoTo msgspot

    'Check all data entered
    If Trim(code.Text) = "" Then
	MsgBox "Code MUST be entered", 16, "Add record Error"
	code.SetFocus
	code.SelStart = 0
	code.SelLength = Len(code.Text)
	Exit Sub
    End If
    If Trim(description.Text) = "" Then
	MsgBox "Code description MUST be entered", 16, "Add record Error"
	description.SetFocus
	description.SelStart = 0
	description.SelLength = Len(description.Text)
	Exit Sub
    End If

    'Add new record
    FMteam.Data1.Recordset.AddNew
    FMteam.Data1.Recordset("code") = code.Text
    FMteam.Data1.Recordset("description") = description.Text
    FMteam.Data1.Recordset("status") = 0
    FMteam.Data1.Recordset.Update
    FMteam.Data1.Refresh
    code.Text = ""
    description.Text = ""
    FMaddteam.Hide

    Exit Sub

msgspot:
    Select Case Err
    Case 3022
	MsgBox "That code already exists.", 16, "Add Code Error"
	code.SetFocus
	code.SelStart = 0
	code.SelLength = Len(code.Text)
    Case 3026
	MsgBox "Disk is full no more additions possible.", 16, "Add code Error"
    Case 3043
	MsgBox "Network error. Operation can't be performed.", 16, "Add code Error"
    Case Else
	MsgBox "Error number " & Err & " occurred at line " & Erl
	Resume Next
    End Select
    
    Exit Sub
End Sub

Sub ButtOK_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = "Click OK to tell system to accept data entered/typed"
End Sub

Sub code_LostFocus ()
code.Text = UCase(code.Text)
End Sub

Sub code_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = "Enter/Type Unique code for this record"
End Sub

Sub description_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Enter/Type description of this record"
End Sub

Sub Form_Activate ()
    code.SetFocus
End Sub

Sub Form_Unload (Cancel As Integer)
    FMteam.Table1.Refresh
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = ""
End Sub

