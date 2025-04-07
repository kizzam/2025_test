VERSION 2.00
Begin Form FMaddeyesgn 
   Caption         =   "Add Eye Sign"
   ClientHeight    =   3015
   ClientLeft      =   1845
   ClientTop       =   2340
   ClientWidth     =   6225
   Height          =   3420
   Left            =   1785
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   6225
   Top             =   1995
   Width           =   6345
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   6255
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6255
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
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   1095
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
         Height          =   495
         Left            =   4800
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin TextBox desc 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   24
         TabIndex        =   1
         Top             =   1320
         Width           =   3015
      End
      Begin TextBox code 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   0
         Top             =   600
         Width           =   1095
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
         Top             =   1320
         Width           =   1335
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
         Top             =   720
         Width           =   735
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    Dim reply As Integer
    If code.Text = "" And desc.Text = "" Then
        MsgBox "Add record cancelled", 16, "Notification only"
        FMaddeyesgn.Hide
        Exit Sub
    End If
    reply% = MsgBox("Please confirm cancellation of add record", 36, "Record Add Cancellation")
    If reply% = 6 Then 'ie YES
        code.Text = ""
        desc.Text = ""
        FMaddeyesgn.Hide
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
        MsgBox "Unique Eyesign Code MUST be entered", 16, "Add record Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
        Exit Sub
    End If
    If Trim(desc.Text) = "" Then
        MsgBox "Eyesign description MUST be entered", 16, "Add record Error"
        desc.SetFocus
        desc.SelStart = 0
        desc.SelLength = Len(desc.Text)
        Exit Sub
    End If

    'Add new record
    FMeyesgn.Data1.Recordset.AddNew
    FMeyesgn.Data1.Recordset("code") = code.Text
    FMeyesgn.Data1.Recordset("desc") = desc.Text
    FMeyesgn.Data1.Recordset("system") = False
    FMeyesgn.Data1.Recordset.Update
    FMeyesgn.Data1.Refresh
    code.Text = ""
    desc.Text = ""
    FMaddeyesgn.Hide

    Exit Sub

msgspot:
    Select Case Err
    Case 3022
        MsgBox "Code already exists - Code MUST be unique", 16, "Add record Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add cage Error"
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add cage Error"
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
        MsgPanel.Caption = "Enter/Type Unique code for this eyesign"
End Sub

Sub desc_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
            MsgPanel.Caption = "Enter/Type description of this eyesign"
End Sub

Sub Form_Activate ()
    code.SetFocus
End Sub

Sub Form_Unload (Cancel As Integer)
    FMeyesgn.Table1.Refresh
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = ""
End Sub

