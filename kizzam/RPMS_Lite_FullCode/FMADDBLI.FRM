VERSION 2.00
Begin Form FMaddBline 
   Caption         =   "Add BloodLine"
   ClientHeight    =   3825
   ClientLeft      =   1620
   ClientTop       =   1665
   ClientWidth     =   6000
   Height          =   4230
   Left            =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6000
   Top             =   1320
   Width           =   6120
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   6015
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6015
      Begin SSCommand CmdCancel 
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
         Left            =   3000
         TabIndex        =   4
         Top             =   2760
         Width           =   1215
      End
      Begin SSCommand CmdOK 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   2760
         Width           =   1215
      End
      Begin TextBox desc 
         Height          =   375
         Left            =   2040
         MaxLength       =   32
         TabIndex        =   1
         Top             =   1680
         Width           =   3735
      End
      Begin TextBox code 
         Height          =   375
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   0
         Top             =   960
         Width           =   855
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
   End
End
Option Explicit

Sub CmdCancel_Click ()
    Dim reply As Integer
    If code.Text = "" And desc.Text = "" Then
        MsgBox "Add record cancelled", 16, "Notification only"
        FMaddBLine.Hide
        Exit Sub
    End If
    reply% = MsgBox("Please confirm cancellation of add record", 36, "Record Add Cancellation")
    If reply% = 6 Then 'ie YES
        code.Text = ""
        desc.Text = ""
        FMaddBLine.Hide
    Else
        Exit Sub
    End If

End Sub

Sub cmdOK_Click ()
    
    On Error GoTo msgspot

    'Check all data entered
    If Trim(code.Text) = "" Then
        MsgBox "Code MUST be entered", 16, "Add record Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
        Exit Sub
    End If
    If Trim(desc.Text) = "" Then
        MsgBox "Description MUST be entered", 16, "Add record Error"
        desc.SetFocus
        desc.SelStart = 0
        desc.SelLength = Len(desc.Text)
        Exit Sub
    End If

    'Add new record
    FMBLine.Data1.Recordset.AddNew
    FMBLine.Data1.Recordset("code") = code.Text
    FMBLine.Data1.Recordset("desc") = desc.Text
    FMBLine.Data1.Recordset("system") = False
    FMBLine.Data1.Recordset("status") = "A"
    FMBLine.Data1.Recordset.Update
    FMBLine.Data1.Refresh
    code.Text = ""
    desc.Text = ""
    FMaddBLine.Hide

    Exit Sub

msgspot:
    Select Case Err
    Case 3022
        MsgBox "That code already exists.", 16, "Add Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
        Exit Sub
    Case 3026
        MsgBox "Disk is full no more additions possible.", 16, "Add cage Error"
        Exit Sub
    Case 3043
        MsgBox "Network error. Operation can't be performed.", 16, "Add cage Error"
    Case Else
        MsgBox "Error" & Err, 16, "Error Code " & Val(Error) & " "
    End Select
    Exit Sub

End Sub

Sub code_LostFocus ()
code.Text = UCase(code.Text)
End Sub

Sub desc_LostFocus ()
desc.Text = UCase(desc.Text)
End Sub

Sub Form_Activate ()
code.SetFocus
End Sub

Sub Form_Unload (Cancel As Integer)
    FMBLine.Table1.Refresh
End Sub

