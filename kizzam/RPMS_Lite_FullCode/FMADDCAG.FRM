VERSION 2.00
Begin Form FMaddcage 
   Caption         =   "Add Cage"
   ClientHeight    =   2760
   ClientLeft      =   1665
   ClientTop       =   1890
   ClientWidth     =   6585
   Height          =   3165
   Left            =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6585
   Top             =   1545
   Width           =   6705
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   6615
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6615
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
         Left            =   5040
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin SSCommand ButtOk 
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
         Left            =   5040
         TabIndex        =   3
         Top             =   360
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
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin TextBox desc 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   360
         Left            =   2160
         MaxLength       =   24
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
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
         Left            =   480
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
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
End
Option Explicit

Sub ButtCancel_Click ()
    Dim reply As Integer
    If code.Text = "" And desc.Text = "" Then
        MsgBox "Add record cancelled", 16, "Notification only"
        FMaddcage.Hide
        Exit Sub
    End If
    reply% = MsgBox("Please confirm cancellation of add record", 36, "Record Add Cancellation")
    If reply% = 6 Then 'ie YES
        code.Text = ""
        desc.Text = ""
        FMaddcage.Hide
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
        MsgBox "Cage Code MUST be entered", 16, "Add record Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
        Exit Sub
    End If
    If Trim(desc.Text) = "" Then
        MsgBox "cage description MUST be entered", 16, "Add record Error"
        desc.SetFocus
        desc.SelStart = 0
        desc.SelLength = Len(desc.Text)
        Exit Sub
    End If

    'Add new record
    FMcage.Data1.Recordset.AddNew
    FMcage.Data1.Recordset("code") = code.Text
    FMcage.Data1.Recordset("desc") = desc.Text
    FMcage.Data1.Recordset("system") = False
    FMcage.Data1.Recordset.Update
    FMcage.Data1.Refresh
    code.Text = ""
    desc.Text = ""
    FMaddcage.Hide

    Exit Sub

msgspot:
    Select Case Err
    Case 3022
        MsgBox "That colour code already exists.", 16, "Add colour Error"
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
    MsgPanel.Caption = "Enter/Type Unique code for this record"
End Sub

Sub desc_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgPanel.Caption = "Enter/Type description of this record"
End Sub

Sub Form_Activate ()
    code.SetFocus
End Sub

Sub Form_Unload (Cancel As Integer)
    FMcage.Table1.Refresh
End Sub

Sub Panel3D1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgPanel.Caption = ""
End Sub

