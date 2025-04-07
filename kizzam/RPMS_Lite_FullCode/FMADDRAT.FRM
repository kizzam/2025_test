VERSION 2.00
Begin Form FMAddRating 
   Caption         =   "Add Rating"
   ClientHeight    =   3360
   ClientLeft      =   2130
   ClientTop       =   2055
   ClientWidth     =   5760
   Height          =   3765
   Left            =   2070
   LinkTopic       =   "Form3"
   ScaleHeight     =   3360
   ScaleWidth      =   5760
   Top             =   1710
   Width           =   5880
   Begin SSPanel MsgPanel2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin SSPanel MsgPanel1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin SSPanel MsgPanel 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   3495
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin TextBox Level 
         Height          =   375
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1800
         Width           =   615
      End
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
         Top             =   2280
         Width           =   1095
      End
      Begin SSCommand CmdAdd 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin TextBox desc 
         Height          =   375
         Left            =   2520
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin TextBox code 
         Height          =   375
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   0
         Top             =   600
         Width           =   615
      End
      Begin Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "(between 1 - 99)"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
      Begin Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(Max 2 char code)"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   720
         Width           =   1935
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
         Left            =   960
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
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
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
   End
End
Option Explicit

Sub CmdAdd_Click ()
    
    On Error GoTo msgspot

    'Check all data entered
    If Trim(code.Text) = "" Then
        MsgBox "Code MUST be entered", 16, "Add Error"
        code.SetFocus
        code.SelStart = 0
        code.SelLength = Len(code.Text)
        Exit Sub
    End If
    If Trim(desc.Text) = "" Then
        MsgBox "Description MUST be entered", 16, "Add Error"
        desc.SetFocus
        desc.SelStart = 0
        desc.SelLength = Len(desc.Text)
        Exit Sub
    End If
    If Trim(Level.Text) = "" Then
        MsgBox "Level MUST be entered; Must be a number between 1 - 99, 1 being the best to 99 worst", 16, "Add Error"
        Level.SetFocus
        Level.SelStart = 0
        Level.SelLength = Len(Level.Text)
        Exit Sub
    End If
    If Not IsNumeric(Level.Text) Then
        MsgBox "Level MUST be numeric number between 1 & 99; 1 being the best to 99 the worst", 16, "Add Error"
        Level.SetFocus
        Level.SelStart = 0
        Level.SelLength = Len(Level.Text)
        Exit Sub
    End If

    'Add new record
    FMLstRatings.Data1.Recordset.AddNew
    FMLstRatings.Data1.Recordset("code") = code.Text
    FMLstRatings.Data1.Recordset("desc") = desc.Text
    FMLstRatings.Data1.Recordset("system") = False
    FMLstRatings.Data1.Recordset("status") = "A"
    FMLstRatings.Data1.Recordset("Level") = CInt(Level.Text)
    FMLstRatings.Data1.Recordset.Update
    FMLstRatings.Data1.Refresh
    code.Text = ""
    desc.Text = ""
    Level.Text = ""
    FMaddrating.Hide

    Exit Sub

msgspot:
    Select Case Err
    Case 3022
        MsgBox "That code already exists.", 16, "Add Error"
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

Sub CmdCancel_Click ()
    Dim reply As Integer
    If code.Text = "" And desc.Text = "" Then
        MsgBox "Add record cancelled", 16, "Notification only"
        FMaddrating.Hide
        Exit Sub
    End If
    reply% = MsgBox("Please confirm cancellation of add record", 36, "Record Add Cancellation")
    If reply% = 6 Then 'ie YES
        code.Text = ""
        desc.Text = ""
        FMaddrating.Hide
    Else
        Exit Sub
    End If


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
    FMLstRatings.Table1.Refresh
End Sub

