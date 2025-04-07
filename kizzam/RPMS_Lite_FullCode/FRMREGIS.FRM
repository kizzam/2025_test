VERSION 2.00
Begin Form frmRegister 
   Caption         =   "Register software"
   ClientHeight    =   5910
   ClientLeft      =   540
   ClientTop       =   1620
   ClientWidth     =   8925
   ControlBox      =   0   'False
   Height          =   6420
   Left            =   480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8925
   Top             =   1170
   Width           =   9045
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00E0E0E0&
      BevelInner      =   1  'Inset
      Caption         =   "The special Demo mode allows limited use of the system for review prior purchase.  Real data, breeding records etc can be entered also reports printed. When you decide to purchase call 0409495571 for code to unlock the full system."
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   480
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   8175
   End
   Begin SSCommand cmdDemoMode 
      Caption         =   "Demo mode"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin TextBox LicNotxt 
      Height          =   375
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
   Begin TextBox username 
      Height          =   375
      Left            =   1440
      MaxLength       =   24
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin SSCommand cmdQuit 
      Caption         =   "&Don't Register Me"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   6735
   End
   Begin SSCommand cmdOK 
      Caption         =   "&Register Me (I agree to be bound by the terms of the licence agreement)"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   4560
      Width           =   6855
   End
   Begin Label Label6 
      Caption         =   "Racing Pigeon Management software, Copyright 1988-2000 by Kieran G Malone, All Rights Reserved"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   8775
   End
   Begin Label Label5 
      Alignment       =   2  'Center
      Caption         =   "To use this software you must first register as the user. Refer to your software licence for this information, ensure you type in the same case as written on your software licence. (Last page of the manual)."
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1095
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   6015
   End
   Begin Label Label4 
      Caption         =   "Note: The LICENCE No. has been allocated to you by the supplier. It was derived from your Details so MUST match your Name as detailed on the software licence agreement."
      Height          =   975
      Left            =   4200
      TabIndex        =   7
      Top             =   2760
      Width           =   4455
   End
   Begin Label Label3 
      Caption         =   "Note: This is the name that you will race pigeons under it will appear on all reports. It MUST be the name that is on your Licence No."
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   1680
      Width           =   4455
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Licence No.:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
End
Option Explicit

Sub cmdDemoMode_Click ()
Dim reply As Integer
Panel3D1.Visible = True
reply% = MsgBox("Click on 'Yes' to setup 'Demonstration Mode' (see above for more detail)", 36, "Enter Demonstration mode?")
If reply% = 6 Then
    UserName.Text = "Demonstration Copy"
    LicNoTxt.Text = "2223190N"
Else
    UserName.SetFocus
End If
Panel3D1.Visible = False
End Sub

Sub cmdOK_Click ()

Dim LicNoa As String
Dim F As String
Dim V As String

GLDemo = -1
UserName.Text = Trim(UserName.Text)
LicNoTxt.Text = Trim(LicNoTxt.Text)

If Trim(UCase(UserName.Text)) = "Demonstration Copy" Then
    UserName.Text = UCase(UserName.Text)
End If

If Trim(UserName.Text) = "" Then
    MsgBox "Must provide User name to continue", 1, "Can't proceed"
    Exit Sub
End If

If Len(UserName.Text) < 6 Then
    MsgBox "Invalid User name - Contact Support", 1, "Can't proceed"
    Exit Sub
End If

If LicNoTxt.Text = "" Then
    MsgBox "Must provide Licence no. to continue", 1, "Can't proceed"
    Exit Sub
End If

F$ = UserName.Text
LicNoa$ = LicNo(F$)
If LicNoa$ = LicNoTxt.Text Then
    'find out if rp.dat exists
    V$ = Dir$(App.Path & "\" & "rp.dat")
    If V$ <> "" Then 'if it does then delete file
    End If
    'encode data and write rp.dat file
    If WrtRPDat(F$, LicNoa$) Then
        GLDemo = 0
        'If Demo set GLDemo to 9
        If Trim(UserName.Text) = "Demonstration Copy" Then
            GLDemo = 9
        End If
        Unload Me
        Exit Sub
    Else
        MsgBox "RP.Dat failed to write", 16, "Cannot register software"
    End If
Else
    MsgBox "The name you typed does not match the Licence No. you entered - Key in Name and Licence No. exactly as in manual, ensure the case is exactly the same OR refer to the supplier of the software", 16, "Cannot register software"
End If

End Sub

Sub cmdQuit_Click ()
GLDemo = -1
Unload Me
End Sub

Sub Form_Activate ()
Screen.MousePointer = 0
End Sub

Sub Form_Load ()
Panel3D1.Visible = False
End Sub

