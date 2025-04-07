VERSION 5.00
Begin VB.Form frmSecurity 
   Caption         =   "Security"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtSecIDa 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtSecId 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtLicNo 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1440
      MaxLength       =   24
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Secutiry Tag:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Security ID:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Lic No.:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()

Dim iUserName(24) As Integer
Dim dLic As Double
Dim iX As Integer
Dim sA As String

dLic# = 0

txtUserName.Text = Trim(txtUserName.Text)
txtLicNo.Text = Trim(txtLicNo.Text)

For iX = 1 To Len(txtUserName.Text)
    iUserName(iX) = Asc(Mid$(txtUserName.Text, iX, 1))
    If iX% = 2 Then
        dLic# = dLic# * 12
    Else
        dLic# = dLic# + (iUserName(iX) * iX)
    End If
Next iX

sA$ = Mid$(Str(dLic#), 3, 2)

If Val(Mid(Str(dLic#), 3, 2)) > Asc("A") And _
    Val(Mid(Str(dLic#), 3, 2)) < Asc("Z") Then
        txtSecIDa.Text = Chr(Val(Mid(Str(dLic#), 3, 2)))
ElseIf Val(Mid(Str(dLic#), 3, 3)) > Asc("a") And _
    Val(Mid(Str(dLic#), 3, 3)) < Asc("z") Then
        txtSecIDa.Text = Chr(Val(Mid(Str(dLic#), 3, 3)))
ElseIf Val(Mid(Str(dLic#), 3, 2)) < 26 Then _
        txtSecIDa.Text = Chr(Val(Mid(Str(dLic#), 3, 2)) + 65)
ElseIf Val(Mid(Str(dLic#), 3, 2)) / 4 < 25 Then _
        txtSecIDa.Text = Chr(Val((Mid(Str(dLic#), 3, 2)) / 4) + 65)
Else
        txtSecIDa.Text = "A"
End If

dLic# = Int(dLic# * 119.2) + 100000

txtSecId.Text = Str(dLic#)

If txtLicNo.Text = Trim(txtSecId.Text & txtSecIDa) Then
    txtStatus.Text = "Passed"
Else
    txtStatus.Text = "Failed"
End If


End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
