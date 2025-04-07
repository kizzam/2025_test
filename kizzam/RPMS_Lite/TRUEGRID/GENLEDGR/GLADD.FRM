VERSION 2.00
Begin Form AddForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Account..."
   ClientHeight    =   1665
   ClientLeft      =   2430
   ClientTop       =   3165
   ClientWidth     =   5775
   Height          =   2070
   Left            =   2370
   LinkTopic       =   "Form2"
   ScaleHeight     =   1665
   ScaleWidth      =   5775
   Top             =   2820
   Width           =   5895
   Begin CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3300
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   4500
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin TextBox Desc 
      Height          =   285
      Left            =   1380
      TabIndex        =   3
      Top             =   540
      Width           =   4215
   End
   Begin TextBox AccountID 
      Height          =   285
      Left            =   1380
      TabIndex        =   2
      Top             =   180
      Width           =   1575
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Description:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   1155
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Account #:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Option Explicit

Sub Command1_Click ()
    
    On Error GoTo ermsg
    ' Create a new account entry and display an error
    ' message if there's an error (such as a duplicate).

    Main.Data1.Recordset.AddNew
    Main.Data1.Recordset!AccountID = Val(AccountID)
    Main.Data1.Recordset!AccountName = Desc
    Main.Data1.Recordset.Update

    ' Refresh the grid so the account appears in proper
    ' sorted order, then position to the account we just
    ' added.

    Main.Data1.Refresh
    Main.Data1.Recordset.FindFirst "AccountID = " & AccountID
    Unload AddForm
    Exit Sub
ermsg:
    MsgBox Error$
    Exit Sub

End Sub

Sub Command2_Click ()
    
    Unload AddForm

End Sub

