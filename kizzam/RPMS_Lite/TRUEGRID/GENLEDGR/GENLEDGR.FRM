VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "General Ledger"
   ClientHeight    =   5250
   ClientLeft      =   1035
   ClientTop       =   1740
   ClientWidth     =   8475
   Height          =   5940
   Icon            =   GENLEDGR.FRX:0000
   Left            =   975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8475
   Top             =   1110
   Width           =   8595
   Begin CommandButton Command1 
      Caption         =   "Add Account"
      Height          =   315
      Left            =   2250
      TabIndex        =   1
      Top             =   4800
      Width           =   1995
   End
   Begin TrueGrid Table1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H0080FFFF&
      DataSource      =   "Data1"
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   1  'Expose upon editing
      FetchMode       =   0  'By cell
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeadingHeight   =   1
      Height          =   4635
      HorzLines       =   1  'Single
      Layout          =   GENLEDGR.FRX:073A
      Left            =   75
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   0   'False
      SplitTabMode    =   1  'Tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   75
      UseBookmarks    =   -1  'True
      Width           =   8295
      WrapCellPointer =   0   'False
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "GENLEDGR.MDB"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   75
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "select * from Accounts order by AccountID"
      Top             =   4800
      Width           =   1995
   End
   Begin Menu ExitMenuOption 
      Caption         =   "E&xit!"
   End
   Begin Menu HelpMenuOption 
      Caption         =   "&Help"
      Begin Menu mHelpOption 
         Caption         =   "&Index"
         Index           =   1
      End
      Begin Menu mHelpOption 
         Caption         =   "&Using Help"
         Index           =   2
      End
      Begin Menu mHelpOption 
         Caption         =   "-"
         Index           =   3
      End
      Begin Menu mHelpOption 
         Caption         =   "&About General Ledger..."
         Index           =   4
      End
   End
End
' ---------------------------------------------------------
'       Copyright (C) 1993 Apex Software Corporation
'
' You have a royalty-free right to use, modify, reproduce,
' and distribute the True Grid sample application files
' (and/or any modified version) in any way you find useful,
' provided that you agree that Apex Software Corporation
' has no warranty, obligations, or liability for any sample
' application files.
' ---------------------------------------------------------

Option Explicit

Const COL_BALANCE = 2

Dim col_account As Integer

Sub CenterForm (F As Form)

' Center the specified form within the screen

    F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Sub Command1_Click ()
    
    ' Bring up the "Add Account" form

    AddForm.Show MODAL

End Sub

Sub ExitMenuOption_Click ()

    Unload AddForm
    Unload Main
    End

End Sub

Function FindCol (hdtxt As String) As Integer
    
    ' Given a text string, this routine finds the column
    ' which contains the given text in the heading.  This
    ' is often more convenient than coordinating column numbers
    ' from design mode.

    Dim i As Integer
    For i = 1 To Table1.Columns
        If Table1.ColumnName(i) = hdtxt Then
            FindCol = i
            Exit Function
        End If
    Next
    FindCol = 0

End Function

Sub Form_Load ()
    Dim dbdir As String

    ' Set the database property of the data control
    dbdir = App.Path
    Data1.DatabaseName = dbdir + "\genledgr.mdb"
    
    ' In the leftmost split, set the general ledger number
    ' so that any value which contains solely numbers (no
    ' decimals) will appear in boldface

    Table1.SplitIndex = 1
    Table1.ParamFontStyle = 2      ' bold
    Table1.ColumnAddRegexAttr(FindCol("GL#")) = "^[0-9]+$"

    ' For the center split, turn margin fetching on for the
    ' account title

    Table1.SplitIndex = 2
    col_account = FindCol("Account")
    Table1.ColumnFetchMargins(col_account) = True

    'Center Form on Application
    CenterForm Main

End Sub

Sub mHelpOption_Click (index As Integer)
    
    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for True Browser
    Select Case index
        Case 1
            HelpContext Main, HELP_GENERALEDGER
        Case 2
            HelpOnHelp Main
        Case 4
            About.Show 1
    End Select

End Sub

Sub Table1_FetchMargins (Split As Integer, Row As Long, Col As Integer, InsetLeft As Integer, InsetRight As Integer)
    
    ' If there is a GL subcode in the GL# column, then set
    ' the left indent for the cell to 2 characters (20 units)

    If InStr(Table1.ColumnText(1), ".") <> 0 Then InsetLeft = 20

End Sub

Sub Table1_Update (Row As Long, Col As Integer, Value As String)
    
    ' When updating the Debit and Credit columns in split 3,
    ' update the actual field (the hidden balance column) with
    ' either a postitive or negative number and also update the
    ' last posting field.

    If Table1.SplitIndex = 3 Then
        If Table1.ColumnName(Col) = "Debit" Then
            Table1.ColumnText(COL_BALANCE) = Val(Value)
            Table1.ColumnText(FindCol("LastPosting")) = Now
        ElseIf Table1.ColumnName(Col) = "Credit" Then
            Table1.ColumnText(COL_BALANCE) = -Val(Value)
            Table1.ColumnText(FindCol("LastPosting")) = Now
        End If
    End If

End Sub

