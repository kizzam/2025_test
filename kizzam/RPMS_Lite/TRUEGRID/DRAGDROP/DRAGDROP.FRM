VERSION 2.00
Begin Form Main 
   Caption         =   "Drag and Drop Sample"
   ClientHeight    =   4545
   ClientLeft      =   1005
   ClientTop       =   1740
   ClientWidth     =   7305
   Height          =   5235
   Icon            =   DRAGDROP.FRX:0000
   Left            =   945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7305
   Top             =   1110
   Width           =   7425
   Begin TrueGrid Table2 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data2"
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeadingHeight   =   1.37
      Height          =   2055
      HorzLines       =   0  'None
      Layout          =   DRAGDROP.FRX:0302
      Left            =   120
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   1
      Top             =   2340
      UseBookmarks    =   -1  'True
      Width           =   7035
      WrapCellPointer =   0   'False
   End
   Begin Data Data2 
      Caption         =   "Data2"
      Connect         =   ""
      DatabaseName    =   "C:\DEV\TGTIPS\TASKS.MDB"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   2130
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Projects"
      Top             =   4380
      Visible         =   0   'False
      Width           =   1755
   End
   Begin TrueGrid Table1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data1"
      DragIcon        =   DRAGDROP.FRX:03DF
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      HeadingHeight   =   1.37
      Height          =   2115
      HorzLines       =   0  'None
      Layout          =   DRAGDROP.FRX:06E1
      Left            =   120
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   120
      UseBookmarks    =   -1  'True
      Width           =   7035
      WrapCellPointer =   0   'False
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "C:\DEV\TGTIPS\PHONE.MDB"
      Exclusive       =   0   'False
      Height          =   315
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "PhoneList"
      Top             =   4380
      Visible         =   0   'False
      Width           =   1755
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
         Caption         =   "&About Drag and Drop..."
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

Sub CenterForm (F As Form)

    ' Center the specified form within the screen

    F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Sub ExitMenuOption_Click ()

    Unload Main
    End

End Sub

Sub Form_Load ()
    Dim dbdir As String

    'Set the databasename property of the data controls
    dbdir = App.Path
    Data1.DatabaseName = dbdir + "\phone.mdb"
    Data2.DatabaseName = dbdir + "\tasks.mdb"

    'Centers the form on the screen
    CenterForm Main
    
End Sub

Sub mHelpOption_Click (index As Integer)
    
    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for True Browser
    
    Select Case index
        Case 1
            HelpContext Main, HELP_DRAGDROP
        Case 2
            HelpOnHelp Main
        Case 4
            About.Show 1
    End Select


End Sub

Sub Table1_DragCell (Split As Integer, Row As Long, Col As Integer)
    
    ' Begin dragging the cell we've pointed to.  In this case
    ' we only want to drag if the mouse is positioned over the
    ' second row.

    If Col = 2 Then
        Table1.RowIndex = Row
        Table1.ColumnIndex = Col
        Table1.Drag 1
    End If
End Sub

Sub Table2_Append ()
    
    ' Add a new project

    If MsgBox("Add a new project?", 1) = 2 Then Exit Sub
    Data2.Recordset.AddNew
    Data2.Recordset("Project") = "New..."

End Sub

Sub Table2_DragDrop (Source As Control, X As Single, Y As Single)
    
    ' We always take the person's name and
    ' phone number from the first grid and
    ' drop it into the proper cells in the
    ' second grid , but we really can do anything
    ' we want here.

    ' Ask if user wants to replace current manager

    If Table2.ColumnText(Table2.ColumnIndex) <> "" Then
        ret% = MsgBox("Are you sure you want to replace current manager?", 1, "Warning")
        If ret% = 2 Then
            Exit Sub
        End If
    End If

    Table2.ColumnText(2) = Table1.ColumnText(2)
    Table2.ColumnText(3) = Table1.ColumnText(3)
    Data2.Recordset.Update

End Sub

Sub Table2_DragOver (Source As Control, X As Single, Y As Single, state As Integer)
    
    ' Set the cell to the one we're pointing at

    Table2.PointX = X
    Table2.PointY = Y
    If Table2.RowAtPoint > 0 Then Table2.RowIndex = Table2.RowAtPoint
    If Table2.ColumnAtPoint Then Table2.ColumnIndex = Table2.ColumnAtPoint
    
End Sub

