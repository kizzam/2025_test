VERSION 2.00
Begin Form ApptForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Appointments"
   ClientHeight    =   5055
   ClientLeft      =   840
   ClientTop       =   1740
   ClientWidth     =   7320
   Height          =   5745
   Icon            =   APPTBOOK.FRX:0000
   Left            =   780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7320
   Top             =   1110
   Width           =   7440
   Begin TrueGrid ApptList 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      Editable        =   0   'False
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      HeadingHeight   =   1
      Height          =   2835
      HorzLines       =   1  'Single
      Layout          =   APPTBOOK.FRX:0302
      Left            =   120
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   16
      Top             =   240
      UseBookmarks    =   -1  'True
      Width           =   4275
      WrapCellPointer =   0   'False
   End
   Begin SSPanel Panel3D2 
      Alignment       =   6  'Center - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "Appointment Types"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1755
      Left            =   4500
      TabIndex        =   11
      Top             =   240
      Width           =   1635
      Begin SSPanel KindFrame 
         BevelInner      =   1  'Inset
         BorderWidth     =   2
         ForeColor       =   &H00C0C0C0&
         Height          =   675
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Tag             =   "Phone"
         Top             =   300
         Width           =   675
         Begin Image KindPict 
            Height          =   480
            Index           =   0
            Left            =   90
            Picture         =   APPTBOOK.FRX:0390
            Top             =   90
            Width           =   480
         End
      End
      Begin SSPanel KindFrame 
         BevelInner      =   1  'Inset
         BorderWidth     =   2
         ForeColor       =   &H00C0C0C0&
         Height          =   675
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Tag             =   "Travel"
         Top             =   1020
         Width           =   675
         Begin Image KindPict 
            Height          =   480
            Index           =   1
            Left            =   90
            Picture         =   APPTBOOK.FRX:0692
            Top             =   90
            Width           =   480
         End
      End
      Begin SSPanel KindFrame 
         BevelInner      =   1  'Inset
         BorderWidth     =   2
         ForeColor       =   &H00C0C0C0&
         Height          =   675
         Index           =   2
         Left            =   840
         TabIndex        =   13
         Tag             =   "Meeting"
         Top             =   300
         Width           =   675
         Begin Image KindPict 
            Height          =   480
            Index           =   2
            Left            =   90
            Picture         =   APPTBOOK.FRX:0994
            Top             =   90
            Width           =   480
         End
      End
      Begin SSPanel KindFrame 
         BevelInner      =   1  'Inset
         BorderWidth     =   2
         ForeColor       =   &H00C0C0C0&
         Height          =   675
         Index           =   3
         Left            =   840
         TabIndex        =   12
         Tag             =   "Lunch"
         Top             =   1020
         Width           =   675
         Begin Image KindPict 
            Height          =   480
            Index           =   3
            Left            =   90
            Picture         =   APPTBOOK.FRX:0C96
            Top             =   90
            Width           =   480
         End
      End
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "Panel3D1"
      ForeColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   4500
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
      Begin CommandButton SaveButton 
         Caption         =   "Save"
         Default         =   -1  'True
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   735
      End
      Begin TextBox ApptText 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   2535
      End
      Begin MaskEdBox ApptTime 
         BackColor       =   &H00C0C0C0&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Format          =   "hh:nn AM/PM"
         Height          =   315
         Left            =   1500
         Mask            =   "##:## ??"
         MaxLength       =   8
         PromptChar      =   "_"
         TabIndex        =   5
         Top             =   840
         Width           =   1155
      End
      Begin TextBox ApptType 
         BackColor       =   &H00C0C0C0&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1500
         TabIndex        =   4
         Top             =   480
         Width           =   1155
      End
      Begin Image Image1 
         Height          =   480
         Left            =   1140
         Picture         =   APPTBOOK.FRX:0F98
         Top             =   60
         Width           =   480
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "When:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   900
         TabIndex        =   6
         Top             =   840
         Width           =   555
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Appointment type:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin Image TrashCan 
      Height          =   480
      Left            =   6480
      Picture         =   APPTBOOK.FRX:129A
      Top             =   900
      Width           =   480
   End
   Begin Image TrashOpened 
      Height          =   480
      Left            =   1740
      Picture         =   APPTBOOK.FRX:159C
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin Image TrashClosed 
      Height          =   480
      Left            =   1260
      Picture         =   APPTBOOK.FRX:189E
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin Label DragArrow 
      Caption         =   "DragArrow"
      DragIcon        =   APPTBOOK.FRX:1BA0
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   915
   End
   Begin Label MoveIcon 
      Caption         =   "MoveIcon"
      DragIcon        =   APPTBOOK.FRX:1EA2
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   4380
      Visible         =   0   'False
      Width           =   915
   End
   Begin Label SaveIcon 
      Caption         =   "SaveIcon"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   915
   End
   Begin Label NoDrag 
      Caption         =   "NoDrag"
      DragIcon        =   APPTBOOK.FRX:21A4
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   3780
      Visible         =   0   'False
      Width           =   915
   End
   Begin Menu ExitMenuOption 
      Caption         =   "E&xit!"
   End
   Begin Menu HelpMenu 
      Caption         =   "&Help"
      Begin Menu HelpOption 
         Caption         =   "&Index"
         Index           =   1
      End
      Begin Menu HelpOption 
         Caption         =   "&Using Help"
         Index           =   2
      End
      Begin Menu HelpOption 
         Caption         =   "-"
         Index           =   3
      End
      Begin Menu HelpOption 
         Caption         =   "&About Appointment Book..."
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

' Drag mode constants to keep track of dragging activity.

Dim DragType As Integer         ' type of object being dragged
Dim Dragging As Integer         ' TRUE when dragging is in progress
Dim DragIndex As Integer        ' Optional index of dragged obj
Dim DragRow As Integer          ' Optional row being dragged in grid

' Miscellaneous variables

Dim valid%                      ' used as return for DragValid

' Bitmasks to describe valid drag objects

Const MASK_NEWAPPT = 1      ' a new appointment
Const MASK_OLDAPPT = 2      ' an old appointment
Const MASK_NONE = 0         ' mask used where no drops are allowed

Sub ApptEdit ()
    
    ' This subroutine moves the data in the current grid row into
    ' the "post-it" editing area.

    Dim aText As String
    Dim colonPos As Integer

    ' This routine copies appointment data to the edit window

    aText = ApptList.ColumnText(2)
    colonPos = InStr(aText, ":")

    ' If no colon, there's no appointment, so clear the post-it
    ' area.  If there is a colon, fill in the information.

    If colonPos = 0 Then
	ApptText.Text = ""
	ApptTime.Text = Format$(0, ApptTime.Format)
	ApptType.Text = ""
    Else
	ApptType.Text = Left$(aText, colonPos - 1)
	ApptText.Text = Mid$(aText, colonPos + 2)
	ApptTime.Text = Format$(ApptList.ColumnText(1), ApptTime.Format)
    End If

End Sub

Sub ApptList_ColumnChange ()

    ' This is to prevent the user from selecting the first column

    ApptList.ColumnIndex = 2
    
End Sub

Sub ApptList_DragCell (Split As Integer, Row As Long, Col As Integer)

    ' This event is intitiating when the user drags the mouse with
    ' the right mouse button depressed.  If the user is in column 2
    ' and there is an appointment then initiate dragging

    If Col = 2 Then
	ApptList.RowIndex = Row
	If InStr(ApptList.Text, ":") <> 0 Then
	    DragRow = ApptList.RowIndex
	    ApptList.RefreshRow = 0
	    ApptList.DragIcon = MoveIcon.DragIcon
	    BeginDragMode ApptList, MASK_OLDAPPT
	End If
    End If

End Sub

Sub ApptList_DragDrop (Source As Control, X As Single, Y As Single)
    
    ' Drop a new appointment or existing appointment at a new
    ' row position.

    Dim aText As String
    Dim i%

    If Not EndDragMode(MASK_NEWAPPT Or MASK_OLDAPPT) Then Exit Sub

    ' Check the location of the drag and drop

    ApptList.PointX = X
    ApptList.PointY = Y
    GridDropRow = ApptList.RowAtPoint

    ' If new appointment add to grid otherwise change appt time
    If DragType = MASK_NEWAPPT Then
	ApptList.RowIndex = GridDropRow
	ApptInfo(GridDropRow) = Source.Tag & ": "
	ApptList.RefreshRow = 0
	ApptEdit
    Else
	aText = ApptList.ColumnText(1)
	ApptList.RowIndex = DragRow
	i% = ChangeApptTime(TimeValue(aText))
    End If
    
    ' Clear drag row

    DragRow = 0
    ApptText.SetFocus
    
End Sub

Sub ApptList_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    
    ' When dragging over the grid, both new and old appointments
    ' are considered.

    If Not DragValid(Source, MASK_NEWAPPT Or MASK_OLDAPPT, State) Then
	Exit Sub
    End If
 
    ' Find location in the grid

    ApptList.PointX = X
    ApptList.PointY = Y

    ' If leaving the grid set the rowindex back to the where dragging
    ' started, else move the RowIndex with the drag icon.

    Select Case State
	Case LEAVE
	    ApptList.RowIndex = DragRow
	Case Else
	    GridDropRow = ApptList.RowAtPoint
	    
	    ' If RowAtPoint is valid then move the RowIndex

	    If ApptList.RowAtPoint <> -1 Then
		ApptList.RowIndex = ApptList.RowAtPoint
	    End If
    End Select

End Sub

Sub ApptList_Fetch (Row As Long, Col As Integer, Value As String)

    ' This event is the main display mechanism in callback mode
    ' If in column 1 display the times.  If in column 2 query the
    ' ApptInfo array and display all appointments

    If Col = 1 Then
	Value = ApptTimes(Row)
    ElseIf Col = 2 Then
	Value = ApptInfo(Row)
    End If
	    
End Sub

Sub ApptList_FetchAttributes (Status As Integer, Split As Integer, Row As Long, Col As Integer, FgColor As Long, BgColor As Long, FontStyle As Integer)

    ' This event lets us set attributes for given cells.  In this case
    ' if the current cell has an appointment then highlight it in
    ' green and white.

    If InStr(ApptList.Text, ":") Then
	If Status = GFS_CURCELL Then
	   BgColor = LTGREEN
	   FgColor = WHITE
	End If
    End If
	
    ' If dragging is occuring then highlight the cell being
    ' dragged from in blue and white.

    If DragRow = Row Then
	BgColor = BLUE
	FgColor = WHITE
    End If

End Sub

Sub ApptList_MarkChange (Row As Long, IsMarked As Integer)

    ' In TrueGrid Pro you can now use multiple selection
    ' callback mode.  The code in the MarkChange event
    ' causes a select array to be updated
    
    ApptSelect(Row) = ApptSelect(Row) Xor True
    ApptList.RefreshRow = Row
    
End Sub

Sub ApptList_QueryMark (Row As Long, IsMarked As Integer)

    ' Here the grid checks to see if it should highlight a
    ' given row.  We provide it our array that we maintain
    ' in MarkChange.

    IsMarked = ApptSelect(Row)


End Sub

Sub ApptList_RowChange ()
	
    ' If a RowChange occurs call ApptEdit to update Post-it area

    ApptEdit

End Sub

Sub ApptText_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub ApptText_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Sub ApptTime_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub ApptTime_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Sub ApptTime_ValidationError (InvalidText As String, StartPosition As Integer)

    ' Check for valid time if user manually changes time
    
    MsgBox "Invalid time"
    ApptTime.SetFocus

End Sub

Sub ApptType_DragDrop (Source As Control, X As Single, Y As Single)
    
    ' Accept a drop only for a NEWAPPT icon, otherwise the
    ' operation will be cancelled.

    If EndDragMode(MASK_NEWAPPT) Then
	ApptType.Text = Source.Tag
    End If

End Sub

Sub ApptType_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NEWAPPT, State)
End Sub

Sub ApptType_KeyPress (KeyAscii As Integer)
    
    ' Don't allow a colon to be entered, since we use a colon to
    ' separate the appointment "kind" from the text.

    If KeyAscii = Asc(":") Then
	Beep
	KeyAscii = 0
    End If

End Sub

Sub BeginDragMode (ctl As Control, objType As Integer)
    
    ' Whenever a drag is about to start, this routine is called.
    ' The type mask of the drag is flagged, and we remember that
    ' dragging is in progress.   This routine MUST be matched
    ' by an EndDragMode function call.

    DragType = objType
    Dragging = True

    ' Change the highlight color for the current cell to blue and white
    ' while dragging is occuring

    ApptList.ParamBackColor = BLUE
    ApptList.ParamForeColor = WHITE
    ApptList.ColumnSetStatusAttr(2) = GFS_CURCELL

    ' Start the drag process

    ctl.Drag BEGIN_DRAG

End Sub

Function ChangeApptTime (newtime As Variant) As Integer
    
    ' Given a new time for an appointment at the current row, this
    ' routine moves the appointment to the new location in the
    ' grid.

    Dim trow As Integer
    Dim oldAppt As String

    trow = TimeRow(newtime)

    ' If we're already there, then do nothing and return False,
    ' indicating no row change occurred.

    If trow = ApptList.RowIndex Then
	ChangeApptTime = False
	Exit Function
    End If

    ' If the user entered in invalid time print an error

    If trow > ApptList.Rows Or trow < 0 Then
	ChangeApptTime = False
	MsgBox "Invalid time"
	Exit Function
    End If
    
    ChangeApptTime = True
    
    ' Actually move the row.

    oldAppt = ApptList.Text
    ApptInfo(ApptList.RowIndex) = ""
    ApptList.RefreshRow = 0
    
    ApptList.RowIndex = trow
    ApptInfo(ApptList.RowIndex) = oldAppt
    ApptList.RefreshRow = 0

    ApptEdit            ' move the data to the post-it area

End Function

Function DragValid (src As Control, mask As Integer, State As Integer) As Integer
    
    ' This function is called by an object's DragOver event to
    ' automatically change the drag cursor to the "no drop"
    ' cursor if necessary.  It also returns True if the object
    ' can legally be dropped according to the input mask.

    If (mask And DragType) Then
	DragValid = True
	Exit Function
    End If

    ' This is not a valid drag.  Return False, but also change the
    ' object's drag icon to the NoDrag icon (remembering the old
    ' value for later restore when we exit this object).

    DragValid = False

    Select Case State
	
	Case ENTER

	    ' Entering, remember old icon

	    SaveIcon.DragIcon = src.DragIcon
	    src.DragIcon = NoDrag.DragIcon

	Case LEAVE
	    
	    ' Exiting, restore old icon

	    src.DragIcon = SaveIcon.DragIcon

    End Select
		
End Function

Function EndDragMode (mask As Integer) As Integer
    
    ' This function is called when a drag has ended, either
    ' successfully or unsuccessfully.  This routine removes any
    ' user feedback related to the drag operation and returns
    ' TRUE if the passed mask matches the dragged object.

    ' Set Current cell colors back to previous settings

    ApptList.ParamBackColor = INHERIT_COLOR
    ApptList.ParamForeColor = INHERIT_COLOR
    ApptList.ColumnSetStatusAttr(2) = GFS_CURCELL
    
    Select Case DragType

	Case MASK_NEWAPPT

	    ' If a "new appointment" icon was dragged, change the
	    ' frame background to LTGREY again so that the drag
	    ' is officially over.

	    KindFrame(DragIndex).BackColor = LTGREY

	Case MASK_OLDAPPT

	    ' If this is an item dragged from the grid, refresh
	    ' the grid in case the drag ended outside the grid
	    ' frame (and the inverted row remains).
    
    End Select

    Dragging = False
    EndDragMode = (mask And DragType) <> 0
    
End Function

Sub ExitMenuOption_Click ()

    'Unload forms and exit

    Unload About
    Unload ApptForm
    End

End Sub

Sub Form_DragDrop (Source As Control, X As Single, Y As Single)
    
    ' Ignore drops which occur on the form

    valid% = EndDragMode(MASK_NONE)

End Sub

Sub Form_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    
    ' Assure that the "no drop" icon is displayed when passing
    ' over the form.

    valid% = DragValid(Source, MASK_NONE, State)

End Sub

Sub Form_Load ()

    Dim curTime As Variant
    Dim curRow As Integer
    Dim rowMax As Integer
    
    ' Fill the leftmost column with appointment times.
    curRow = 1
    For curTime = Prefs.timeStart To Prefs.timeEnd Step Prefs.timeIncrement
	ApptTimes(curRow) = Format$(curTime, "hh:mm am/pm")
	curRow = curRow + 1
    Next curTime

    ' Set up ApptList
    
    ApptList.SelectMode = 1  ' Enable multiple selection
    rowMax = (Prefs.timeEnd - Prefs.timeStart) / Prefs.timeIncrement
    ApptList.Rows = rowMax
    ApptList.RowIndex = 1
    ApptList.Height = (ApptList.CellRectHeight + (Screen.TwipsPerPixelY + 1)) * rowMax

    ' ApptList column colors

    ApptList.ColumnBackColor(1) = LTGREY     'Grey background
    ApptList.ColumnBackColor(2) = OFF_YELLOW     'Yellow background
    
    ' Forces the FetchAttributes to be called for column 2

    ApptList.ColumnCellAttrs(2) = True

End Sub

Sub Form_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Since we can't trap a "drop" which occurs outside of our
    ' application, this is a pretty good solution.  Whenever the
    ' cursor passes over the form, if we're still dragging check
    ' to see if the button is now up.  If so, just cancel the
    ' operation

    If Dragging Then
	If (Button And LEFT_BUTTON) = 0 Then
	    valid% = EndDragMode(MASK_NONE)
	End If
    End If

End Sub

Sub HelpOption_Click (index As Integer)
    
    'This event calls the WinHelp EXE and a location to goto based on which selection the user has chosen
    'Case 4 shows the about box for True Browser
    Select Case index
	Case 1
	    HelpContext ApptForm, HELP_APPTBOOK
	Case 2
	    HelpOnHelp ApptForm
	Case 4
	    About.Show 1
    End Select


End Sub

Sub Image1_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub Image1_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Sub KindFrame_DragDrop (index As Integer, Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub KindFrame_DragOver (index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NEWAPPT, State)
End Sub

Sub KindPict_DragDrop (index As Integer, Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub KindPict_DragOver (index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NEWAPPT, State)
End Sub

Sub KindPict_MouseDown (index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' When the left button goes down over an "appointment type"
    ' icon, drag its image in NEWAPPT mode.  Copy the DragIcon
    ' each time, since it may still be set to the "no drop" icon
    ' from a previous cancellation.

    If Button And LEFT_BUTTON Then

	KindFrame(index).DragIcon = DragArrow.DragIcon
	BeginDragMode KindFrame(index), MASK_NEWAPPT
	KindFrame(index).BackColor = CYAN
	
	' Save the index, we'll need it in EndDragMode

	DragIndex = index
	ApptList.RowIndex = 0

    End If

End Sub

Sub Label1_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub Label1_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Sub Label2_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub Label2_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Sub Panel3D1_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub Panel3D1_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Sub Panel3D2_DragDrop (Source As Control, X As Single, Y As Single)
    valid% = EndDragMode(MASK_NONE)
End Sub

Sub Panel3D2_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)

End Sub

Sub SaveButton_Click ()
    
    ' Save all data in the post-it area to the grid.

    Dim i%
    
    ' We can only save if there's an appointment on the current
    ' grid row already (at least a blank one).

    If InStr(ApptList.ColumnText(2), ":") = 0 Then
	MsgBox "No appointment at current row"
	Exit Sub
    End If
    
    ApptInfo(ApptList.RowIndex) = ApptType.Text & ": " & ApptText.Text
    ApptList.RefreshRow = 0
    ApptList.RowIndex = ApptList.RowIndex
    IgnoreRowChange = False
    
    ' If the time was changed manually, then move the row to the new
    ' location.

    i% = ChangeApptTime(TimeValue(ApptTime.Text))

End Sub

Sub SaveButton_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    valid% = DragValid(Source, MASK_NONE, State)
End Sub

Function TimeRow (thetime As Variant) As Integer
    
    ' Given a time value, return the row number within the grid
    ' where the specified time slot is located.

	TimeRow = (thetime - TimeValue("08:00am")) / TimeValue("00:30") + 1

End Function

Sub TrashCan_DragDrop (Source As Control, X As Single, Y As Single)
    
    ' The trash can only accepts drops for "old appointments" from
    ' the grid.

    If EndDragMode(MASK_OLDAPPT) Then

	' Get rid of feedback

	TrashCan.Picture = TrashClosed.Picture

	' Clear the grid row and update the post-it area
	
	ApptInfo(ApptList.RowIndex) = ""
	ApptList.RefreshRow = 0
	DragRow = 0
	ApptEdit
	ApptList.SetFocus
    
    End If

End Sub

Sub TrashCan_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    
    ' Provide feedback by "opening the trashcan" whenever an
    ' old appointment is dragged over the trash.

    If DragValid(Source, MASK_OLDAPPT, State) Then
	Select Case State
	    Case ENTER
		' Open when entering
		TrashCan.Picture = TrashOpened.Picture
	    Case LEAVE
		' Close when leaving
		TrashCan.Picture = TrashClosed.Picture
	End Select
    End If
End Sub

