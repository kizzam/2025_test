Option Explicit

Function TgFindColumn (ctl As TrueGrid, ByVal txt As String) As Integer
    
    ' ==========================================================
    ' Sub: TgFindColumn
    '
    ' Given a string which represents a field name (or a heading)
    ' this routine returns the column index of the column, or
    ' zero.
    ' ==========================================================

    Dim i As Integer

    TgFindColumn = 0
    txt = UCase$(txt)

    ' Search fields first

    For i = 1 To ctl.Columns
	If txt = UCase$(ctl.ColumnField(i)) Then
	    TgFindColumn = i
	    Exit Function
	End If
    Next

    ' Now search for headings

    For i = 1 To ctl.Columns
	If txt = UCase$(ctl.ColumnName(i)) Then
	    TgFindColumn = i
	    Exit Function
	End If
    Next

End Function

Function TgGetVisibleCols (ctl As TrueGrid) As Integer
    ' ==========================================================
    ' Sub: TgGetVisibleCols
    '
    ' Returns the number of columns which are visible in the
    ' current split.  The setting of SplitPropsGlobal is
    ' irrelevant since ColumnVisible always RETURNS the
    ' value for the current split.
    ' ==========================================================

    Dim ret As Integer
    Dim i As Integer

    ret = 0

    For i = 1 To ctl.Columns
	If ctl.ColumnVisible(i) Then ret = ret + 1
    Next i

    TgGetVisibleCols = ret

End Function

Sub TgLockColumn (ctl As TrueGrid, ByVal col As Integer, atleft As Integer)
    
    ' ==========================================================
    ' Sub: TGLockColumn
    '
    ' Locks a column in the leftmost or rightmost split,
    ' depending upon the value of atleft.  The column will
    ' be hidden in any other split it is present in.  If there
    ' is no split, then one will be created.
    ' ==========================================================

    Dim oldSplitIndex As Integer
    Dim oldPropsGlobal As Integer
    Dim newSplit As Integer
    Dim i As Integer

    ' Save the existing split index and global setting

    oldSplitIndex = ctl.SplitIndex
    oldPropsGlobal = ctl.SplitPropsGlobal

    ctl.SplitPropsGlobal = False

    ' First, determine whether or not we need to create a new
    ' split for locked columns.

    If atleft Then
	newSplit = 1
    Else
	newSplit = ctl.Splits
    End If

    ctl.SplitIndex = newSplit

    If ctl.SplitSizeMode <> GSPT_COLUMNS Then
	' Insert a split at the left and modify the caller's original
	' split index.
	
	If atleft Then
	    ctl.InsertSplit = 1
	    oldSplitIndex = oldSplitIndex + 1
	Else
	    ctl.InsertSplit = ctl.Splits + 1
	    newSplit = newSplit + 1
	End If

	' Hide all columns in the newly created split

	For i = 1 To ctl.Columns
	    ctl.ColumnVisible(i) = False
	Next i
    End If

    ' Make the column visible in the new split, make sure it's
    ' invisible elsewhere, but don't touch splits which are
    ' set to a fixed number of columns.

    ctl.ColumnVisible(col) = True

    For i = 1 To ctl.Splits
	ctl.SplitIndex = i
	If i <> newSplit And ctl.SplitSizeMode <> GSPT_COLUMNS Then
	    ctl.ColumnVisible(col) = False
	End If
    Next i

    ' Now, set the sizemode and size for the split.  Lock the split
    ' so that it has no split bar

    ctl.SplitIndex = newSplit
    ctl.SplitSizeMode = GSPT_COLUMNS
    ctl.SplitSize = TgGetVisibleCols(ctl)
    ctl.SplitLocked = True

    ' Reset the old split number and the global properties flag

    ctl.SplitIndex = oldSplitIndex
    ctl.SplitPropsGlobal = oldPropsGlobal

End Sub

Sub TgLockColumnLeft (ctl As TrueGrid, ByVal col As Integer)

    ' ==========================================================
    ' Sub: TgLockColumnLeft
    '
    ' Locks a column in the leftmost split of the grid,
    ' and creates a new split there for locked columns if
    ' one doesn't exist.
    ' ==========================================================

    TgLockColumn ctl, col, True

End Sub

Sub TgLockColumnRight (ctl As TrueGrid, ByVal col As Integer)

    ' ==========================================================
    ' Sub: TgLockColumnRight
    '
    ' Locks a column in the rightmost split of the grid,
    ' and creates a new split there for locked columns if
    ' one doesn't exist.
    ' ==========================================================

    TgLockColumn ctl, col, False

End Sub

Sub TgSetCurCellColor (ctl As TrueGrid, fg As Long, bg As Long)

    ' ==========================================================
    ' Sub: TgSetCurCellColor
    '
    ' This routine sets the color of the current cell.  If you
    ' want the entire highlighted row marquee to be a different
    ' color, then execute TgSetMarqueeColor FIRST.
    ' ==========================================================

    ' Set parameters

    ctl.ParamForeColor = fg
    ctl.ParamBackColor = bg

    ' Set the color for all possible current cell combinations,
    ' but not for other cells in the highlighted row

    ctl.SetStatusAttr = GFS_CURCELL
    ctl.SetStatusAttr = GFS_CURCELL + GFS_SELECTED
    ctl.SetStatusAttr = GFS_CURCELL + GFS_CHANGED
    ctl.SetStatusAttr = GFS_CURCELL + GFS_CHANGED + GFS_SELECTED
    ctl.SetStatusAttr = GFS_HIGHROW + GFS_CURCELL
    ctl.SetStatusAttr = GFS_HIGHROW + GFS_CURCELL + GFS_SELECTED
    ctl.SetStatusAttr = GFS_HIGHROW + GFS_CURCELL + GFS_CHANGED
    ctl.SetStatusAttr = GFS_HIGHROW + GFS_CURCELL + GFS_CHANGED + GFS_SELECTED
    
End Sub

Sub TgSetMarqueeColor (ctl As TrueGrid, fg As Long, bg As Long)

    ' ==========================================================
    ' Sub: TgSetMarqueeColor
    '
    ' This routine sets the color of the marquee, including
    ' the current cell within the marquee.  If you want to
    ' make the current cell a different color, then execute
    ' TgSetCurCellColor AFTER executing this subroutine.
    ' ==========================================================

    ' Set parameters first

    ctl.ParamForeColor = fg
    ctl.ParamBackColor = bg
    ctl.ParamFontStyle = -1

    ' Set the color for all possible marquee status combinations

    ctl.SetStatusAttr = GFS_CURCELL
    ctl.SetStatusAttr = GFS_CURCELL + GFS_HIGHROW
    ctl.SetStatusAttr = GFS_CURCELL + GFS_HIGHROW + GFS_SELECTED
    ctl.SetStatusAttr = GFS_HIGHROW
    ctl.SetStatusAttr = GFS_HIGHROW + GFS_SELECTED

    ' For changed data, then maintain the foreground color

    ctl.ParamForeColor = -1

    ctl.SetStatusAttr = GFS_CHANGED + GFS_CURCELL
    ctl.SetStatusAttr = GFS_CHANGED + GFS_CURCELL + GFS_HIGHROW
    ctl.SetStatusAttr = GFS_CHANGED + GFS_CURCELL + GFS_HIGHROW + GFS_SELECTED
    ctl.SetStatusAttr = GFS_CHANGED + GFS_HIGHROW
    ctl.SetStatusAttr = GFS_CHANGED + GFS_HIGHROW + GFS_SELECTED

End Sub

Sub TgUnlockColumn (ctl As TrueGrid, ByVal col As Integer)
    
    ' ==========================================================
    ' Sub: TgUnlockColumn
    '
    ' Unlocks a column which was locked with TgLockColumnLeft
    ' or TgLockColumnRight.  A column is considered locked if
    ' it is part of a locked split which has SplitSizeMode
    ' set to GSPT_COLUMNS.
    ' ==========================================================

    Dim oldSplitIndex As Integer
    Dim oldPropsGlobal As Integer
    Dim lockSplit As Integer
    Dim visible As Integer
    Dim i As Integer

    ' Save the existing split index and global settings

    oldSplitIndex = ctl.SplitIndex
    oldPropsGlobal = ctl.SplitPropsGlobal

    ctl.SplitPropsGlobal = False

    ' Find the split which contains the locked column

    ctl.SplitIndex = 1
    Do While ctl.SplitIndex <= ctl.Splits
	If ctl.SplitSizeMode = GSPT_COLUMNS And ctl.SplitLocked Then
	    If ctl.ColumnVisible(col) Then Exit Do
	End If

	' Exit the sub if we didn't find the column

	If ctl.SplitIndex = ctl.Splits Then
	    ctl.SplitIndex = oldSplitIndex
	    ctl.SplitPropsGlobal = oldPropsGlobal
	    Exit Sub
	End If
	ctl.SplitIndex = ctl.SplitIndex + 1
    Loop

    lockSplit = ctl.SplitIndex

    ' The locked column is in the current split.  Hide it in the
    ' current split, and unhide it in all other splits which aren't
    ' set to a fixed number of columns.

    ctl.ColumnVisible(col) = False
    For i = 1 To ctl.Splits
	ctl.SplitIndex = i
	If i <> lockSplit And ctl.SplitSizeMode <> GSPT_COLUMNS Then
	    ctl.ColumnVisible(col) = True
	End If
    Next i

    ' If there are no columns visible in the current split, then remove
    ' the split

    ctl.SplitIndex = lockSplit
    If TgGetVisibleCols(ctl) = 0 Then
	ctl.RemoveSplit = lockSplit
	If oldSplitIndex > lockSplit Or oldSplitIndex > ctl.Splits Then oldSplitIndex = oldSplitIndex - 1
    End If

    ctl.SplitIndex = oldSplitIndex
    ctl.SplitPropsGlobal = oldPropsGlobal

End Sub

Sub TgSetNegativeColColor (ctl As TrueGrid, col As Integer, fg As Long, bg As Long)
    
    ' ==========================================================
    ' Sub: SetNegativeColColor
    '
    ' For a given column, sets the colors to be used for cells
    ' which appear to have negative numbers (start with a leading
    ' minus sign, or are enclosed in parenthesis)
    ' ==========================================================

    ctl.ParamForeColor = fg
    ctl.ParamBackColor = bg
    ctl.ParamFontStyle = -1
    ctl.ParamStatus = -1

    ctl.ColumnAddRegexAttr(col) = "^ *-"
    ctl.ColumnAddRegexAttr(col) = "^(.+)$"

End Sub

