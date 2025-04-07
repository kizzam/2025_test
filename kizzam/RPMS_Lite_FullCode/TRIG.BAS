Option Explicit
'
' FUNCTIONS HEREIN:
'
' ARC SINE IN RADIANS   fnArcSineRad
'          OR DEGREES   fnArcSineDeg
'
'  RADIANS TO DEGREES   fnRadsToDegs
'  DEGREES TO RADIANS   fnDegsToRads
'
'
' NOTE: there is also a testing section below these functions.
'
'
'
Const mcPI = 3.14159265358979                                   ' Approximation of Pi.
'
'
'        Name:      ARC SINE IN RADIANS
'     Purpose:      This function determines the ArcSine of the passed argument.
'      Inputs:      vntSine     The sine of an angle.
' Assumptions:      None.
'     Effects:      Will raise "Overflow" error if vntSine is not a valid sine or
'                   if any other error occurs during execution.
'     Returns:      The angle in radians.
'                   If an error occurs fnArcSine is indeterminate.
'
'
'
'        Name:      ARC SINE IN DEGREES
'     Purpose:      This function determines the ArcSine of the passed argument.
'      Inputs:      vntSine     The sine of an angle.
' Assumptions:      None.
'     Effects:      None.
'     Returns:      The angle in degrees.
'                   If an error occurs fnArcSine is indeterminate.
'
'
'
'        Name:      DEGREES TO RADIANS
'     Purpose:      This function converts passed degrees into radians.
'      Inputs:      dblDegrees  A quantity in degrees.
' Assumptions:      Passed arg is a Double.
'     Effects:      None.
'     Returns:      The angle in radians.
'                   If an error occurs fnDegsToRads is indeterminate.
'
'
'
'        Name:      RADIANS TO DEGREES
'     Purpose:      This function converts passed radians into degrees.
'      Inputs:      dblRadians  A quantity in radians.
' Assumptions:      Passed arg is a Double.
'     Effects:      None.
'     Returns:      The angle in degrees.
'                   If an error occurs fnRadsToDegs is indeterminate.
'
'
'
'
'
'
' Convert (-infinity degrees to +infinity degrees) to -90 to +90.
' -90 to +90 is the range of values the correspond to ArcSine(-1) to ArcSine(1).
'
' Why?  Because the Sin(390) is 0.5 but the ArcSine(0.5) is 30.
' Therefore to test if ArcSin(Sin(390))=390 I must translate 390 to
' the range returned by ArcSin.
'
'
'
' Modulus with fractional part of remainder also.

Function ArcCos (X As Double) As Double
   ArcCos = .5 * 3.14159265358979 - ArcSin(X)
End Function

Function ArcSin (X As Double) As Double
Dim value0 As Double
Dim value1 As Double

   value0 = X 'argument
   value1 = 1
   If value0 < 0 Then
        value0 = -(value0)
        value1 = -1
   ElseIf value0 = 0 Then
          value1 = 0
   End If

   If value0 > 1 Then
         MsgBox "ILLEGAL ARCSIN ARGUMENT", 16, "FATAL ERROR"
         Exit Function
   ElseIf value0 = 1 Then
         X = value1 * .5 * 3.14159265358979
   Else
         X = X / Sqr(1 - X * X)
         X = Atn(X)
   End If
   ArcSin = X
End Function

'Function ATN2X1 (SinOfx As Double, CosOfx As Double) As Double
'  If CosOfx <> 0 Then
'     ATN2X1 = Atn(SinOfx / CosOfx)
'    Else
'     ATN2X1 = 1.570796327
'    End If
'End Function
Function ATN2X (y As Double, X As Double) As Double
  If X < 0 Then ATN2X = Atn(y / X) + mcPI
  If ((X > 0) And (y >= 0)) Then ATN2X = Atn(y / X)
  If ((X > 0) And (y < 0)) Then ATN2X = Atn(y / X) + 2 * mcPI
  If ((X = 0) And (y > 0)) Then ATN2X = mcPI / 2
  If ((X = 0) And (y < 0)) Then ATN2X = 3 * mcPI / 2
  If ((X = 0) And (y = 0)) Then MsgBox ("atan2(0,0) undefined"): Stop
End Function

Function fnArcSineDeg (vntSine As Variant) As Double
                                                                ' No error trap; allow fnArcSineRad errors to be seen.
    fnArcSineDeg = fnRadsToDegs(fnArcSineRad(vntSine))          ' Get Degree ArcSin by converting Radian ArcSin.

End Function

Function fnArcSineRad (vntSine As Variant) As Double
On Error GoTo ERROR_ArcSine                                     ' Trap strangeness.

Const cOVERFLOW = 6                                             ' "Overflow" message #.

Dim blnEditPassed As Integer ' Boolean VB5                                    ' Edit results variable.
Dim dblTemp       As Double                                     ' Temporary double.

    blnEditPassed = False                                       ' It hasn't passed yet!
    If IsNumeric(vntSine) Then                                  ' Is arg numeric?
        If vntSine >= -1 And vntSine <= 1 Then                  ' Yup, is it within limits?
            blnEditPassed = True                                ' Yup, employ Derived Math Function:
                                                                ' Arcsin(X) = Atn(X / Sqr(-X * X + 1))
            dblTemp = Sqr(-vntSine * vntSine + 1)               ' Calculate denominator.
            If dblTemp = 0 Then                                 ' Is it 0?  Can only happen if arg is +1 or -1.
                fnArcSineRad = Sgn(vntSine) * mcPI / 2          ' Yup, assign +Pi/2 or -Pi/2 based on sign of arg.
            Else
                fnArcSineRad = Atn(vntSine / dblTemp)           ' Complete derived math function.
            End If
        End If
    End If

EXIT__ArcSine:
    'If Not blnEditPassed Then Err.Raise cOVERFLOW   'VB5            ' Raise overflow error if arg failed edits.
    Exit Function

ERROR_ArcSine:
    On Error GoTo 0                                             ' Turn off error trapping.
    blnEditPassed = False                                       ' Should never happen!
    Resume EXIT__ArcSine                                        ' But if it does give overflow error.


End Function

Function fnDegsToRads (dblDegrees As Double) As Double
On Error Resume Next                                            ' Trap strangeness.

    fnDegsToRads = dblDegrees * mcPI / 180                      ' Convert.

End Function

Function fnRadsToDegs (dblRadians As Double) As Double
On Error Resume Next                                            ' Trap strangeness.

    fnRadsToDegs = dblRadians * 180 / mcPI                      ' Convert.

End Function

Function fnRealMod (dblReal As Double, lngModulus As Long) As Double
Dim dblTimes As Double                                          ' # of complete divisions.
    dblTimes = dblReal / lngModulus                             ' Divide.
    fnRealMod = dblReal - (Int(dblTimes) * lngModulus)          ' Get the real remainder.
End Function

Function fnSinRange (ByVal dblAngle As Double) As Double
    dblAngle = fnRealMod(dblAngle, 360)                         ' Remove extra cycles.
    If dblAngle < 0 Then dblAngle = 360 + dblAngle              ' Compensate for negatives.
    Select Case Int(dblAngle / 90)                              ' Find the quadrant.
        Case 0: fnSinRange = dblAngle                           ' 1st quadrant:   0 to  89.9999 =>   0 to  89.9999
        Case 1: fnSinRange = 180 - dblAngle                     ' 2nd quadrant:  90 to 179.9999 =>  90 to   0.0001
        Case 2: fnSinRange = -(dblAngle - 180)                  ' 3rd quadrant: 180 to 269.9999 =>   0 to -89.9999
        Case 3: fnSinRange = -(360 - dblAngle)                  ' 4th quadrant: 270 to 359.9999 => -90 to  -0.0001
    End Select
End Function

Function MakDeg (A As Double) As Double 'convert radians to degrees.
   MakDeg = A * (180 / 3.14159265358979)
End Function

Function MakRad (A As Double) As Double 'Convert degrees to Radians
    MakRad = A# * (3.14159265358979 / 180)
End Function

Function PI () As Double ' The value of PI to 15 places.
   PI = 3.14159265358979
End Function

