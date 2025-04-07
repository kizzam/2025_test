Imports System.IO
Public Class clsRPMSSecure
    Public Function CalcLicCode(ByVal userName As String) As String

        Dim iUserName(24) As Short
        Dim dLic As Double = 0
        Dim sA As String
        Dim LicCode As String

        For cnt As Short = 1 To Len(userName)
            iUserName(cnt) = Asc(Mid(userName, cnt, 1))
            If cnt = 2 Then
                dLic = dLic * 12
            Else
                dLic = dLic + (iUserName(cnt) * cnt)
            End If
        Next cnt

        sA = Mid(Str(dLic), 3, 2)

        If Val(Mid(Str(dLic), 3, 2)) > Asc("A") And Val(Mid(Str(dLic), 3, 2)) < Asc("Z") Then
            LicCode = Chr(Val(Mid(Str(dLic), 3, 2)))
        ElseIf Val(Mid(Str(dLic), 3, 3)) > Asc("a") And Val(Mid(Str(dLic), 3, 3)) < Asc("z") Then
            LicCode = Chr(Val(Mid(Str(dLic), 3, 3)))
        ElseIf Val(Mid(Str(dLic), 3, 2)) < 26 Then
            LicCode = Chr(Val(Mid(Str(dLic), 3, 2)) + 65)
        ElseIf Val(Mid(Str(dLic), 3, 2)) / 4 < 25 Then
            LicCode = Chr(Val(CStr(CDbl(Mid(Str(dLic), 3, 2)) / 4)) + 65)
        Else
            LicCode = "A"
        End If

        dLic = Int(dLic * 119.2) + 100000

        CalcLicCode = Str(dLic) & LicCode.Trim
    End Function

    Public Function RPdat_Write(ByVal MyStr As String) As Boolean

        ' Write this array to the file.
        Dim array() As Char = MyStr.ToCharArray

        ' Create the BinaryWriter and use File.Open to create the file.
        Using writer As BinaryWriter = New BinaryWriter(File.Open("rp.dat", FileMode.Create))
            ' Write each integer.
            For Each value As Char In array
                writer.Write(value)
            Next
        End Using
        Return True

    End Function

    Public Function rpdat_read() As String()
        ' Create the reader in a Using statement.
        ' ... Use File.Open to open the existing binary file.
        Dim usr As String
        Dim lic As String
        Dim data As String = ""
        Using reader As New BinaryReader(File.Open("rp.dat", FileMode.Open))
            ' Loop through length of file.
            Dim pos As Integer = 0
            Dim length As Integer = reader.BaseStream.Length
            While pos < length - 5
                ' Read the integer.
                Dim value As Char = reader.ReadChar
                data = data & value.ToString
                ' Write to screen.
                Console.WriteLine(value)
                ' Add length of integer in bytes to position.
                pos += 2
            End While
        End Using
        usr = Mid(data, 1, 24)
        lic = Mid(data, 25, 8)
        Dim ret(1) As String
        ret(0) = usr
        ret(1) = lic
        Return ret
    End Function

    Public Function encrypt_str(ByVal MyUsr As String, MyLic As String) As String()
        Dim c1 As Integer
        Dim usr As String = MyUsr.Trim
        usr = usr.PadRight(24)
        Dim retusr As String = ""
        Dim retlic As String = ""

        For i As Integer = 0 To 23 '24 char OwnerName
            'C1 = Asc(data(i))
            'C1 = Asc(filedata(i))
            c1 = Asc(usr(i))
            retusr = retusr + Chr(c1 Xor &HFF)
        Next

        'Process Lic
        Dim lic As String = MyLic.Trim
        lic = lic.PadRight(24)
        For i As Integer = 0 To 7 '8 char OwnerName
            'C1 = Asc(data(i))
            'C1 = Asc(filedata(i))
            c1 = Asc(MyLic(i))
            retlic = retlic + Chr(c1 Xor &HFF)
        Next
        Dim u(1) As String
        u(0) = retusr.ToCharArray
        u(1) = retlic.ToCharArray
        Return u
    End Function

    Public Function decrypt_str(ByVal MyUsr As String, MyLic As String) As String()
        Dim c1 As Integer
        Dim usr As String = MyUsr
        'usr = usr.PadRight(24)
        Dim retusr As String = ""
        Dim retlic As String = ""

        For i As Integer = 0 To 23 '24 char OwnerName
            'C1 = Asc(data(i))
            'C1 = Asc(filedata(i))
            c1 = Asc(usr(i))
            retusr = retusr + Chr(c1 Xor &HFF)
        Next

        'Process Lic
        Dim lic As String = MyLic
        'lic = lic.PadRight(24)
        For i As Integer = 0 To 7 '8 char OwnerName
            'C1 = Asc(data(i))
            'C1 = Asc(filedata(i))
            c1 = Asc(MyLic(i))
            retlic = retlic + Chr(c1 Xor &HFF)
        Next
        Dim u(1) As String
        u(0) = retusr.ToCharArray
        u(1) = retlic.ToCharArray
        Return u
    End Function

End Class
