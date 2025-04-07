Option Strict Off
Option Explicit On
Imports System.IO
Imports Microsoft.VisualBasic.ControlChars

Friend Class frmSecurity
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCalc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCalc.Click

        screenreset()

        Dim resultLic As String

        Dim clsLic As New clsRPMSSecure

        resultLic = clsLic.CalcLicCode(txtUserName.Text)

        txtSecId.Text = Trim(Mid(resultLic.Trim, 1, 7))
        txtSecIDa.Text = Trim(Mid(resultLic.Trim, 8, 1))
		
		If txtLicNo.Text = Trim(txtSecId.Text & txtSecIDa.Text) Then
			txtStatus.Text = "Passed"
		Else
            txtStatus.Text = "Failed"
            If txtLicNo.Text = "" Then
                txtLicNo.Text = Trim(txtSecId.Text & txtSecIDa.Text)
            End If
        End If

        'Dim iUserName(24) As Short
        'Dim dLic As Double
        'Dim iX As Short
        'Dim sA As String

        'dLic = 0

        'txtUserName.Text = Trim(txtUserName.Text)
        'txtLicNo.Text = Trim(txtLicNo.Text)

        'For iX = 1 To Len(txtUserName.Text)
        '	iUserName(iX) = Asc(Mid(txtUserName.Text, iX, 1))
        '	If iX = 2 Then
        '		dLic = dLic * 12
        '	Else
        '		dLic = dLic + (iUserName(iX) * iX)
        '	End If
        'Next iX

        'sA = Mid(Str(dLic), 3, 2)

        'If Val(Mid(Str(dLic), 3, 2)) > Asc("A") And Val(Mid(Str(dLic), 3, 2)) < Asc("Z") Then
        '	txtSecIDa.Text = Chr(Val(Mid(Str(dLic), 3, 2)))
        'ElseIf Val(Mid(Str(dLic), 3, 3)) > Asc("a") And Val(Mid(Str(dLic), 3, 3)) < Asc("z") Then 
        '	txtSecIDa.Text = Chr(Val(Mid(Str(dLic), 3, 3)))
        'ElseIf Val(Mid(Str(dLic), 3, 2)) < 26 Then 
        '	txtSecIDa.Text = Chr(Val(Mid(Str(dLic), 3, 2)) + 65)
        'ElseIf Val(Mid(Str(dLic), 3, 2)) / 4 < 25 Then 
        '	txtSecIDa.Text = Chr(Val(CStr(CDbl(Mid(Str(dLic), 3, 2)) / 4)) + 65)
        'Else
        '	txtSecIDa.Text = "A"
        'End If

        'dLic = Int(dLic * 119.2) + 100000

        '      txtSecId.Text = Str(dLic)

		
	End Sub
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		Me.Close()
	End Sub

    Private Sub btnEncrypt_Click(sender As Object, e As EventArgs) Handles btnEncrypt.Click
        Dim clsLic As New clsRPMSSecure
        Dim myencrypt() As String = clsLic.encrypt_str(txtUserName.Text, txtLicNo.Text)
        txtUserE.Text = myencrypt(0)
        txtLicE.Text = myencrypt(1)

    End Sub


    Private Sub frmSecurity_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub btnDecrypt_Click(sender As Object, e As EventArgs) Handles btnDecrypt.Click
        Dim clsLic As New clsRPMSSecure
        Dim mydecrypt() As String = clsLic.decrypt_str(txtUserE.Text, txtLicE.Text)
        txtUsrD.Text = mydecrypt(0)
        txtLicD.Text = mydecrypt(1)
    End Sub

    Private Sub screenreset()
        txtLicD.Text = ""
        txtUsrD.Text = ""
        txtLicE.Text = ""
        txtUserE.Text = ""
        'txtUserName.Text = ""
        'txtLicNo.Text = ""
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        screenreset()
    End Sub

    Private Sub btnWrtRpdat_Click(sender As Object, e As EventArgs) Handles btnWrtRpdat.Click
        Dim clsrpsec As New clsRPMSSecure
        Dim myArrv As Boolean = clsrpsec.RPdat_Write(txtUserE.Text & txtLicE.Text)
    End Sub

    Private Sub btnReadRPDat_Click(sender As Object, e As EventArgs) Handles btnReadRPDat.Click
        Dim clsrpsec As New clsRPMSSecure
        Dim mydecrypt As String() = clsrpsec.rpdat_read()
        txtUsrE1.Text = mydecrypt(0)
        txtLicE1.Text = mydecrypt(1)

    End Sub

    Private Sub btnDecryptFile_Click(sender As Object, e As EventArgs) Handles btnDecryptFile.Click
        Dim clsLic As New clsRPMSSecure
        Dim mydecrypt() As String = clsLic.decrypt_str(txtUsrE1.Text, txtLicE1.Text)
        txtUsrDe.Text = mydecrypt(0)
        txtLicDe.Text = mydecrypt(1)

    End Sub
End Class