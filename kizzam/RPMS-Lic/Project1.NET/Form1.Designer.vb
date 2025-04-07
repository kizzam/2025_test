<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSecurity
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtStatus As System.Windows.Forms.TextBox
	Public WithEvents txtSecIDa As System.Windows.Forms.TextBox
	Public WithEvents cmdCalc As System.Windows.Forms.Button
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents txtSecId As System.Windows.Forms.TextBox
	Public WithEvents txtLicNo As System.Windows.Forms.TextBox
	Public WithEvents txtUserName As System.Windows.Forms.TextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.txtSecIDa = New System.Windows.Forms.TextBox()
        Me.cmdCalc = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.txtSecId = New System.Windows.Forms.TextBox()
        Me.txtLicNo = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnEncrypt = New System.Windows.Forms.Button()
        Me.txtUserE = New System.Windows.Forms.TextBox()
        Me.txtLicE = New System.Windows.Forms.TextBox()
        Me.btnDecrypt = New System.Windows.Forms.Button()
        Me.txtUsrD = New System.Windows.Forms.TextBox()
        Me.txtLicD = New System.Windows.Forms.TextBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnWrtRpdat = New System.Windows.Forms.Button()
        Me.btnReadRPDat = New System.Windows.Forms.Button()
        Me.txtLicE1 = New System.Windows.Forms.TextBox()
        Me.txtUsrE1 = New System.Windows.Forms.TextBox()
        Me.txtUsrDe = New System.Windows.Forms.TextBox()
        Me.txtLicDe = New System.Windows.Forms.TextBox()
        Me.btnDecryptFile = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatus.Location = New System.Drawing.Point(232, 144)
        Me.txtStatus.MaxLength = 0
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(65, 22)
        Me.txtStatus.TabIndex = 10
        '
        'txtSecIDa
        '
        Me.txtSecIDa.AcceptsReturn = True
        Me.txtSecIDa.BackColor = System.Drawing.SystemColors.Window
        Me.txtSecIDa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSecIDa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSecIDa.Location = New System.Drawing.Point(96, 144)
        Me.txtSecIDa.MaxLength = 0
        Me.txtSecIDa.Name = "txtSecIDa"
        Me.txtSecIDa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSecIDa.Size = New System.Drawing.Size(73, 22)
        Me.txtSecIDa.TabIndex = 8
        '
        'cmdCalc
        '
        Me.cmdCalc.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCalc.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCalc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCalc.Location = New System.Drawing.Point(184, 199)
        Me.cmdCalc.Name = "cmdCalc"
        Me.cmdCalc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCalc.Size = New System.Drawing.Size(113, 33)
        Me.cmdCalc.TabIndex = 7
        Me.cmdCalc.Text = "&Calculate"
        Me.cmdCalc.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(678, 253)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(113, 33)
        Me.cmdExit.TabIndex = 6
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'txtSecId
        '
        Me.txtSecId.AcceptsReturn = True
        Me.txtSecId.BackColor = System.Drawing.SystemColors.Window
        Me.txtSecId.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSecId.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSecId.Location = New System.Drawing.Point(96, 104)
        Me.txtSecId.MaxLength = 0
        Me.txtSecId.Name = "txtSecId"
        Me.txtSecId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSecId.Size = New System.Drawing.Size(201, 22)
        Me.txtSecId.TabIndex = 4
        '
        'txtLicNo
        '
        Me.txtLicNo.AcceptsReturn = True
        Me.txtLicNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtLicNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLicNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLicNo.Location = New System.Drawing.Point(96, 64)
        Me.txtLicNo.MaxLength = 0
        Me.txtLicNo.Name = "txtLicNo"
        Me.txtLicNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLicNo.Size = New System.Drawing.Size(201, 22)
        Me.txtLicNo.TabIndex = 1
        '
        'txtUserName
        '
        Me.txtUserName.AcceptsReturn = True
        Me.txtUserName.BackColor = System.Drawing.SystemColors.Window
        Me.txtUserName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUserName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUserName.Location = New System.Drawing.Point(96, 24)
        Me.txtUserName.MaxLength = 24
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUserName.Size = New System.Drawing.Size(201, 22)
        Me.txtUserName.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(176, 144)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(49, 17)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Status:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(8, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(81, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Secutiry Tag:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(24, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(65, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Security ID:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Lic No.:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "User Name:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnEncrypt
        '
        Me.btnEncrypt.Location = New System.Drawing.Point(367, 104)
        Me.btnEncrypt.Name = "btnEncrypt"
        Me.btnEncrypt.Size = New System.Drawing.Size(110, 23)
        Me.btnEncrypt.TabIndex = 12
        Me.btnEncrypt.Text = "Encrypt"
        Me.btnEncrypt.UseVisualStyleBackColor = True
        '
        'txtUserE
        '
        Me.txtUserE.Location = New System.Drawing.Point(312, 24)
        Me.txtUserE.Name = "txtUserE"
        Me.txtUserE.Size = New System.Drawing.Size(242, 22)
        Me.txtUserE.TabIndex = 13
        '
        'txtLicE
        '
        Me.txtLicE.Location = New System.Drawing.Point(312, 64)
        Me.txtLicE.Name = "txtLicE"
        Me.txtLicE.Size = New System.Drawing.Size(242, 22)
        Me.txtLicE.TabIndex = 14
        '
        'btnDecrypt
        '
        Me.btnDecrypt.Location = New System.Drawing.Point(641, 104)
        Me.btnDecrypt.Name = "btnDecrypt"
        Me.btnDecrypt.Size = New System.Drawing.Size(75, 31)
        Me.btnDecrypt.TabIndex = 15
        Me.btnDecrypt.Text = "Decrypt"
        Me.btnDecrypt.UseVisualStyleBackColor = True
        '
        'txtUsrD
        '
        Me.txtUsrD.Location = New System.Drawing.Point(572, 24)
        Me.txtUsrD.Name = "txtUsrD"
        Me.txtUsrD.Size = New System.Drawing.Size(202, 22)
        Me.txtUsrD.TabIndex = 16
        '
        'txtLicD
        '
        Me.txtLicD.Location = New System.Drawing.Point(572, 64)
        Me.txtLicD.Name = "txtLicD"
        Me.txtLicD.Size = New System.Drawing.Size(202, 22)
        Me.txtLicD.TabIndex = 17
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(27, 199)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(105, 33)
        Me.btnClear.TabIndex = 18
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnWrtRpdat
        '
        Me.btnWrtRpdat.Location = New System.Drawing.Point(367, 138)
        Me.btnWrtRpdat.Name = "btnWrtRpdat"
        Me.btnWrtRpdat.Size = New System.Drawing.Size(105, 35)
        Me.btnWrtRpdat.TabIndex = 19
        Me.btnWrtRpdat.Text = "Write RP.Dat"
        Me.btnWrtRpdat.UseVisualStyleBackColor = True
        '
        'btnReadRPDat
        '
        Me.btnReadRPDat.Location = New System.Drawing.Point(623, 141)
        Me.btnReadRPDat.Name = "btnReadRPDat"
        Me.btnReadRPDat.Size = New System.Drawing.Size(113, 35)
        Me.btnReadRPDat.TabIndex = 20
        Me.btnReadRPDat.Text = "Read RP.dat"
        Me.btnReadRPDat.UseVisualStyleBackColor = True
        '
        'txtLicE1
        '
        Me.txtLicE1.Location = New System.Drawing.Point(572, 179)
        Me.txtLicE1.Name = "txtLicE1"
        Me.txtLicE1.Size = New System.Drawing.Size(207, 22)
        Me.txtLicE1.TabIndex = 21
        '
        'txtUsrE1
        '
        Me.txtUsrE1.Location = New System.Drawing.Point(312, 179)
        Me.txtUsrE1.Name = "txtUsrE1"
        Me.txtUsrE1.Size = New System.Drawing.Size(242, 22)
        Me.txtUsrE1.TabIndex = 22
        '
        'txtUsrDe
        '
        Me.txtUsrDe.Location = New System.Drawing.Point(312, 210)
        Me.txtUsrDe.Name = "txtUsrDe"
        Me.txtUsrDe.Size = New System.Drawing.Size(242, 22)
        Me.txtUsrDe.TabIndex = 23
        '
        'txtLicDe
        '
        Me.txtLicDe.Location = New System.Drawing.Point(572, 210)
        Me.txtLicDe.Name = "txtLicDe"
        Me.txtLicDe.Size = New System.Drawing.Size(207, 22)
        Me.txtLicDe.TabIndex = 24
        '
        'btnDecryptFile
        '
        Me.btnDecryptFile.Location = New System.Drawing.Point(504, 253)
        Me.btnDecryptFile.Name = "btnDecryptFile"
        Me.btnDecryptFile.Size = New System.Drawing.Size(114, 35)
        Me.btnDecryptFile.TabIndex = 25
        Me.btnDecryptFile.Text = "Decrypt File Read"
        Me.btnDecryptFile.UseVisualStyleBackColor = True
        '
        'frmSecurity
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(803, 300)
        Me.Controls.Add(Me.btnDecryptFile)
        Me.Controls.Add(Me.txtLicDe)
        Me.Controls.Add(Me.txtUsrDe)
        Me.Controls.Add(Me.txtUsrE1)
        Me.Controls.Add(Me.txtLicE1)
        Me.Controls.Add(Me.btnReadRPDat)
        Me.Controls.Add(Me.btnWrtRpdat)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.txtLicD)
        Me.Controls.Add(Me.txtUsrD)
        Me.Controls.Add(Me.btnDecrypt)
        Me.Controls.Add(Me.txtLicE)
        Me.Controls.Add(Me.txtUserE)
        Me.Controls.Add(Me.btnEncrypt)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.txtSecIDa)
        Me.Controls.Add(Me.cmdCalc)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.txtSecId)
        Me.Controls.Add(Me.txtLicNo)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmSecurity"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Security"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnEncrypt As System.Windows.Forms.Button
    Friend WithEvents txtUserE As System.Windows.Forms.TextBox
    Friend WithEvents txtLicE As System.Windows.Forms.TextBox
    Friend WithEvents btnDecrypt As System.Windows.Forms.Button
    Friend WithEvents txtUsrD As System.Windows.Forms.TextBox
    Friend WithEvents txtLicD As System.Windows.Forms.TextBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnWrtRpdat As System.Windows.Forms.Button
    Friend WithEvents btnReadRPDat As System.Windows.Forms.Button
    Friend WithEvents txtLicE1 As System.Windows.Forms.TextBox
    Friend WithEvents txtUsrE1 As System.Windows.Forms.TextBox
    Friend WithEvents txtUsrDe As System.Windows.Forms.TextBox
    Friend WithEvents txtLicDe As System.Windows.Forms.TextBox
    Friend WithEvents btnDecryptFile As System.Windows.Forms.Button
#End Region 
End Class