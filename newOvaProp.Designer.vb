<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class newOVActl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtIP = New System.Windows.Forms.TextBox()
        Me.cboOva = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDNS = New System.Windows.Forms.TextBox()
        Me.txtGW = New System.Windows.Forms.TextBox()
        Me.txtMask = New System.Windows.Forms.TextBox()
        Me.txtProxy = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtNTP = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(345, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 21)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "OVA ID:"
        '
        'txtIP
        '
        Me.txtIP.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtIP.Location = New System.Drawing.Point(104, 55)
        Me.txtIP.Multiline = True
        Me.txtIP.Name = "txtIP"
        Me.txtIP.Size = New System.Drawing.Size(205, 27)
        Me.txtIP.TabIndex = 1
        Me.txtIP.Text = "192.168.1.200"
        '
        'cboOva
        '
        Me.cboOva.FormattingEnabled = True
        Me.cboOva.Location = New System.Drawing.Point(418, 55)
        Me.cboOva.Name = "cboOva"
        Me.cboOva.Size = New System.Drawing.Size(87, 21)
        Me.cboOva.Sorted = True
        Me.cboOva.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(12, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 21)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "IP Address:"
        '
        'txtDNS
        '
        Me.txtDNS.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDNS.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtDNS.Location = New System.Drawing.Point(104, 88)
        Me.txtDNS.Multiline = True
        Me.txtDNS.Name = "txtDNS"
        Me.txtDNS.Size = New System.Drawing.Size(205, 27)
        Me.txtDNS.TabIndex = 2
        Me.txtDNS.Text = "8.8.8.8,8.8.8.4"
        '
        'txtGW
        '
        Me.txtGW.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGW.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtGW.Location = New System.Drawing.Point(104, 121)
        Me.txtGW.Multiline = True
        Me.txtGW.Name = "txtGW"
        Me.txtGW.Size = New System.Drawing.Size(205, 27)
        Me.txtGW.TabIndex = 3
        Me.txtGW.Text = "192.168.1.1"
        '
        'txtMask
        '
        Me.txtMask.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMask.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtMask.Location = New System.Drawing.Point(104, 154)
        Me.txtMask.Multiline = True
        Me.txtMask.Name = "txtMask"
        Me.txtMask.Size = New System.Drawing.Size(205, 27)
        Me.txtMask.TabIndex = 4
        Me.txtMask.Text = "255.255.0.0"
        '
        'txtProxy
        '
        Me.txtProxy.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProxy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtProxy.Location = New System.Drawing.Point(104, 220)
        Me.txtProxy.Multiline = True
        Me.txtProxy.Name = "txtProxy"
        Me.txtProxy.Size = New System.Drawing.Size(205, 27)
        Me.txtProxy.TabIndex = 6
        Me.txtProxy.Text = "https://my.proxy.com"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(-3, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(101, 21)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "DNS Servers:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(25, 121)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 21)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Gateway:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(23, 154)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 21)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Netmask:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(46, 220)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 21)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Proxy:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label7.Location = New System.Drawing.Point(0, 187)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 21)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "NTP Servers:"
        '
        'txtNTP
        '
        Me.txtNTP.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNTP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtNTP.Location = New System.Drawing.Point(104, 187)
        Me.txtNTP.Multiline = True
        Me.txtNTP.Name = "txtNTP"
        Me.txtNTP.Size = New System.Drawing.Size(205, 27)
        Me.txtNTP.TabIndex = 5
        Me.txtNTP.Text = "1.2.3.4,my.ntp.com"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label8.Location = New System.Drawing.Point(12, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(483, 17)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "ENTER OVA SETUP INFORMATION (leave fields blank for DHCP self-configuration)"
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Navy
        Me.btnCancel.Location = New System.Drawing.Point(349, 222)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(78, 25)
        Me.btnCancel.TabIndex = 24
        Me.btnCancel.Text = "CANCEL"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnGo
        '
        Me.btnGo.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGo.ForeColor = System.Drawing.Color.Navy
        Me.btnGo.Location = New System.Drawing.Point(427, 222)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(78, 25)
        Me.btnGo.TabIndex = 25
        Me.btnGo.Text = "GET OVA"
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'newOVActl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.btnGo)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtNTP)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtProxy)
        Me.Controls.Add(Me.txtMask)
        Me.Controls.Add(Me.txtGW)
        Me.Controls.Add(Me.txtDNS)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboOva)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtIP)
        Me.Name = "newOVActl"
        Me.Size = New System.Drawing.Size(509, 263)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents txtIP As TextBox
    Friend WithEvents cboOva As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtDNS As TextBox
    Friend WithEvents txtGW As TextBox
    Friend WithEvents txtMask As TextBox
    Friend WithEvents txtProxy As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents txtNTP As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnGo As Button
End Class
