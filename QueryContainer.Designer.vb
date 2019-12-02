<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QueryContainer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(QueryContainer))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblDescript = New System.Windows.Forms.Label()
        Me.btnXls = New System.Windows.Forms.Button()
        Me.btnCsv = New System.Windows.Forms.Button()
        Me.btnJson = New System.Windows.Forms.Button()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TextBox1.Location = New System.Drawing.Point(3, 26)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(545, 63)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "in:devices"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(11, 2)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 21)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Query"
        '
        'lblDescript
        '
        Me.lblDescript.AutoSize = True
        Me.lblDescript.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescript.ForeColor = System.Drawing.Color.Navy
        Me.lblDescript.Location = New System.Drawing.Point(12, 103)
        Me.lblDescript.Name = "lblDescript"
        Me.lblDescript.Size = New System.Drawing.Size(0, 17)
        Me.lblDescript.TabIndex = 2
        Me.lblDescript.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnXls
        '
        Me.btnXls.AutoSize = True
        Me.btnXls.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnXls.Image = CType(resources.GetObject("btnXls.Image"), System.Drawing.Image)
        Me.btnXls.Location = New System.Drawing.Point(492, 157)
        Me.btnXls.Name = "btnXls"
        Me.btnXls.Size = New System.Drawing.Size(56, 56)
        Me.btnXls.TabIndex = 3
        Me.btnXls.UseVisualStyleBackColor = True
        Me.btnXls.Visible = False
        '
        'btnCsv
        '
        Me.btnCsv.AutoSize = True
        Me.btnCsv.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnCsv.Image = CType(resources.GetObject("btnCsv.Image"), System.Drawing.Image)
        Me.btnCsv.Location = New System.Drawing.Point(433, 157)
        Me.btnCsv.Name = "btnCsv"
        Me.btnCsv.Size = New System.Drawing.Size(56, 56)
        Me.btnCsv.TabIndex = 4
        Me.btnCsv.UseVisualStyleBackColor = True
        Me.btnCsv.Visible = False
        '
        'btnJson
        '
        Me.btnJson.AutoSize = True
        Me.btnJson.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnJson.Image = CType(resources.GetObject("btnJson.Image"), System.Drawing.Image)
        Me.btnJson.Location = New System.Drawing.Point(374, 157)
        Me.btnJson.Name = "btnJson"
        Me.btnJson.Size = New System.Drawing.Size(56, 56)
        Me.btnJson.TabIndex = 5
        Me.btnJson.UseVisualStyleBackColor = True
        '
        'btnGo
        '
        Me.btnGo.AutoSize = True
        Me.btnGo.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnGo.Image = CType(resources.GetObject("btnGo.Image"), System.Drawing.Image)
        Me.btnGo.Location = New System.Drawing.Point(492, 95)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(56, 56)
        Me.btnGo.TabIndex = 6
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'QueryContainer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.btnGo)
        Me.Controls.Add(Me.btnJson)
        Me.Controls.Add(Me.btnCsv)
        Me.Controls.Add(Me.btnXls)
        Me.Controls.Add(Me.lblDescript)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "QueryContainer"
        Me.Size = New System.Drawing.Size(551, 217)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblDescript As Label
    Friend WithEvents btnXls As Button
    Friend WithEvents btnCsv As Button
    Friend WithEvents btnJson As Button
    Friend WithEvents btnGo As Button
End Class
