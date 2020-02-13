<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptDevice
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRptDevice))
        Me.cbLB = New System.Windows.Forms.CheckedListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnUP = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cbLB
        '
        Me.cbLB.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbLB.FormattingEnabled = True
        Me.cbLB.Location = New System.Drawing.Point(73, 34)
        Me.cbLB.Name = "cbLB"
        Me.cbLB.Size = New System.Drawing.Size(195, 244)
        Me.cbLB.TabIndex = 0
        Me.cbLB.ThreeDCheckBoxes = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Navy
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(99, 303)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(93, 28)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "BUILD"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'btnUP
        '
        Me.btnUP.BackColor = System.Drawing.Color.White
        Me.btnUP.FlatAppearance.BorderColor = System.Drawing.Color.Navy
        Me.btnUP.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUP.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUP.ForeColor = System.Drawing.Color.White
        Me.btnUP.Image = CType(resources.GetObject("btnUP.Image"), System.Drawing.Image)
        Me.btnUP.Location = New System.Drawing.Point(263, 51)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(36, 39)
        Me.btnUP.TabIndex = 17
        Me.btnUP.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.White
        Me.Button2.FlatAppearance.BorderColor = System.Drawing.Color.Navy
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(263, 96)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(36, 39)
        Me.Button2.TabIndex = 18
        Me.Button2.UseVisualStyleBackColor = False
        '
        'frmRptDevice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(337, 355)
        Me.ControlBox = False
        Me.Controls.Add(Me.cbLB)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnUP)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmRptDevice"
        Me.Text = "Report Format"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cbLB As CheckedListBox
    Friend WithEvents Button1 As Button
    Friend WithEvents btnUP As Button
    Friend WithEvents Button2 As Button
End Class
