﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainUI
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainUI))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.logTB = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.xlsEngine = New System.ComponentModel.BackgroundWorker()
        Me.getOVA = New System.ComponentModel.BackgroundWorker()
        Me.btnOVA = New System.Windows.Forms.Button()
        Me.NewOVActl1 = New AQL2XLS_WinApp.newOVActl()
        Me.QueryContainer2 = New AQL2XLS_WinApp.QueryContainer()
        Me.QueryContainer1 = New AQL2XLS_WinApp.QueryContainer()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(1, 7)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(189, 98)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.CheckBox1.Location = New System.Drawing.Point(46, 413)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(114, 25)
        Me.CheckBox1.TabIndex = 3
        Me.CheckBox1.Text = "Multi-Query"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'logTB
        '
        Me.logTB.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.logTB.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.logTB.Location = New System.Drawing.Point(1, 111)
        Me.logTB.Multiline = True
        Me.logTB.Name = "logTB"
        Me.logTB.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.logTB.Size = New System.Drawing.Size(189, 292)
        Me.logTB.TabIndex = 9
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'xlsEngine
        '
        '
        'getOVA
        '
        '
        'btnOVA
        '
        Me.btnOVA.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOVA.ForeColor = System.Drawing.Color.Navy
        Me.btnOVA.Location = New System.Drawing.Point(59, 414)
        Me.btnOVA.Name = "btnOVA"
        Me.btnOVA.Size = New System.Drawing.Size(78, 25)
        Me.btnOVA.TabIndex = 10
        Me.btnOVA.Text = "Get OVA"
        Me.btnOVA.UseVisualStyleBackColor = True
        '
        'NewOVActl1
        '
        Me.NewOVActl1.BackColor = System.Drawing.Color.White
        Me.NewOVActl1.Location = New System.Drawing.Point(206, 7)
        Me.NewOVActl1.Name = "NewOVActl1"
        Me.NewOVActl1.Size = New System.Drawing.Size(528, 265)
        Me.NewOVActl1.TabIndex = 11
        Me.NewOVActl1.Visible = False
        '
        'QueryContainer2
        '
        Me.QueryContainer2.BackColor = System.Drawing.Color.White
        Me.QueryContainer2.Location = New System.Drawing.Point(196, 222)
        Me.QueryContainer2.Name = "QueryContainer2"
        Me.QueryContainer2.Size = New System.Drawing.Size(551, 215)
        Me.QueryContainer2.TabIndex = 2
        '
        'QueryContainer1
        '
        Me.QueryContainer1.BackColor = System.Drawing.Color.White
        Me.QueryContainer1.Location = New System.Drawing.Point(196, 7)
        Me.QueryContainer1.Name = "QueryContainer1"
        Me.QueryContainer1.Size = New System.Drawing.Size(551, 219)
        Me.QueryContainer1.TabIndex = 1
        '
        'MainUI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(759, 441)
        Me.Controls.Add(Me.NewOVActl1)
        Me.Controls.Add(Me.btnOVA)
        Me.Controls.Add(Me.logTB)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.QueryContainer2)
        Me.Controls.Add(Me.QueryContainer1)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "MainUI"
        Me.Text = "Armis AQL2XLS"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents QueryContainer1 As QueryContainer
    Friend WithEvents QueryContainer2 As QueryContainer
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents logTB As TextBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents xlsEngine As System.ComponentModel.BackgroundWorker
    Friend WithEvents getOVA As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnOVA As Button
    Friend WithEvents NewOVActl1 As newOVActl
End Class
