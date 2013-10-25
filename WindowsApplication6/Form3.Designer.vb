<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.SaveC = New System.Windows.Forms.Button()
        Me.CopyC = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(700, 500)
        Me.DataGridView1.TabIndex = 0
        '
        'SaveC
        '
        Me.SaveC.Location = New System.Drawing.Point(157, 527)
        Me.SaveC.Name = "SaveC"
        Me.SaveC.Size = New System.Drawing.Size(139, 23)
        Me.SaveC.TabIndex = 4
        Me.SaveC.Text = "Save Cross Curves"
        Me.SaveC.UseVisualStyleBackColor = True
        '
        'CopyC
        '
        Me.CopyC.Location = New System.Drawing.Point(12, 527)
        Me.CopyC.Name = "CopyC"
        Me.CopyC.Size = New System.Drawing.Size(139, 23)
        Me.CopyC.TabIndex = 3
        Me.CopyC.Text = "Copy from Clipboard"
        Me.CopyC.UseVisualStyleBackColor = True
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.SaveC)
        Me.Controls.Add(Me.CopyC)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Form3"
        Me.Text = "Cross Curves"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents SaveC As System.Windows.Forms.Button
    Friend WithEvents CopyC As System.Windows.Forms.Button
End Class
