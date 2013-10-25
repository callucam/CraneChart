<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Me.CopyH = New System.Windows.Forms.Button()
        Me.SaveH = New System.Windows.Forms.Button()
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
        'CopyH
        '
        Me.CopyH.Location = New System.Drawing.Point(12, 527)
        Me.CopyH.Name = "CopyH"
        Me.CopyH.Size = New System.Drawing.Size(139, 23)
        Me.CopyH.TabIndex = 1
        Me.CopyH.Text = "Paste from Clipboard"
        Me.CopyH.UseVisualStyleBackColor = True
        '
        'SaveH
        '
        Me.SaveH.Location = New System.Drawing.Point(157, 527)
        Me.SaveH.Name = "SaveH"
        Me.SaveH.Size = New System.Drawing.Size(139, 23)
        Me.SaveH.TabIndex = 2
        Me.SaveH.Text = "Save Hydrostatics"
        Me.SaveH.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.SaveH)
        Me.Controls.Add(Me.CopyH)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Form2"
        Me.Text = "Hydrostatics"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents CopyH As System.Windows.Forms.Button
    Friend WithEvents SaveH As System.Windows.Forms.Button
End Class
