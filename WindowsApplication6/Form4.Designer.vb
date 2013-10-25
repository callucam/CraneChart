<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
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
        Me.Label15 = New System.Windows.Forms.Label()
        Me.DataGridViewBoomLengthWeight = New System.Windows.Forms.DataGridView()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.DataGridViewWAC = New System.Windows.Forms.DataGridView()
        Me.SaveWACButton = New System.Windows.Forms.Button()
        Me.SaveBoomWeightButton = New System.Windows.Forms.Button()
        Me.PasteWACButton = New System.Windows.Forms.Button()
        Me.PasteBoomWeightButton = New System.Windows.Forms.Button()
        CType(Me.DataGridViewBoomLengthWeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridViewWAC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(9, 278)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(71, 13)
        Me.Label15.TabIndex = 69
        Me.Label15.Text = "Boom Weight"
        '
        'DataGridViewBoomLengthWeight
        '
        Me.DataGridViewBoomLengthWeight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewBoomLengthWeight.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewBoomLengthWeight.Location = New System.Drawing.Point(12, 294)
        Me.DataGridViewBoomLengthWeight.Name = "DataGridViewBoomLengthWeight"
        Me.DataGridViewBoomLengthWeight.Size = New System.Drawing.Size(560, 256)
        Me.DataGridViewBoomLengthWeight.TabIndex = 68
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(12, 9)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(106, 13)
        Me.Label13.TabIndex = 67
        Me.Label13.Text = "Weights and Centres"
        '
        'DataGridViewWAC
        '
        Me.DataGridViewWAC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewWAC.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewWAC.Location = New System.Drawing.Point(12, 25)
        Me.DataGridViewWAC.Name = "DataGridViewWAC"
        Me.DataGridViewWAC.Size = New System.Drawing.Size(560, 250)
        Me.DataGridViewWAC.TabIndex = 66
        '
        'SaveWACButton
        '
        Me.SaveWACButton.Location = New System.Drawing.Point(578, 252)
        Me.SaveWACButton.Name = "SaveWACButton"
        Me.SaveWACButton.Size = New System.Drawing.Size(75, 23)
        Me.SaveWACButton.TabIndex = 70
        Me.SaveWACButton.Text = "Save"
        Me.SaveWACButton.UseVisualStyleBackColor = True
        '
        'SaveBoomWeightButton
        '
        Me.SaveBoomWeightButton.Location = New System.Drawing.Point(578, 527)
        Me.SaveBoomWeightButton.Name = "SaveBoomWeightButton"
        Me.SaveBoomWeightButton.Size = New System.Drawing.Size(75, 23)
        Me.SaveBoomWeightButton.TabIndex = 71
        Me.SaveBoomWeightButton.Text = "Save"
        Me.SaveBoomWeightButton.UseVisualStyleBackColor = True
        '
        'PasteWACButton
        '
        Me.PasteWACButton.Location = New System.Drawing.Point(578, 223)
        Me.PasteWACButton.Name = "PasteWACButton"
        Me.PasteWACButton.Size = New System.Drawing.Size(75, 23)
        Me.PasteWACButton.TabIndex = 72
        Me.PasteWACButton.Text = "Paste"
        Me.PasteWACButton.UseVisualStyleBackColor = True
        '
        'PasteBoomWeightButton
        '
        Me.PasteBoomWeightButton.Location = New System.Drawing.Point(578, 498)
        Me.PasteBoomWeightButton.Name = "PasteBoomWeightButton"
        Me.PasteBoomWeightButton.Size = New System.Drawing.Size(75, 23)
        Me.PasteBoomWeightButton.TabIndex = 73
        Me.PasteBoomWeightButton.Text = "Paste"
        Me.PasteBoomWeightButton.UseVisualStyleBackColor = True
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(662, 562)
        Me.Controls.Add(Me.PasteBoomWeightButton)
        Me.Controls.Add(Me.PasteWACButton)
        Me.Controls.Add(Me.SaveBoomWeightButton)
        Me.Controls.Add(Me.SaveWACButton)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.DataGridViewBoomLengthWeight)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.DataGridViewWAC)
        Me.Name = "Form4"
        Me.Text = "Crane Weights and Centres"
        CType(Me.DataGridViewBoomLengthWeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridViewWAC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewBoomLengthWeight As System.Windows.Forms.DataGridView
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewWAC As System.Windows.Forms.DataGridView
    Friend WithEvents SaveWACButton As System.Windows.Forms.Button
    Friend WithEvents SaveBoomWeightButton As System.Windows.Forms.Button
    Friend WithEvents PasteWACButton As System.Windows.Forms.Button
    Friend WithEvents PasteBoomWeightButton As System.Windows.Forms.Button
End Class
