<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
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
        Me.DataGridViewListChart = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PasteChart = New System.Windows.Forms.Button()
        CType(Me.DataGridViewListChart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewListChart
        '
        Me.DataGridViewListChart.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewListChart.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridViewListChart.Location = New System.Drawing.Point(12, 36)
        Me.DataGridViewListChart.Name = "DataGridViewListChart"
        Me.DataGridViewListChart.Size = New System.Drawing.Size(560, 470)
        Me.DataGridViewListChart.TabIndex = 10
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(467, 527)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(105, 23)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "Save Changes"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'PasteChart
        '
        Me.PasteChart.Location = New System.Drawing.Point(386, 527)
        Me.PasteChart.Name = "PasteChart"
        Me.PasteChart.Size = New System.Drawing.Size(75, 23)
        Me.PasteChart.TabIndex = 12
        Me.PasteChart.Text = "Paste"
        Me.PasteChart.UseVisualStyleBackColor = True
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 562)
        Me.Controls.Add(Me.PasteChart)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridViewListChart)
        Me.Name = "Form5"
        Me.Text = "Land-Based Lift Capacities"
        CType(Me.DataGridViewListChart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridViewListChart As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents PasteChart As System.Windows.Forms.Button
End Class
