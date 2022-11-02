<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Split_deflection_form
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
        Me.Button_split = New System.Windows.Forms.Button()
        Me.TextBox_distance = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_split_angle = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button_split
        '
        Me.Button_split.Location = New System.Drawing.Point(12, 75)
        Me.Button_split.Name = "Button_split"
        Me.Button_split.Size = New System.Drawing.Size(251, 32)
        Me.Button_split.TabIndex = 0
        Me.Button_split.Text = "Split"
        Me.Button_split.UseVisualStyleBackColor = True
        '
        'TextBox_distance
        '
        Me.TextBox_distance.BackColor = System.Drawing.Color.White
        Me.TextBox_distance.ForeColor = System.Drawing.Color.Black
        Me.TextBox_distance.Location = New System.Drawing.Point(12, 12)
        Me.TextBox_distance.Name = "TextBox_distance"
        Me.TextBox_distance.Size = New System.Drawing.Size(94, 21)
        Me.TextBox_distance.TabIndex = 1
        Me.TextBox_distance.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(112, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Joint Length"
        '
        'TextBox_split_angle
        '
        Me.TextBox_split_angle.BackColor = System.Drawing.Color.White
        Me.TextBox_split_angle.ForeColor = System.Drawing.Color.Black
        Me.TextBox_split_angle.Location = New System.Drawing.Point(12, 47)
        Me.TextBox_split_angle.Name = "TextBox_split_angle"
        Me.TextBox_split_angle.Size = New System.Drawing.Size(94, 21)
        Me.TextBox_split_angle.TabIndex = 1
        Me.TextBox_split_angle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(112, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(151, 30)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Maximum allowable bend" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[Decimal degrees]"
        '
        'Split_deflection_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(270, 115)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_split_angle)
        Me.Controls.Add(Me.TextBox_distance)
        Me.Controls.Add(Me.Button_split)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Split_deflection_form"
        Me.Text = "Split deflections"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_split As System.Windows.Forms.Button
    Friend WithEvents TextBox_distance As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_split_angle As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
