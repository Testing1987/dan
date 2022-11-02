<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Easement_builder_Form
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
        Me.TextBox_left = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_right = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button_offset = New System.Windows.Forms.Button()
        Me.Button_draw_polyline = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox_left
        '
        Me.TextBox_left.Location = New System.Drawing.Point(17, 29)
        Me.TextBox_left.Name = "TextBox_left"
        Me.TextBox_left.Size = New System.Drawing.Size(43, 21)
        Me.TextBox_left.TabIndex = 0
        Me.TextBox_left.Text = "50"
        Me.TextBox_left.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Left"
        '
        'TextBox_right
        '
        Me.TextBox_right.Location = New System.Drawing.Point(83, 29)
        Me.TextBox_right.Name = "TextBox_right"
        Me.TextBox_right.Size = New System.Drawing.Size(43, 21)
        Me.TextBox_right.TabIndex = 1
        Me.TextBox_right.Text = "30"
        Me.TextBox_right.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(80, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 15)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Right"
        '
        'Button_offset
        '
        Me.Button_offset.Location = New System.Drawing.Point(17, 66)
        Me.Button_offset.Name = "Button_offset"
        Me.Button_offset.Size = New System.Drawing.Size(109, 27)
        Me.Button_offset.TabIndex = 2
        Me.Button_offset.Text = "Offset"
        Me.Button_offset.UseVisualStyleBackColor = True
        '
        'Button_draw_polyline
        '
        Me.Button_draw_polyline.Location = New System.Drawing.Point(193, 5)
        Me.Button_draw_polyline.Name = "Button_draw_polyline"
        Me.Button_draw_polyline.Size = New System.Drawing.Size(126, 27)
        Me.Button_draw_polyline.TabIndex = 3
        Me.Button_draw_polyline.Text = "Draw easement"
        Me.Button_draw_polyline.UseVisualStyleBackColor = True
        '
        'Easement_builder_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(331, 133)
        Me.Controls.Add(Me.Button_draw_polyline)
        Me.Controls.Add(Me.Button_offset)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_right)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_left)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Easement_builder_Form"
        Me.Text = "Easement_builder_Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_left As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_right As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button_offset As System.Windows.Forms.Button
    Friend WithEvents Button_draw_polyline As System.Windows.Forms.Button
End Class
