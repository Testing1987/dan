<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Viewport_to_poly_form
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
        Me.Button_pick = New System.Windows.Forms.Button()
        Me.Button_draw = New System.Windows.Forms.Button()
        Me.CheckBox_close = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Button_pick
        '
        Me.Button_pick.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_pick.Location = New System.Drawing.Point(12, 12)
        Me.Button_pick.Name = "Button_pick"
        Me.Button_pick.Size = New System.Drawing.Size(86, 36)
        Me.Button_pick.TabIndex = 0
        Me.Button_pick.Text = "Pick"
        Me.Button_pick.UseVisualStyleBackColor = True
        '
        'Button_draw
        '
        Me.Button_draw.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_draw.Location = New System.Drawing.Point(12, 54)
        Me.Button_draw.Name = "Button_draw"
        Me.Button_draw.Size = New System.Drawing.Size(86, 36)
        Me.Button_draw.TabIndex = 0
        Me.Button_draw.Text = "Draw"
        Me.Button_draw.UseVisualStyleBackColor = True
        '
        'CheckBox_close
        '
        Me.CheckBox_close.AutoSize = True
        Me.CheckBox_close.Location = New System.Drawing.Point(12, 96)
        Me.CheckBox_close.Name = "CheckBox_close"
        Me.CheckBox_close.Size = New System.Drawing.Size(198, 17)
        Me.CheckBox_close.TabIndex = 1
        Me.CheckBox_close.Text = "Close current DWG and do not save"
        Me.CheckBox_close.UseVisualStyleBackColor = True
        '
        'Viewport_to_poly_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(240, 108)
        Me.Controls.Add(Me.CheckBox_close)
        Me.Controls.Add(Me.Button_draw)
        Me.Controls.Add(Me.Button_pick)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Viewport_to_poly_form"
        Me.Text = "Viewport to poly"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_pick As System.Windows.Forms.Button
    Friend WithEvents Button_draw As System.Windows.Forms.Button
    Friend WithEvents CheckBox_close As System.Windows.Forms.CheckBox
End Class
