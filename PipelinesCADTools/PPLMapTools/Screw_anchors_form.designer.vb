<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Screw_anchors_form
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
        Me.TextBox_screw_anchors_spacing = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel_screw_anchors = New System.Windows.Forms.Panel()
        Me.Button_clear_text = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ListBox_Recalculated_chainages = New System.Windows.Forms.ListBox()
        Me.ListBox_Picked_chainages = New System.Windows.Forms.ListBox()
        Me.Button_insert_block_recalculated_chainages = New System.Windows.Forms.Button()
        Me.ListBox_no_of_screw_anchors = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel_screw_anchors.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_pick
        '
        Me.Button_pick.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_pick.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick.ForeColor = System.Drawing.Color.Black
        Me.Button_pick.Location = New System.Drawing.Point(17, 113)
        Me.Button_pick.Name = "Button_pick"
        Me.Button_pick.Size = New System.Drawing.Size(158, 39)
        Me.Button_pick.TabIndex = 100
        Me.Button_pick.Text = "Pick Points"
        Me.Button_pick.UseVisualStyleBackColor = False
        '
        'TextBox_screw_anchors_spacing
        '
        Me.TextBox_screw_anchors_spacing.BackColor = System.Drawing.Color.White
        Me.TextBox_screw_anchors_spacing.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_screw_anchors_spacing.ForeColor = System.Drawing.Color.Black
        Me.TextBox_screw_anchors_spacing.Location = New System.Drawing.Point(68, 26)
        Me.TextBox_screw_anchors_spacing.Multiline = True
        Me.TextBox_screw_anchors_spacing.Name = "TextBox_screw_anchors_spacing"
        Me.TextBox_screw_anchors_spacing.Size = New System.Drawing.Size(53, 26)
        Me.TextBox_screw_anchors_spacing.TabIndex = 4
        Me.TextBox_screw_anchors_spacing.Text = "23.8"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(10, 7)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(171, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Screw Anchors spacing"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Panel_screw_anchors
        '
        Me.Panel_screw_anchors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_screw_anchors.Controls.Add(Me.TextBox_screw_anchors_spacing)
        Me.Panel_screw_anchors.Controls.Add(Me.Label5)
        Me.Panel_screw_anchors.Location = New System.Drawing.Point(2, 25)
        Me.Panel_screw_anchors.Name = "Panel_screw_anchors"
        Me.Panel_screw_anchors.Size = New System.Drawing.Size(206, 66)
        Me.Panel_screw_anchors.TabIndex = 9
        '
        'Button_clear_text
        '
        Me.Button_clear_text.BackColor = System.Drawing.Color.OrangeRed
        Me.Button_clear_text.ForeColor = System.Drawing.Color.White
        Me.Button_clear_text.Location = New System.Drawing.Point(439, 4)
        Me.Button_clear_text.Name = "Button_clear_text"
        Me.Button_clear_text.Size = New System.Drawing.Size(75, 23)
        Me.Button_clear_text.TabIndex = 114
        Me.Button_clear_text.Text = "Clear"
        Me.Button_clear_text.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(520, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(76, 14)
        Me.Label6.TabIndex = 112
        Me.Label6.Text = "Recalculated"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(317, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(44, 14)
        Me.Label7.TabIndex = 113
        Me.Label7.Text = "Picked"
        '
        'ListBox_Recalculated_chainages
        '
        Me.ListBox_Recalculated_chainages.BackColor = System.Drawing.Color.White
        Me.ListBox_Recalculated_chainages.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_Recalculated_chainages.ForeColor = System.Drawing.Color.Black
        Me.ListBox_Recalculated_chainages.FormattingEnabled = True
        Me.ListBox_Recalculated_chainages.HorizontalScrollbar = True
        Me.ListBox_Recalculated_chainages.ItemHeight = 14
        Me.ListBox_Recalculated_chainages.Location = New System.Drawing.Point(520, 25)
        Me.ListBox_Recalculated_chainages.Name = "ListBox_Recalculated_chainages"
        Me.ListBox_Recalculated_chainages.Size = New System.Drawing.Size(194, 396)
        Me.ListBox_Recalculated_chainages.TabIndex = 110
        '
        'ListBox_Picked_chainages
        '
        Me.ListBox_Picked_chainages.BackColor = System.Drawing.Color.White
        Me.ListBox_Picked_chainages.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_Picked_chainages.ForeColor = System.Drawing.Color.Black
        Me.ListBox_Picked_chainages.FormattingEnabled = True
        Me.ListBox_Picked_chainages.HorizontalScrollbar = True
        Me.ListBox_Picked_chainages.ItemHeight = 14
        Me.ListBox_Picked_chainages.Location = New System.Drawing.Point(320, 25)
        Me.ListBox_Picked_chainages.Name = "ListBox_Picked_chainages"
        Me.ListBox_Picked_chainages.Size = New System.Drawing.Size(194, 396)
        Me.ListBox_Picked_chainages.TabIndex = 111
        '
        'Button_insert_block_recalculated_chainages
        '
        Me.Button_insert_block_recalculated_chainages.BackColor = System.Drawing.Color.Lime
        Me.Button_insert_block_recalculated_chainages.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_insert_block_recalculated_chainages.Location = New System.Drawing.Point(14, 352)
        Me.Button_insert_block_recalculated_chainages.Name = "Button_insert_block_recalculated_chainages"
        Me.Button_insert_block_recalculated_chainages.Size = New System.Drawing.Size(194, 52)
        Me.Button_insert_block_recalculated_chainages.TabIndex = 116
        Me.Button_insert_block_recalculated_chainages.Text = "Insert block" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Screw Anchor"
        Me.Button_insert_block_recalculated_chainages.UseVisualStyleBackColor = False
        '
        'ListBox_no_of_screw_anchors
        '
        Me.ListBox_no_of_screw_anchors.BackColor = System.Drawing.Color.White
        Me.ListBox_no_of_screw_anchors.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_no_of_screw_anchors.ForeColor = System.Drawing.Color.Black
        Me.ListBox_no_of_screw_anchors.FormattingEnabled = True
        Me.ListBox_no_of_screw_anchors.HorizontalScrollbar = True
        Me.ListBox_no_of_screw_anchors.ItemHeight = 14
        Me.ListBox_no_of_screw_anchors.Location = New System.Drawing.Point(214, 25)
        Me.ListBox_no_of_screw_anchors.Name = "ListBox_no_of_screw_anchors"
        Me.ListBox_no_of_screw_anchors.Size = New System.Drawing.Size(100, 396)
        Me.ListBox_no_of_screw_anchors.TabIndex = 111
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(226, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 14)
        Me.Label1.TabIndex = 113
        Me.Label1.Text = "No of Anchors"
        '
        'Screw_anchors_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(722, 430)
        Me.Controls.Add(Me.Button_insert_block_recalculated_chainages)
        Me.Controls.Add(Me.Button_clear_text)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.ListBox_Recalculated_chainages)
        Me.Controls.Add(Me.ListBox_no_of_screw_anchors)
        Me.Controls.Add(Me.ListBox_Picked_chainages)
        Me.Controls.Add(Me.Panel_screw_anchors)
        Me.Controls.Add(Me.Button_pick)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Screw_anchors_form"
        Me.Text = "Chainage for screw anchors"
        Me.Panel_screw_anchors.ResumeLayout(False)
        Me.Panel_screw_anchors.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_pick As System.Windows.Forms.Button
    Friend WithEvents TextBox_screw_anchors_spacing As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel_screw_anchors As System.Windows.Forms.Panel
    Friend WithEvents Button_clear_text As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ListBox_Recalculated_chainages As System.Windows.Forms.ListBox
    Friend WithEvents ListBox_Picked_chainages As System.Windows.Forms.ListBox
    Friend WithEvents Button_insert_block_recalculated_chainages As System.Windows.Forms.Button
    Friend WithEvents ListBox_no_of_screw_anchors As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
