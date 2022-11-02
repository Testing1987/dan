<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Chainage_operations_form
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
        Me.TextBox_chainage_middle = New System.Windows.Forms.TextBox()
        Me.RadioButton_keep_low = New System.Windows.Forms.RadioButton()
        Me.RadioButton_keep_middle = New System.Windows.Forms.RadioButton()
        Me.RadioButton_keep_high = New System.Windows.Forms.RadioButton()
        Me.TextBox_chainage_high = New System.Windows.Forms.TextBox()
        Me.TextBox_chainage_low = New System.Windows.Forms.TextBox()
        Me.TextBox_amount = New System.Windows.Forms.TextBox()
        Me.Button_plus = New System.Windows.Forms.Button()
        Me.Button_minus = New System.Windows.Forms.Button()
        Me.Button_pick_chainage = New System.Windows.Forms.Button()
        Me.Button_pick_amount = New System.Windows.Forms.Button()
        Me.Button_split = New System.Windows.Forms.Button()
        Me.CheckBox_low = New System.Windows.Forms.CheckBox()
        Me.CheckBox_high = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'TextBox_chainage_middle
        '
        Me.TextBox_chainage_middle.Location = New System.Drawing.Point(65, 34)
        Me.TextBox_chainage_middle.Name = "TextBox_chainage_middle"
        Me.TextBox_chainage_middle.Size = New System.Drawing.Size(114, 21)
        Me.TextBox_chainage_middle.TabIndex = 0
        '
        'RadioButton_keep_low
        '
        Me.RadioButton_keep_low.AutoSize = True
        Me.RadioButton_keep_low.Location = New System.Drawing.Point(6, 8)
        Me.RadioButton_keep_low.Name = "RadioButton_keep_low"
        Me.RadioButton_keep_low.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_keep_low.Size = New System.Drawing.Size(53, 19)
        Me.RadioButton_keep_low.TabIndex = 1
        Me.RadioButton_keep_low.Text = "Hold "
        Me.RadioButton_keep_low.UseVisualStyleBackColor = True
        '
        'RadioButton_keep_middle
        '
        Me.RadioButton_keep_middle.AutoSize = True
        Me.RadioButton_keep_middle.Checked = True
        Me.RadioButton_keep_middle.Location = New System.Drawing.Point(9, 34)
        Me.RadioButton_keep_middle.Name = "RadioButton_keep_middle"
        Me.RadioButton_keep_middle.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_keep_middle.Size = New System.Drawing.Size(50, 19)
        Me.RadioButton_keep_middle.TabIndex = 1
        Me.RadioButton_keep_middle.TabStop = True
        Me.RadioButton_keep_middle.Text = "Hold"
        Me.RadioButton_keep_middle.UseVisualStyleBackColor = True
        '
        'RadioButton_keep_high
        '
        Me.RadioButton_keep_high.AutoSize = True
        Me.RadioButton_keep_high.Location = New System.Drawing.Point(6, 61)
        Me.RadioButton_keep_high.Name = "RadioButton_keep_high"
        Me.RadioButton_keep_high.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_keep_high.Size = New System.Drawing.Size(53, 19)
        Me.RadioButton_keep_high.TabIndex = 1
        Me.RadioButton_keep_high.Text = "Hold "
        Me.RadioButton_keep_high.UseVisualStyleBackColor = True
        '
        'TextBox_chainage_high
        '
        Me.TextBox_chainage_high.Location = New System.Drawing.Point(65, 61)
        Me.TextBox_chainage_high.Name = "TextBox_chainage_high"
        Me.TextBox_chainage_high.Size = New System.Drawing.Size(114, 21)
        Me.TextBox_chainage_high.TabIndex = 0
        '
        'TextBox_chainage_low
        '
        Me.TextBox_chainage_low.Location = New System.Drawing.Point(65, 8)
        Me.TextBox_chainage_low.Name = "TextBox_chainage_low"
        Me.TextBox_chainage_low.Size = New System.Drawing.Size(114, 21)
        Me.TextBox_chainage_low.TabIndex = 0
        '
        'TextBox_amount
        '
        Me.TextBox_amount.Location = New System.Drawing.Point(218, 40)
        Me.TextBox_amount.Name = "TextBox_amount"
        Me.TextBox_amount.Size = New System.Drawing.Size(100, 21)
        Me.TextBox_amount.TabIndex = 2
        '
        'Button_plus
        '
        Me.Button_plus.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Button_plus.Location = New System.Drawing.Point(218, 68)
        Me.Button_plus.Name = "Button_plus"
        Me.Button_plus.Size = New System.Drawing.Size(48, 45)
        Me.Button_plus.TabIndex = 3
        Me.Button_plus.Text = "+"
        Me.Button_plus.UseVisualStyleBackColor = True
        '
        'Button_minus
        '
        Me.Button_minus.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Button_minus.Location = New System.Drawing.Point(275, 68)
        Me.Button_minus.Name = "Button_minus"
        Me.Button_minus.Size = New System.Drawing.Size(48, 45)
        Me.Button_minus.TabIndex = 3
        Me.Button_minus.Text = "-"
        Me.Button_minus.UseVisualStyleBackColor = True
        '
        'Button_pick_chainage
        '
        Me.Button_pick_chainage.Location = New System.Drawing.Point(65, 88)
        Me.Button_pick_chainage.Name = "Button_pick_chainage"
        Me.Button_pick_chainage.Size = New System.Drawing.Size(114, 24)
        Me.Button_pick_chainage.TabIndex = 4
        Me.Button_pick_chainage.Text = "Pick chainage"
        Me.Button_pick_chainage.UseVisualStyleBackColor = True
        '
        'Button_pick_amount
        '
        Me.Button_pick_amount.Location = New System.Drawing.Point(218, 10)
        Me.Button_pick_amount.Name = "Button_pick_amount"
        Me.Button_pick_amount.Size = New System.Drawing.Size(100, 24)
        Me.Button_pick_amount.TabIndex = 4
        Me.Button_pick_amount.Text = "Pick amount"
        Me.Button_pick_amount.UseVisualStyleBackColor = True
        '
        'Button_split
        '
        Me.Button_split.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_split.Location = New System.Drawing.Point(329, 10)
        Me.Button_split.Name = "Button_split"
        Me.Button_split.Size = New System.Drawing.Size(31, 103)
        Me.Button_split.TabIndex = 3
        Me.Button_split.Text = "S" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "P" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "L" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "I" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "T"
        Me.Button_split.UseVisualStyleBackColor = True
        '
        'CheckBox_low
        '
        Me.CheckBox_low.AutoSize = True
        Me.CheckBox_low.Location = New System.Drawing.Point(185, 13)
        Me.CheckBox_low.Name = "CheckBox_low"
        Me.CheckBox_low.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox_low.TabIndex = 5
        Me.CheckBox_low.UseVisualStyleBackColor = True
        '
        'CheckBox_high
        '
        Me.CheckBox_high.AutoSize = True
        Me.CheckBox_high.Location = New System.Drawing.Point(185, 64)
        Me.CheckBox_high.Name = "CheckBox_high"
        Me.CheckBox_high.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox_high.TabIndex = 5
        Me.CheckBox_high.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(185, 37)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox1.TabIndex = 5
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Chainage_operations_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(378, 121)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CheckBox_high)
        Me.Controls.Add(Me.CheckBox_low)
        Me.Controls.Add(Me.Button_pick_amount)
        Me.Controls.Add(Me.Button_pick_chainage)
        Me.Controls.Add(Me.Button_split)
        Me.Controls.Add(Me.Button_minus)
        Me.Controls.Add(Me.Button_plus)
        Me.Controls.Add(Me.TextBox_amount)
        Me.Controls.Add(Me.RadioButton_keep_high)
        Me.Controls.Add(Me.RadioButton_keep_middle)
        Me.Controls.Add(Me.RadioButton_keep_low)
        Me.Controls.Add(Me.TextBox_chainage_low)
        Me.Controls.Add(Me.TextBox_chainage_high)
        Me.Controls.Add(Me.TextBox_chainage_middle)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Chainage_operations_form"
        Me.Text = "Chainage_operations_form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_chainage_middle As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton_keep_low As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_keep_middle As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_keep_high As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox_chainage_high As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_chainage_low As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_amount As System.Windows.Forms.TextBox
    Friend WithEvents Button_plus As System.Windows.Forms.Button
    Friend WithEvents Button_minus As System.Windows.Forms.Button
    Friend WithEvents Button_pick_chainage As System.Windows.Forms.Button
    Friend WithEvents Button_pick_amount As System.Windows.Forms.Button
    Friend WithEvents Button_split As System.Windows.Forms.Button
    Friend WithEvents CheckBox_low As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_high As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
