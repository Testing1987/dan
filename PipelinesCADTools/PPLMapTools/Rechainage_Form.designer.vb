<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Rechainage_Form
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
        Me.TextBox_BEG_chainage = New System.Windows.Forms.TextBox()
        Me.TextBox_END_chainage = New System.Windows.Forms.TextBox()
        Me.TextBox_diference = New System.Windows.Forms.TextBox()
        Me.Button_pick = New System.Windows.Forms.Button()
        Me.Button_clear = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button_UPDATE_LENGTH = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_screw_anchors_number = New System.Windows.Forms.TextBox()
        Me.TextBox_screw_anchors_spacing = New System.Windows.Forms.TextBox()
        Me.Button_calculate_screw_anchors = New System.Windows.Forms.Button()
        Me.Panel_color = New System.Windows.Forms.Panel()
        Me.Label_updated_chain1_for_screw_anchors = New System.Windows.Forms.Label()
        Me.Label_updated_chain2_for_screw_anchors = New System.Windows.Forms.Label()
        Me.Panel_screw_anchors = New System.Windows.Forms.Panel()
        Me.CheckBox_BEG_STA = New System.Windows.Forms.CheckBox()
        Me.CheckBox_END_STA = New System.Windows.Forms.CheckBox()
        Me.CheckBox_diference = New System.Windows.Forms.CheckBox()
        Me.RadioButton_minus = New System.Windows.Forms.RadioButton()
        Me.RadioButton_plus = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_switch = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button_RECHAINAGE = New System.Windows.Forms.Button()
        Me.TextBox_AMOUNT_FOR_RECHAIN = New System.Windows.Forms.TextBox()
        Me.Button_push_chainage = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel_color.SuspendLayout()
        Me.Panel_screw_anchors.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_BEG_chainage
        '
        Me.TextBox_BEG_chainage.BackColor = System.Drawing.Color.White
        Me.TextBox_BEG_chainage.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_BEG_chainage.ForeColor = System.Drawing.Color.Black
        Me.TextBox_BEG_chainage.Location = New System.Drawing.Point(33, 7)
        Me.TextBox_BEG_chainage.Name = "TextBox_BEG_chainage"
        Me.TextBox_BEG_chainage.Size = New System.Drawing.Size(122, 22)
        Me.TextBox_BEG_chainage.TabIndex = 0
        '
        'TextBox_END_chainage
        '
        Me.TextBox_END_chainage.BackColor = System.Drawing.Color.White
        Me.TextBox_END_chainage.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_END_chainage.ForeColor = System.Drawing.Color.Black
        Me.TextBox_END_chainage.Location = New System.Drawing.Point(33, 35)
        Me.TextBox_END_chainage.Name = "TextBox_END_chainage"
        Me.TextBox_END_chainage.Size = New System.Drawing.Size(122, 22)
        Me.TextBox_END_chainage.TabIndex = 1
        '
        'TextBox_diference
        '
        Me.TextBox_diference.BackColor = System.Drawing.Color.PeachPuff
        Me.TextBox_diference.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_diference.ForeColor = System.Drawing.Color.Black
        Me.TextBox_diference.Location = New System.Drawing.Point(33, 63)
        Me.TextBox_diference.Name = "TextBox_diference"
        Me.TextBox_diference.ReadOnly = True
        Me.TextBox_diference.Size = New System.Drawing.Size(122, 22)
        Me.TextBox_diference.TabIndex = 100
        '
        'Button_pick
        '
        Me.Button_pick.BackColor = System.Drawing.Color.Green
        Me.Button_pick.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick.ForeColor = System.Drawing.Color.White
        Me.Button_pick.Location = New System.Drawing.Point(4, 3)
        Me.Button_pick.Name = "Button_pick"
        Me.Button_pick.Size = New System.Drawing.Size(191, 41)
        Me.Button_pick.TabIndex = 100
        Me.Button_pick.Text = "Pick Chainage"
        Me.Button_pick.UseVisualStyleBackColor = False
        '
        'Button_clear
        '
        Me.Button_clear.BackColor = System.Drawing.Color.Red
        Me.Button_clear.Location = New System.Drawing.Point(2, 7)
        Me.Button_clear.Name = "Button_clear"
        Me.Button_clear.Size = New System.Drawing.Size(27, 50)
        Me.Button_clear.TabIndex = 7
        Me.Button_clear.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Button_UPDATE_LENGTH)
        Me.Panel1.Location = New System.Drawing.Point(12, 150)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(202, 53)
        Me.Panel1.TabIndex = 9
        '
        'Button_UPDATE_LENGTH
        '
        Me.Button_UPDATE_LENGTH.BackColor = System.Drawing.Color.DodgerBlue
        Me.Button_UPDATE_LENGTH.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_UPDATE_LENGTH.ForeColor = System.Drawing.Color.White
        Me.Button_UPDATE_LENGTH.Location = New System.Drawing.Point(4, 3)
        Me.Button_UPDATE_LENGTH.Name = "Button_UPDATE_LENGTH"
        Me.Button_UPDATE_LENGTH.Size = New System.Drawing.Size(191, 41)
        Me.Button_UPDATE_LENGTH.TabIndex = 100
        Me.Button_UPDATE_LENGTH.Text = "Update field"
        Me.Button_UPDATE_LENGTH.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Button_pick)
        Me.Panel2.Location = New System.Drawing.Point(12, 91)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(202, 53)
        Me.Panel2.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 7)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(92, 14)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Screw Anchors"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(47, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(18, 14)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "@"
        '
        'TextBox_screw_anchors_number
        '
        Me.TextBox_screw_anchors_number.BackColor = System.Drawing.Color.White
        Me.TextBox_screw_anchors_number.ForeColor = System.Drawing.Color.Black
        Me.TextBox_screw_anchors_number.Location = New System.Drawing.Point(2, 25)
        Me.TextBox_screw_anchors_number.Multiline = True
        Me.TextBox_screw_anchors_number.Name = "TextBox_screw_anchors_number"
        Me.TextBox_screw_anchors_number.Size = New System.Drawing.Size(39, 22)
        Me.TextBox_screw_anchors_number.TabIndex = 3
        '
        'TextBox_screw_anchors_spacing
        '
        Me.TextBox_screw_anchors_spacing.BackColor = System.Drawing.Color.White
        Me.TextBox_screw_anchors_spacing.ForeColor = System.Drawing.Color.Black
        Me.TextBox_screw_anchors_spacing.Location = New System.Drawing.Point(71, 26)
        Me.TextBox_screw_anchors_spacing.Multiline = True
        Me.TextBox_screw_anchors_spacing.Name = "TextBox_screw_anchors_spacing"
        Me.TextBox_screw_anchors_spacing.Size = New System.Drawing.Size(39, 22)
        Me.TextBox_screw_anchors_spacing.TabIndex = 4
        Me.TextBox_screw_anchors_spacing.Text = "23.8"
        '
        'Button_calculate_screw_anchors
        '
        Me.Button_calculate_screw_anchors.Location = New System.Drawing.Point(2, 54)
        Me.Button_calculate_screw_anchors.Name = "Button_calculate_screw_anchors"
        Me.Button_calculate_screw_anchors.Size = New System.Drawing.Size(191, 23)
        Me.Button_calculate_screw_anchors.TabIndex = 100
        Me.Button_calculate_screw_anchors.Text = "Calculate"
        Me.Button_calculate_screw_anchors.UseVisualStyleBackColor = True
        '
        'Panel_color
        '
        Me.Panel_color.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_color.Controls.Add(Me.Label_updated_chain1_for_screw_anchors)
        Me.Panel_color.Controls.Add(Me.Label_updated_chain2_for_screw_anchors)
        Me.Panel_color.Location = New System.Drawing.Point(115, 10)
        Me.Panel_color.Name = "Panel_color"
        Me.Panel_color.Size = New System.Drawing.Size(78, 38)
        Me.Panel_color.TabIndex = 9
        '
        'Label_updated_chain1_for_screw_anchors
        '
        Me.Label_updated_chain1_for_screw_anchors.AutoSize = True
        Me.Label_updated_chain1_for_screw_anchors.Location = New System.Drawing.Point(3, 0)
        Me.Label_updated_chain1_for_screw_anchors.Name = "Label_updated_chain1_for_screw_anchors"
        Me.Label_updated_chain1_for_screw_anchors.Size = New System.Drawing.Size(67, 14)
        Me.Label_updated_chain1_for_screw_anchors.TabIndex = 3
        Me.Label_updated_chain1_for_screw_anchors.Text = "Chainage 1"
        '
        'Label_updated_chain2_for_screw_anchors
        '
        Me.Label_updated_chain2_for_screw_anchors.AutoSize = True
        Me.Label_updated_chain2_for_screw_anchors.Location = New System.Drawing.Point(3, 18)
        Me.Label_updated_chain2_for_screw_anchors.Name = "Label_updated_chain2_for_screw_anchors"
        Me.Label_updated_chain2_for_screw_anchors.Size = New System.Drawing.Size(67, 14)
        Me.Label_updated_chain2_for_screw_anchors.TabIndex = 3
        Me.Label_updated_chain2_for_screw_anchors.Text = "Chainage 2"
        '
        'Panel_screw_anchors
        '
        Me.Panel_screw_anchors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_screw_anchors.Controls.Add(Me.Panel_color)
        Me.Panel_screw_anchors.Controls.Add(Me.Button_calculate_screw_anchors)
        Me.Panel_screw_anchors.Controls.Add(Me.TextBox_screw_anchors_spacing)
        Me.Panel_screw_anchors.Controls.Add(Me.TextBox_screw_anchors_number)
        Me.Panel_screw_anchors.Controls.Add(Me.Label4)
        Me.Panel_screw_anchors.Controls.Add(Me.Label5)
        Me.Panel_screw_anchors.Location = New System.Drawing.Point(458, 13)
        Me.Panel_screw_anchors.Name = "Panel_screw_anchors"
        Me.Panel_screw_anchors.Size = New System.Drawing.Size(206, 87)
        Me.Panel_screw_anchors.TabIndex = 9
        '
        'CheckBox_BEG_STA
        '
        Me.CheckBox_BEG_STA.AutoSize = True
        Me.CheckBox_BEG_STA.Location = New System.Drawing.Point(240, 7)
        Me.CheckBox_BEG_STA.Name = "CheckBox_BEG_STA"
        Me.CheckBox_BEG_STA.Size = New System.Drawing.Size(105, 18)
        Me.CheckBox_BEG_STA.TabIndex = 102
        Me.CheckBox_BEG_STA.Text = "BEGIN STATION"
        Me.CheckBox_BEG_STA.UseVisualStyleBackColor = True
        '
        'CheckBox_END_STA
        '
        Me.CheckBox_END_STA.AutoSize = True
        Me.CheckBox_END_STA.Location = New System.Drawing.Point(240, 35)
        Me.CheckBox_END_STA.Name = "CheckBox_END_STA"
        Me.CheckBox_END_STA.Size = New System.Drawing.Size(94, 18)
        Me.CheckBox_END_STA.TabIndex = 102
        Me.CheckBox_END_STA.Text = "END STATION"
        Me.CheckBox_END_STA.UseVisualStyleBackColor = True
        '
        'CheckBox_diference
        '
        Me.CheckBox_diference.AutoSize = True
        Me.CheckBox_diference.Location = New System.Drawing.Point(240, 59)
        Me.CheckBox_diference.Name = "CheckBox_diference"
        Me.CheckBox_diference.Size = New System.Drawing.Size(81, 18)
        Me.CheckBox_diference.TabIndex = 102
        Me.CheckBox_diference.Text = "Ammount"
        Me.CheckBox_diference.UseVisualStyleBackColor = True
        '
        'RadioButton_minus
        '
        Me.RadioButton_minus.AutoSize = True
        Me.RadioButton_minus.Checked = True
        Me.RadioButton_minus.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_minus.Location = New System.Drawing.Point(3, 3)
        Me.RadioButton_minus.Name = "RadioButton_minus"
        Me.RadioButton_minus.Size = New System.Drawing.Size(30, 22)
        Me.RadioButton_minus.TabIndex = 103
        Me.RadioButton_minus.TabStop = True
        Me.RadioButton_minus.Text = "-"
        Me.RadioButton_minus.UseVisualStyleBackColor = True
        '
        'RadioButton_plus
        '
        Me.RadioButton_plus.AutoSize = True
        Me.RadioButton_plus.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_plus.Location = New System.Drawing.Point(3, 31)
        Me.RadioButton_plus.Name = "RadioButton_plus"
        Me.RadioButton_plus.Size = New System.Drawing.Size(35, 22)
        Me.RadioButton_plus.TabIndex = 103
        Me.RadioButton_plus.Text = "+"
        Me.RadioButton_plus.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.RadioButton_plus)
        Me.Panel3.Controls.Add(Me.RadioButton_minus)
        Me.Panel3.Location = New System.Drawing.Point(194, 7)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(40, 63)
        Me.Panel3.TabIndex = 9
        '
        'Button_switch
        '
        Me.Button_switch.BackColor = System.Drawing.Color.Yellow
        Me.Button_switch.Location = New System.Drawing.Point(161, 7)
        Me.Button_switch.Name = "Button_switch"
        Me.Button_switch.Size = New System.Drawing.Size(27, 56)
        Me.Button_switch.TabIndex = 7
        Me.Button_switch.UseVisualStyleBackColor = False
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Button_RECHAINAGE)
        Me.Panel4.Controls.Add(Me.TextBox_AMOUNT_FOR_RECHAIN)
        Me.Panel4.Location = New System.Drawing.Point(12, 209)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(202, 84)
        Me.Panel4.TabIndex = 9
        '
        'Button_RECHAINAGE
        '
        Me.Button_RECHAINAGE.BackColor = System.Drawing.Color.DodgerBlue
        Me.Button_RECHAINAGE.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_RECHAINAGE.ForeColor = System.Drawing.Color.White
        Me.Button_RECHAINAGE.Location = New System.Drawing.Point(3, 31)
        Me.Button_RECHAINAGE.Name = "Button_RECHAINAGE"
        Me.Button_RECHAINAGE.Size = New System.Drawing.Size(192, 41)
        Me.Button_RECHAINAGE.TabIndex = 100
        Me.Button_RECHAINAGE.Text = "RECHAINAGE"
        Me.Button_RECHAINAGE.UseVisualStyleBackColor = False
        '
        'TextBox_AMOUNT_FOR_RECHAIN
        '
        Me.TextBox_AMOUNT_FOR_RECHAIN.BackColor = System.Drawing.Color.White
        Me.TextBox_AMOUNT_FOR_RECHAIN.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_AMOUNT_FOR_RECHAIN.ForeColor = System.Drawing.Color.Black
        Me.TextBox_AMOUNT_FOR_RECHAIN.Location = New System.Drawing.Point(3, 3)
        Me.TextBox_AMOUNT_FOR_RECHAIN.Name = "TextBox_AMOUNT_FOR_RECHAIN"
        Me.TextBox_AMOUNT_FOR_RECHAIN.Size = New System.Drawing.Size(192, 22)
        Me.TextBox_AMOUNT_FOR_RECHAIN.TabIndex = 1
        '
        'Button_push_chainage
        '
        Me.Button_push_chainage.BackColor = System.Drawing.Color.DodgerBlue
        Me.Button_push_chainage.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_push_chainage.ForeColor = System.Drawing.Color.White
        Me.Button_push_chainage.Location = New System.Drawing.Point(240, 83)
        Me.Button_push_chainage.Name = "Button_push_chainage"
        Me.Button_push_chainage.Size = New System.Drawing.Size(105, 41)
        Me.Button_push_chainage.TabIndex = 100
        Me.Button_push_chainage.Text = "Push Chainage"
        Me.Button_push_chainage.UseVisualStyleBackColor = False
        '
        'Rechainage_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(353, 301)
        Me.Controls.Add(Me.Button_push_chainage)
        Me.Controls.Add(Me.CheckBox_diference)
        Me.Controls.Add(Me.CheckBox_END_STA)
        Me.Controls.Add(Me.CheckBox_BEG_STA)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel_screw_anchors)
        Me.Controls.Add(Me.Button_switch)
        Me.Controls.Add(Me.Button_clear)
        Me.Controls.Add(Me.TextBox_diference)
        Me.Controls.Add(Me.TextBox_END_chainage)
        Me.Controls.Add(Me.TextBox_BEG_chainage)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Rechainage_Form"
        Me.Text = "Rechainage"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel_color.ResumeLayout(False)
        Me.Panel_color.PerformLayout()
        Me.Panel_screw_anchors.ResumeLayout(False)
        Me.Panel_screw_anchors.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_BEG_chainage As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_END_chainage As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_diference As System.Windows.Forms.TextBox
    Friend WithEvents Button_pick As System.Windows.Forms.Button
    Friend WithEvents Button_clear As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button_UPDATE_LENGTH As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_screw_anchors_number As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_screw_anchors_spacing As System.Windows.Forms.TextBox
    Friend WithEvents Button_calculate_screw_anchors As System.Windows.Forms.Button
    Friend WithEvents Panel_color As System.Windows.Forms.Panel
    Friend WithEvents Label_updated_chain1_for_screw_anchors As System.Windows.Forms.Label
    Friend WithEvents Label_updated_chain2_for_screw_anchors As System.Windows.Forms.Label
    Friend WithEvents Panel_screw_anchors As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_BEG_STA As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_END_STA As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_diference As System.Windows.Forms.CheckBox
    Friend WithEvents RadioButton_minus As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_plus As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button_switch As System.Windows.Forms.Button
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Button_RECHAINAGE As System.Windows.Forms.Button
    Friend WithEvents TextBox_AMOUNT_FOR_RECHAIN As System.Windows.Forms.TextBox
    Friend WithEvents Button_push_chainage As System.Windows.Forms.Button
End Class
