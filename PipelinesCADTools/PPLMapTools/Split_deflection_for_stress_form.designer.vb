<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Split_deflection_for_stress_form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Split_deflection_for_stress_form))
        Me.Button_split = New System.Windows.Forms.Button()
        Me.TextBox_distance = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_number_of_splits = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox_double_joint = New System.Windows.Forms.CheckBox()
        Me.Button_calculate_MAX_BEND_ANGLE = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ComboBox_NPS = New System.Windows.Forms.ComboBox()
        Me.TextBox_tan_length = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label_IM2 = New System.Windows.Forms.Label()
        Me.Label_IM1 = New System.Windows.Forms.Label()
        Me.TextBox_max_bend_angle = New System.Windows.Forms.TextBox()
        Me.TextBox_joint_length_calcs = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.RadioButton_m = New System.Windows.Forms.RadioButton()
        Me.RadioButton_i = New System.Windows.Forms.RadioButton()
        Me.RadioButton_f = New System.Windows.Forms.RadioButton()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox_number_of_bends = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox_degg_per_bend = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_split
        '
        Me.Button_split.Location = New System.Drawing.Point(9, 62)
        Me.Button_split.Name = "Button_split"
        Me.Button_split.Size = New System.Drawing.Size(209, 32)
        Me.Button_split.TabIndex = 0
        Me.Button_split.Text = "Draw Split"
        Me.Button_split.UseVisualStyleBackColor = True
        '
        'TextBox_distance
        '
        Me.TextBox_distance.BackColor = System.Drawing.Color.White
        Me.TextBox_distance.ForeColor = System.Drawing.Color.Black
        Me.TextBox_distance.Location = New System.Drawing.Point(12, 8)
        Me.TextBox_distance.Name = "TextBox_distance"
        Me.TextBox_distance.Size = New System.Drawing.Size(94, 21)
        Me.TextBox_distance.TabIndex = 1
        Me.TextBox_distance.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(112, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Joint Length"
        '
        'TextBox_number_of_splits
        '
        Me.TextBox_number_of_splits.BackColor = System.Drawing.Color.White
        Me.TextBox_number_of_splits.ForeColor = System.Drawing.Color.Black
        Me.TextBox_number_of_splits.Location = New System.Drawing.Point(12, 35)
        Me.TextBox_number_of_splits.Name = "TextBox_number_of_splits"
        Me.TextBox_number_of_splits.Size = New System.Drawing.Size(94, 21)
        Me.TextBox_number_of_splits.TabIndex = 1
        Me.TextBox_number_of_splits.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(112, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Number of splits"
        '
        'CheckBox_double_joint
        '
        Me.CheckBox_double_joint.AutoSize = True
        Me.CheckBox_double_joint.Location = New System.Drawing.Point(31, 131)
        Me.CheckBox_double_joint.Name = "CheckBox_double_joint"
        Me.CheckBox_double_joint.Size = New System.Drawing.Size(92, 19)
        Me.CheckBox_double_joint.TabIndex = 3
        Me.CheckBox_double_joint.Text = "Double joint"
        Me.CheckBox_double_joint.UseVisualStyleBackColor = True
        '
        'Button_calculate_MAX_BEND_ANGLE
        '
        Me.Button_calculate_MAX_BEND_ANGLE.Location = New System.Drawing.Point(14, 164)
        Me.Button_calculate_MAX_BEND_ANGLE.Name = "Button_calculate_MAX_BEND_ANGLE"
        Me.Button_calculate_MAX_BEND_ANGLE.Size = New System.Drawing.Size(209, 34)
        Me.Button_calculate_MAX_BEND_ANGLE.TabIndex = 4
        Me.Button_calculate_MAX_BEND_ANGLE.Text = "Calculate Maximum Bend Angle"
        Me.Button_calculate_MAX_BEND_ANGLE.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.ComboBox_NPS)
        Me.Panel1.Controls.Add(Me.TextBox_tan_length)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label_IM2)
        Me.Panel1.Controls.Add(Me.Label_IM1)
        Me.Panel1.Controls.Add(Me.TextBox_number_of_bends)
        Me.Panel1.Controls.Add(Me.TextBox_max_bend_angle)
        Me.Panel1.Controls.Add(Me.TextBox_degg_per_bend)
        Me.Panel1.Controls.Add(Me.TextBox_joint_length_calcs)
        Me.Panel1.Controls.Add(Me.Button_calculate_MAX_BEND_ANGLE)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.CheckBox_double_joint)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(315, 289)
        Me.Panel1.TabIndex = 5
        '
        'ComboBox_NPS
        '
        Me.ComboBox_NPS.BackColor = System.Drawing.Color.White
        Me.ComboBox_NPS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_NPS.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_NPS.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_NPS.FormattingEnabled = True
        Me.ComboBox_NPS.Location = New System.Drawing.Point(128, 12)
        Me.ComboBox_NPS.Name = "ComboBox_NPS"
        Me.ComboBox_NPS.Size = New System.Drawing.Size(121, 32)
        Me.ComboBox_NPS.TabIndex = 6
        '
        'TextBox_tan_length
        '
        Me.TextBox_tan_length.BackColor = System.Drawing.Color.White
        Me.TextBox_tan_length.ForeColor = System.Drawing.Color.Black
        Me.TextBox_tan_length.Location = New System.Drawing.Point(31, 88)
        Me.TextBox_tan_length.Name = "TextBox_tan_length"
        Me.TextBox_tan_length.Size = New System.Drawing.Size(64, 21)
        Me.TextBox_tan_length.TabIndex = 1
        Me.TextBox_tan_length.Text = "1.8"
        Me.TextBox_tan_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(7, 94)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(21, 15)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "2x"
        '
        'Label_IM2
        '
        Me.Label_IM2.AutoSize = True
        Me.Label_IM2.Location = New System.Drawing.Point(210, 94)
        Me.Label_IM2.Name = "Label_IM2"
        Me.Label_IM2.Size = New System.Drawing.Size(26, 15)
        Me.Label_IM2.TabIndex = 2
        Me.Label_IM2.Text = "[m]"
        '
        'Label_IM1
        '
        Me.Label_IM1.AutoSize = True
        Me.Label_IM1.Location = New System.Drawing.Point(205, 57)
        Me.Label_IM1.Name = "Label_IM1"
        Me.Label_IM1.Size = New System.Drawing.Size(26, 15)
        Me.Label_IM1.TabIndex = 2
        Me.Label_IM1.Text = "[m]"
        '
        'TextBox_max_bend_angle
        '
        Me.TextBox_max_bend_angle.BackColor = System.Drawing.Color.LightSkyBlue
        Me.TextBox_max_bend_angle.ForeColor = System.Drawing.Color.Black
        Me.TextBox_max_bend_angle.Location = New System.Drawing.Point(14, 255)
        Me.TextBox_max_bend_angle.Name = "TextBox_max_bend_angle"
        Me.TextBox_max_bend_angle.ReadOnly = True
        Me.TextBox_max_bend_angle.Size = New System.Drawing.Size(94, 21)
        Me.TextBox_max_bend_angle.TabIndex = 1
        '
        'TextBox_joint_length_calcs
        '
        Me.TextBox_joint_length_calcs.BackColor = System.Drawing.Color.White
        Me.TextBox_joint_length_calcs.ForeColor = System.Drawing.Color.Black
        Me.TextBox_joint_length_calcs.Location = New System.Drawing.Point(31, 51)
        Me.TextBox_joint_length_calcs.Name = "TextBox_joint_length_calcs"
        Me.TextBox_joint_length_calcs.Size = New System.Drawing.Size(64, 21)
        Me.TextBox_joint_length_calcs.TabIndex = 1
        Me.TextBox_joint_length_calcs.Text = "24"
        Me.TextBox_joint_length_calcs.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(101, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 30)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Tangent Length " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "required at welds"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(114, 261)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 15)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Maximum bend angle"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(78, 15)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 24)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "[in]"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(23, 15)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 24)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "NPS"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(122, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Joint Length"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_distance)
        Me.Panel2.Controls.Add(Me.Button_split)
        Me.Panel2.Controls.Add(Me.TextBox_number_of_splits)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Location = New System.Drawing.Point(333, 14)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(235, 113)
        Me.Panel2.TabIndex = 5
        '
        'Panel3
        '
        Me.Panel3.BackgroundImage = CType(resources.GetObject("Panel3.BackgroundImage"), System.Drawing.Image)
        Me.Panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Location = New System.Drawing.Point(12, 307)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(559, 61)
        Me.Panel3.TabIndex = 5
        '
        'RadioButton_m
        '
        Me.RadioButton_m.AutoSize = True
        Me.RadioButton_m.Checked = True
        Me.RadioButton_m.Location = New System.Drawing.Point(333, 144)
        Me.RadioButton_m.Name = "RadioButton_m"
        Me.RadioButton_m.Size = New System.Drawing.Size(65, 19)
        Me.RadioButton_m.TabIndex = 6
        Me.RadioButton_m.TabStop = True
        Me.RadioButton_m.Text = "Meters"
        Me.RadioButton_m.UseVisualStyleBackColor = True
        '
        'RadioButton_i
        '
        Me.RadioButton_i.AutoSize = True
        Me.RadioButton_i.Location = New System.Drawing.Point(333, 169)
        Me.RadioButton_i.Name = "RadioButton_i"
        Me.RadioButton_i.Size = New System.Drawing.Size(63, 19)
        Me.RadioButton_i.TabIndex = 6
        Me.RadioButton_i.Text = "Inches"
        Me.RadioButton_i.UseVisualStyleBackColor = True
        '
        'RadioButton_f
        '
        Me.RadioButton_f.AutoSize = True
        Me.RadioButton_f.Location = New System.Drawing.Point(333, 194)
        Me.RadioButton_f.Name = "RadioButton_f"
        Me.RadioButton_f.Size = New System.Drawing.Size(49, 19)
        Me.RadioButton_f.TabIndex = 6
        Me.RadioButton_f.Text = "Feet"
        Me.RadioButton_f.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(114, 201)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(105, 15)
        Me.Label8.TabIndex = 2
        Me.Label8.Text = "Number of Bends"
        '
        'TextBox_number_of_bends
        '
        Me.TextBox_number_of_bends.BackColor = System.Drawing.Color.LightSkyBlue
        Me.TextBox_number_of_bends.ForeColor = System.Drawing.Color.Black
        Me.TextBox_number_of_bends.Location = New System.Drawing.Point(14, 201)
        Me.TextBox_number_of_bends.Name = "TextBox_number_of_bends"
        Me.TextBox_number_of_bends.ReadOnly = True
        Me.TextBox_number_of_bends.Size = New System.Drawing.Size(94, 21)
        Me.TextBox_number_of_bends.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(122, 234)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 15)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "Degrees per Bend"
        '
        'TextBox_degg_per_bend
        '
        Me.TextBox_degg_per_bend.BackColor = System.Drawing.Color.White
        Me.TextBox_degg_per_bend.ForeColor = System.Drawing.Color.Black
        Me.TextBox_degg_per_bend.Location = New System.Drawing.Point(31, 228)
        Me.TextBox_degg_per_bend.Name = "TextBox_degg_per_bend"
        Me.TextBox_degg_per_bend.Size = New System.Drawing.Size(64, 21)
        Me.TextBox_degg_per_bend.TabIndex = 0
        Me.TextBox_degg_per_bend.Text = "1"
        Me.TextBox_degg_per_bend.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Split_deflection_for_stress_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(578, 376)
        Me.Controls.Add(Me.RadioButton_f)
        Me.Controls.Add(Me.RadioButton_i)
        Me.Controls.Add(Me.RadioButton_m)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Split_deflection_for_stress_form"
        Me.Text = "Bend Calcs"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_split As System.Windows.Forms.Button
    Friend WithEvents TextBox_distance As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_number_of_splits As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_double_joint As System.Windows.Forms.CheckBox
    Friend WithEvents Button_calculate_MAX_BEND_ANGLE As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ComboBox_NPS As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox_tan_length As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_max_bend_angle As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_joint_length_calcs As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label_IM2 As System.Windows.Forms.Label
    Friend WithEvents Label_IM1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents RadioButton_m As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_i As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_f As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox_number_of_bends As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_degg_per_bend As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class
