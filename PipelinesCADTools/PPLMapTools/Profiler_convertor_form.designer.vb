<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Profiler_convertor_form
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
        Me.Button_Draw = New System.Windows.Forms.Button()
        Me.TextBox_Horizontal_scale = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TextBox_printing_scale = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_Vertical_scale = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ComboBox_Horizontal_scale_HMM = New System.Windows.Forms.ComboBox()
        Me.ComboBox_Vertical_scale_HMM = New System.Windows.Forms.ComboBox()
        Me.TextBox_Printing_scale_HMM = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox_text_style = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ComboBox_linetype = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBox_PROFILE_LENGTH = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TextBox_vertical_increment = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.TextBox_horizontal_increment = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TextBox_nr_vert_labels = New System.Windows.Forms.TextBox()
        Me.CheckBox_recalculate_chainage = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_Draw
        '
        Me.Button_Draw.Location = New System.Drawing.Point(12, 272)
        Me.Button_Draw.Name = "Button_Draw"
        Me.Button_Draw.Size = New System.Drawing.Size(431, 33)
        Me.Button_Draw.TabIndex = 0
        Me.Button_Draw.Text = "Draw"
        Me.Button_Draw.UseVisualStyleBackColor = True
        '
        'TextBox_Horizontal_scale
        '
        Me.TextBox_Horizontal_scale.Location = New System.Drawing.Point(3, 3)
        Me.TextBox_Horizontal_scale.Name = "TextBox_Horizontal_scale"
        Me.TextBox_Horizontal_scale.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_Horizontal_scale.TabIndex = 0
        Me.TextBox_Horizontal_scale.Text = "1000"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(98, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 14)
        Me.Label1.TabIndex = 200
        Me.Label1.Text = "Horizontal Scale "
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TextBox_printing_scale)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.TextBox_Vertical_scale)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.TextBox_Horizontal_scale)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(12, 64)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(213, 91)
        Me.Panel1.TabIndex = 3
        '
        'TextBox_printing_scale
        '
        Me.TextBox_printing_scale.Location = New System.Drawing.Point(3, 59)
        Me.TextBox_printing_scale.Name = "TextBox_printing_scale"
        Me.TextBox_printing_scale.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_printing_scale.TabIndex = 2
        Me.TextBox_printing_scale.Text = "1000"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(98, 62)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 14)
        Me.Label3.TabIndex = 204
        Me.Label3.Text = "Printing Scale "
        '
        'TextBox_Vertical_scale
        '
        Me.TextBox_Vertical_scale.Location = New System.Drawing.Point(3, 31)
        Me.TextBox_Vertical_scale.Name = "TextBox_Vertical_scale"
        Me.TextBox_Vertical_scale.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_Vertical_scale.TabIndex = 1
        Me.TextBox_Vertical_scale.Text = "1000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(98, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 14)
        Me.Label2.TabIndex = 202
        Me.Label2.Text = "Vertical Scale "
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Location = New System.Drawing.Point(36, 13)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(145, 44)
        Me.Panel2.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(24, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 14)
        Me.Label4.TabIndex = 203
        Me.Label4.Text = "GRAPH SOURCE"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.ComboBox_Horizontal_scale_HMM)
        Me.Panel3.Controls.Add(Me.ComboBox_Vertical_scale_HMM)
        Me.Panel3.Controls.Add(Me.TextBox_Printing_scale_HMM)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Location = New System.Drawing.Point(239, 64)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(202, 91)
        Me.Panel3.TabIndex = 3
        '
        'ComboBox_Horizontal_scale_HMM
        '
        Me.ComboBox_Horizontal_scale_HMM.BackColor = System.Drawing.Color.Silver
        Me.ComboBox_Horizontal_scale_HMM.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_Horizontal_scale_HMM.FormattingEnabled = True
        Me.ComboBox_Horizontal_scale_HMM.Items.AddRange(New Object() {"1000", "2000"})
        Me.ComboBox_Horizontal_scale_HMM.Location = New System.Drawing.Point(3, 2)
        Me.ComboBox_Horizontal_scale_HMM.Name = "ComboBox_Horizontal_scale_HMM"
        Me.ComboBox_Horizontal_scale_HMM.Size = New System.Drawing.Size(89, 22)
        Me.ComboBox_Horizontal_scale_HMM.TabIndex = 205
        '
        'ComboBox_Vertical_scale_HMM
        '
        Me.ComboBox_Vertical_scale_HMM.BackColor = System.Drawing.Color.Silver
        Me.ComboBox_Vertical_scale_HMM.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_Vertical_scale_HMM.FormattingEnabled = True
        Me.ComboBox_Vertical_scale_HMM.Items.AddRange(New Object() {"1000", "100", "200", "500"})
        Me.ComboBox_Vertical_scale_HMM.Location = New System.Drawing.Point(3, 31)
        Me.ComboBox_Vertical_scale_HMM.Name = "ComboBox_Vertical_scale_HMM"
        Me.ComboBox_Vertical_scale_HMM.Size = New System.Drawing.Size(89, 22)
        Me.ComboBox_Vertical_scale_HMM.TabIndex = 205
        '
        'TextBox_Printing_scale_HMM
        '
        Me.TextBox_Printing_scale_HMM.BackColor = System.Drawing.Color.Silver
        Me.TextBox_Printing_scale_HMM.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Printing_scale_HMM.Location = New System.Drawing.Point(3, 59)
        Me.TextBox_Printing_scale_HMM.Name = "TextBox_Printing_scale_HMM"
        Me.TextBox_Printing_scale_HMM.ReadOnly = True
        Me.TextBox_Printing_scale_HMM.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_Printing_scale_HMM.TabIndex = 5
        Me.TextBox_Printing_scale_HMM.Text = "1000"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(98, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 14)
        Me.Label5.TabIndex = 204
        Me.Label5.Text = "Printing Scale "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(98, 34)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 14)
        Me.Label6.TabIndex = 202
        Me.Label6.Text = "Vertical Scale "
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(98, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(97, 14)
        Me.Label7.TabIndex = 200
        Me.Label7.Text = "Horizontal Scale "
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Label8)
        Me.Panel4.Location = New System.Drawing.Point(263, 13)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(145, 44)
        Me.Panel4.TabIndex = 4
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 13)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(116, 14)
        Me.Label8.TabIndex = 203
        Me.Label8.Text = "GRAPH DESTINATION"
        '
        'ComboBox_text_style
        '
        Me.ComboBox_text_style.BackColor = System.Drawing.Color.Silver
        Me.ComboBox_text_style.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_text_style.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_text_style.FormattingEnabled = True
        Me.ComboBox_text_style.Items.AddRange(New Object() {"ROMANS"})
        Me.ComboBox_text_style.Location = New System.Drawing.Point(12, 162)
        Me.ComboBox_text_style.Name = "ComboBox_text_style"
        Me.ComboBox_text_style.Size = New System.Drawing.Size(140, 22)
        Me.ComboBox_text_style.TabIndex = 205
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(161, 165)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(60, 14)
        Me.Label9.TabIndex = 204
        Me.Label9.Text = "Text Style"
        '
        'ComboBox_linetype
        '
        Me.ComboBox_linetype.BackColor = System.Drawing.Color.Silver
        Me.ComboBox_linetype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_linetype.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_linetype.FormattingEnabled = True
        Me.ComboBox_linetype.Items.AddRange(New Object() {"ROMANS"})
        Me.ComboBox_linetype.Location = New System.Drawing.Point(12, 196)
        Me.ComboBox_linetype.Name = "ComboBox_linetype"
        Me.ComboBox_linetype.Size = New System.Drawing.Size(140, 22)
        Me.ComboBox_linetype.TabIndex = 205
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(161, 199)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(60, 14)
        Me.Label10.TabIndex = 204
        Me.Label10.Text = "Line Type"
        '
        'TextBox_PROFILE_LENGTH
        '
        Me.TextBox_PROFILE_LENGTH.Location = New System.Drawing.Point(342, 162)
        Me.TextBox_PROFILE_LENGTH.Name = "TextBox_PROFILE_LENGTH"
        Me.TextBox_PROFILE_LENGTH.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_PROFILE_LENGTH.TabIndex = 2
        Me.TextBox_PROFILE_LENGTH.Text = "750"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(236, 165)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(91, 14)
        Me.Label11.TabIndex = 204
        Me.Label11.Text = "Profile H length"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(236, 193)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(75, 14)
        Me.Label12.TabIndex = 204
        Me.Label12.Text = "V Increment"
        '
        'TextBox_vertical_increment
        '
        Me.TextBox_vertical_increment.Location = New System.Drawing.Point(342, 190)
        Me.TextBox_vertical_increment.Name = "TextBox_vertical_increment"
        Me.TextBox_vertical_increment.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_vertical_increment.TabIndex = 2
        Me.TextBox_vertical_increment.Text = "10"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(236, 220)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(74, 14)
        Me.Label13.TabIndex = 204
        Me.Label13.Text = "H Increment"
        '
        'TextBox_horizontal_increment
        '
        Me.TextBox_horizontal_increment.Location = New System.Drawing.Point(342, 216)
        Me.TextBox_horizontal_increment.Name = "TextBox_horizontal_increment"
        Me.TextBox_horizontal_increment.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_horizontal_increment.TabIndex = 2
        Me.TextBox_horizontal_increment.Text = "100"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(235, 248)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(68, 14)
        Me.Label14.TabIndex = 204
        Me.Label14.Text = "No V labels"
        '
        'TextBox_nr_vert_labels
        '
        Me.TextBox_nr_vert_labels.Location = New System.Drawing.Point(341, 244)
        Me.TextBox_nr_vert_labels.Name = "TextBox_nr_vert_labels"
        Me.TextBox_nr_vert_labels.Size = New System.Drawing.Size(89, 20)
        Me.TextBox_nr_vert_labels.TabIndex = 2
        Me.TextBox_nr_vert_labels.Text = "11"
        '
        'CheckBox_recalculate_chainage
        '
        Me.CheckBox_recalculate_chainage.AutoSize = True
        Me.CheckBox_recalculate_chainage.Location = New System.Drawing.Point(12, 234)
        Me.CheckBox_recalculate_chainage.Name = "CheckBox_recalculate_chainage"
        Me.CheckBox_recalculate_chainage.Size = New System.Drawing.Size(142, 32)
        Me.CheckBox_recalculate_chainage.TabIndex = 206
        Me.CheckBox_recalculate_chainage.Text = "Recalculate Chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Based on new 0+000"
        Me.CheckBox_recalculate_chainage.UseVisualStyleBackColor = True
        '
        'Profiler_convertor_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(463, 317)
        Me.Controls.Add(Me.CheckBox_recalculate_chainage)
        Me.Controls.Add(Me.TextBox_nr_vert_labels)
        Me.Controls.Add(Me.TextBox_horizontal_increment)
        Me.Controls.Add(Me.TextBox_vertical_increment)
        Me.Controls.Add(Me.TextBox_PROFILE_LENGTH)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.ComboBox_linetype)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.ComboBox_text_style)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button_Draw)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Profiler_convertor_form"
        Me.Text = "Profile Convertor"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_Draw As System.Windows.Forms.Button
    Friend WithEvents TextBox_Horizontal_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_printing_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Vertical_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_Printing_scale_HMM As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_Vertical_scale_HMM As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_text_style As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_linetype As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBox_PROFILE_LENGTH As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextBox_vertical_increment As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TextBox_horizontal_increment As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox_nr_vert_labels As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_recalculate_chainage As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox_Horizontal_scale_HMM As System.Windows.Forms.ComboBox
End Class
