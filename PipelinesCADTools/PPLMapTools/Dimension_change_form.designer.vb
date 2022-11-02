<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Dimension_change_form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Dimension_change_form))
        Me.Button_change = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_SCALE_PS = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_viewport_Scale = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadioButtonm = New System.Windows.Forms.RadioButton()
        Me.RadioButton_mm = New System.Windows.Forms.RadioButton()
        Me.CheckBox_Scale_linear = New System.Windows.Forms.CheckBox()
        Me.TextBox_printScale = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button_transfer_info = New System.Windows.Forms.Button()
        Me.Button_hide_show_ext_line_1 = New System.Windows.Forms.Button()
        Me.Button_hide_show_ext_line_2 = New System.Windows.Forms.Button()
        Me.Button_align_to_linear = New System.Windows.Forms.Button()
        Me.Button_double_arrow2 = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_open30_2 = New System.Windows.Forms.Button()
        Me.Button_open30_1 = New System.Windows.Forms.Button()
        Me.Button_double_arrow1 = New System.Windows.Forms.Button()
        Me.ComboBox_decimals = New System.Windows.Forms.ComboBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Button_dim_rotated = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_change
        '
        Me.Button_change.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_change.ForeColor = System.Drawing.Color.DarkRed
        Me.Button_change.Location = New System.Drawing.Point(297, 3)
        Me.Button_change.Name = "Button_change"
        Me.Button_change.Size = New System.Drawing.Size(78, 110)
        Me.Button_change.TabIndex = 3
        Me.Button_change.Text = "Change"
        Me.Button_change.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkRed
        Me.Label1.Location = New System.Drawing.Point(67, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Decimals"
        '
        'TextBox_SCALE_PS
        '
        Me.TextBox_SCALE_PS.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_SCALE_PS.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_SCALE_PS.ForeColor = System.Drawing.Color.DarkRed
        Me.TextBox_SCALE_PS.Location = New System.Drawing.Point(76, 41)
        Me.TextBox_SCALE_PS.Name = "TextBox_SCALE_PS"
        Me.TextBox_SCALE_PS.Size = New System.Drawing.Size(58, 21)
        Me.TextBox_SCALE_PS.TabIndex = 1
        Me.TextBox_SCALE_PS.Text = "1000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.DarkRed
        Me.Label2.Location = New System.Drawing.Point(18, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "SCALE 1:"
        '
        'TextBox_viewport_Scale
        '
        Me.TextBox_viewport_Scale.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_viewport_Scale.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_viewport_Scale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_viewport_Scale.Location = New System.Drawing.Point(85, 148)
        Me.TextBox_viewport_Scale.Name = "TextBox_viewport_Scale"
        Me.TextBox_viewport_Scale.Size = New System.Drawing.Size(90, 21)
        Me.TextBox_viewport_Scale.TabIndex = 3
        Me.TextBox_viewport_Scale.Text = "1"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(9, 130)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(139, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Viewport custom Scale"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.RadioButtonm)
        Me.Panel1.Controls.Add(Me.RadioButton_mm)
        Me.Panel1.Location = New System.Drawing.Point(207, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(67, 52)
        Me.Panel1.TabIndex = 4
        '
        'RadioButtonm
        '
        Me.RadioButtonm.AutoSize = True
        Me.RadioButtonm.Checked = True
        Me.RadioButtonm.ForeColor = System.Drawing.Color.DarkRed
        Me.RadioButtonm.Location = New System.Drawing.Point(4, 3)
        Me.RadioButtonm.Name = "RadioButtonm"
        Me.RadioButtonm.Size = New System.Drawing.Size(36, 19)
        Me.RadioButtonm.TabIndex = 0
        Me.RadioButtonm.TabStop = True
        Me.RadioButtonm.Text = "m"
        Me.RadioButtonm.UseVisualStyleBackColor = True
        '
        'RadioButton_mm
        '
        Me.RadioButton_mm.AutoSize = True
        Me.RadioButton_mm.ForeColor = System.Drawing.Color.DarkRed
        Me.RadioButton_mm.Location = New System.Drawing.Point(4, 28)
        Me.RadioButton_mm.Name = "RadioButton_mm"
        Me.RadioButton_mm.Size = New System.Drawing.Size(47, 19)
        Me.RadioButton_mm.TabIndex = 0
        Me.RadioButton_mm.Text = "mm"
        Me.RadioButton_mm.UseVisualStyleBackColor = True
        '
        'CheckBox_Scale_linear
        '
        Me.CheckBox_Scale_linear.AutoSize = True
        Me.CheckBox_Scale_linear.ForeColor = System.Drawing.Color.DarkRed
        Me.CheckBox_Scale_linear.Location = New System.Drawing.Point(20, 1)
        Me.CheckBox_Scale_linear.Name = "CheckBox_Scale_linear"
        Me.CheckBox_Scale_linear.Size = New System.Drawing.Size(162, 34)
        Me.CheckBox_Scale_linear.TabIndex = 5
        Me.CheckBox_Scale_linear.Text = "Change DimScaleLinear" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Parameter"
        Me.CheckBox_Scale_linear.UseVisualStyleBackColor = True
        '
        'TextBox_printScale
        '
        Me.TextBox_printScale.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_printScale.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_printScale.ForeColor = System.Drawing.Color.DarkSlateGray
        Me.TextBox_printScale.Location = New System.Drawing.Point(81, 198)
        Me.TextBox_printScale.Name = "TextBox_printScale"
        Me.TextBox_printScale.Size = New System.Drawing.Size(90, 21)
        Me.TextBox_printScale.TabIndex = 3
        Me.TextBox_printScale.Text = "1000"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(4, 177)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 15)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Printing Scale"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.DarkSlateGray
        Me.Label5.Location = New System.Drawing.Point(17, 201)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 15)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "SCALE 1:"
        '
        'Button_transfer_info
        '
        Me.Button_transfer_info.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_transfer_info.Location = New System.Drawing.Point(10, 428)
        Me.Button_transfer_info.Name = "Button_transfer_info"
        Me.Button_transfer_info.Size = New System.Drawing.Size(293, 43)
        Me.Button_transfer_info.TabIndex = 3
        Me.Button_transfer_info.Text = "Transfer info" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "between dimension objects"
        Me.Button_transfer_info.UseVisualStyleBackColor = True
        '
        'Button_hide_show_ext_line_1
        '
        Me.Button_hide_show_ext_line_1.BackColor = System.Drawing.Color.Yellow
        Me.Button_hide_show_ext_line_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_hide_show_ext_line_1.ForeColor = System.Drawing.Color.Black
        Me.Button_hide_show_ext_line_1.Location = New System.Drawing.Point(7, 6)
        Me.Button_hide_show_ext_line_1.Name = "Button_hide_show_ext_line_1"
        Me.Button_hide_show_ext_line_1.Size = New System.Drawing.Size(115, 45)
        Me.Button_hide_show_ext_line_1.TabIndex = 3
        Me.Button_hide_show_ext_line_1.Text = "Extension Line 1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Show - Hide"
        Me.Button_hide_show_ext_line_1.UseVisualStyleBackColor = False
        '
        'Button_hide_show_ext_line_2
        '
        Me.Button_hide_show_ext_line_2.BackColor = System.Drawing.Color.Yellow
        Me.Button_hide_show_ext_line_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_hide_show_ext_line_2.ForeColor = System.Drawing.Color.Black
        Me.Button_hide_show_ext_line_2.Location = New System.Drawing.Point(256, 6)
        Me.Button_hide_show_ext_line_2.Name = "Button_hide_show_ext_line_2"
        Me.Button_hide_show_ext_line_2.Size = New System.Drawing.Size(115, 45)
        Me.Button_hide_show_ext_line_2.TabIndex = 3
        Me.Button_hide_show_ext_line_2.Text = "Extension Line 2" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Show - Hide"
        Me.Button_hide_show_ext_line_2.UseVisualStyleBackColor = False
        '
        'Button_align_to_linear
        '
        Me.Button_align_to_linear.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_align_to_linear.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_align_to_linear.Location = New System.Drawing.Point(128, 57)
        Me.Button_align_to_linear.Name = "Button_align_to_linear"
        Me.Button_align_to_linear.Size = New System.Drawing.Size(122, 100)
        Me.Button_align_to_linear.TabIndex = 3
        Me.Button_align_to_linear.Text = "HORIZONTAL" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Aligned Dimension" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Linear Dimension"
        Me.Button_align_to_linear.UseVisualStyleBackColor = False
        '
        'Button_double_arrow2
        '
        Me.Button_double_arrow2.BackColor = System.Drawing.Color.SpringGreen
        Me.Button_double_arrow2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_double_arrow2.Image = CType(resources.GetObject("Button_double_arrow2.Image"), System.Drawing.Image)
        Me.Button_double_arrow2.Location = New System.Drawing.Point(256, 112)
        Me.Button_double_arrow2.Name = "Button_double_arrow2"
        Me.Button_double_arrow2.Size = New System.Drawing.Size(115, 45)
        Me.Button_double_arrow2.TabIndex = 3
        Me.Button_double_arrow2.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.CheckBox_Scale_linear)
        Me.Panel2.Controls.Add(Me.TextBox_SCALE_PS)
        Me.Panel2.Controls.Add(Me.Panel1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Location = New System.Drawing.Point(3, 44)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(281, 73)
        Me.Panel2.TabIndex = 6
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Button_hide_show_ext_line_1)
        Me.Panel3.Controls.Add(Me.Button_align_to_linear)
        Me.Panel3.Controls.Add(Me.Button_open30_2)
        Me.Panel3.Controls.Add(Me.Button_hide_show_ext_line_2)
        Me.Panel3.Controls.Add(Me.Button_open30_1)
        Me.Panel3.Controls.Add(Me.Button_double_arrow1)
        Me.Panel3.Controls.Add(Me.Button_double_arrow2)
        Me.Panel3.Location = New System.Drawing.Point(10, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(382, 170)
        Me.Panel3.TabIndex = 7
        '
        'Button_open30_2
        '
        Me.Button_open30_2.BackColor = System.Drawing.Color.Yellow
        Me.Button_open30_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_open30_2.Image = CType(resources.GetObject("Button_open30_2.Image"), System.Drawing.Image)
        Me.Button_open30_2.Location = New System.Drawing.Point(256, 57)
        Me.Button_open30_2.Name = "Button_open30_2"
        Me.Button_open30_2.Size = New System.Drawing.Size(115, 45)
        Me.Button_open30_2.TabIndex = 3
        Me.Button_open30_2.UseVisualStyleBackColor = False
        '
        'Button_open30_1
        '
        Me.Button_open30_1.BackColor = System.Drawing.Color.Yellow
        Me.Button_open30_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_open30_1.Image = CType(resources.GetObject("Button_open30_1.Image"), System.Drawing.Image)
        Me.Button_open30_1.Location = New System.Drawing.Point(7, 57)
        Me.Button_open30_1.Name = "Button_open30_1"
        Me.Button_open30_1.Size = New System.Drawing.Size(115, 45)
        Me.Button_open30_1.TabIndex = 3
        Me.Button_open30_1.UseVisualStyleBackColor = False
        '
        'Button_double_arrow1
        '
        Me.Button_double_arrow1.BackColor = System.Drawing.Color.SpringGreen
        Me.Button_double_arrow1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_double_arrow1.Image = CType(resources.GetObject("Button_double_arrow1.Image"), System.Drawing.Image)
        Me.Button_double_arrow1.Location = New System.Drawing.Point(7, 112)
        Me.Button_double_arrow1.Name = "Button_double_arrow1"
        Me.Button_double_arrow1.Size = New System.Drawing.Size(115, 45)
        Me.Button_double_arrow1.TabIndex = 3
        Me.Button_double_arrow1.UseVisualStyleBackColor = False
        '
        'ComboBox_decimals
        '
        Me.ComboBox_decimals.BackColor = System.Drawing.Color.Gainsboro
        Me.ComboBox_decimals.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_decimals.ForeColor = System.Drawing.Color.DarkRed
        Me.ComboBox_decimals.FormattingEnabled = True
        Me.ComboBox_decimals.Items.AddRange(New Object() {"0", "1", "2", "3", "4", "5", "6", "7", "8"})
        Me.ComboBox_decimals.Location = New System.Drawing.Point(3, 15)
        Me.ComboBox_decimals.Name = "ComboBox_decimals"
        Me.ComboBox_decimals.Size = New System.Drawing.Size(53, 23)
        Me.ComboBox_decimals.TabIndex = 8
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Button_change)
        Me.Panel4.Location = New System.Drawing.Point(10, 188)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(382, 119)
        Me.Panel4.TabIndex = 6
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Controls.Add(Me.ComboBox_decimals)
        Me.Panel5.Controls.Add(Me.Label4)
        Me.Panel5.Controls.Add(Me.Label5)
        Me.Panel5.Controls.Add(Me.TextBox_printScale)
        Me.Panel5.Controls.Add(Me.Label1)
        Me.Panel5.Controls.Add(Me.TextBox_viewport_Scale)
        Me.Panel5.Controls.Add(Me.Panel2)
        Me.Panel5.Controls.Add(Me.Label3)
        Me.Panel5.Location = New System.Drawing.Point(10, 188)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(293, 234)
        Me.Panel5.TabIndex = 6
        '
        'Button_dim_rotated
        '
        Me.Button_dim_rotated.BackColor = System.Drawing.Color.SpringGreen
        Me.Button_dim_rotated.Location = New System.Drawing.Point(309, 313)
        Me.Button_dim_rotated.Name = "Button_dim_rotated"
        Me.Button_dim_rotated.Size = New System.Drawing.Size(87, 158)
        Me.Button_dim_rotated.TabIndex = 8
        Me.Button_dim_rotated.Text = "Dim Linear" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Rotated"
        Me.Button_dim_rotated.UseVisualStyleBackColor = False
        '
        'Dimension_change_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(399, 474)
        Me.Controls.Add(Me.Button_dim_rotated)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Button_transfer_info)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "Dimension_change_form"
        Me.Text = "Dimension Edit"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button_change As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_SCALE_PS As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_viewport_Scale As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadioButtonm As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_mm As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBox_Scale_linear As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox_printScale As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button_transfer_info As System.Windows.Forms.Button
    Friend WithEvents Button_hide_show_ext_line_1 As System.Windows.Forms.Button
    Friend WithEvents Button_hide_show_ext_line_2 As System.Windows.Forms.Button
    Friend WithEvents Button_align_to_linear As System.Windows.Forms.Button
    Friend WithEvents Button_double_arrow1 As System.Windows.Forms.Button
    Friend WithEvents Button_double_arrow2 As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button_open30_1 As System.Windows.Forms.Button
    Friend WithEvents Button_open30_2 As System.Windows.Forms.Button
    Friend WithEvents ComboBox_decimals As System.Windows.Forms.ComboBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Button_dim_rotated As System.Windows.Forms.Button
End Class
