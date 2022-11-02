<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HDD_Form
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
        Me.TextBox_angle_left = New System.Windows.Forms.TextBox()
        Me.TextBox_angle_right = New System.Windows.Forms.TextBox()
        Me.TextBox_radius_Left = New System.Windows.Forms.TextBox()
        Me.TextBox_straight_length = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_Radius_right = New System.Windows.Forms.TextBox()
        Me.Label_STRAIGHT_LEFT = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_3d_angle_left = New System.Windows.Forms.TextBox()
        Me.Label_3d_angle_left = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox_3d_angle_right = New System.Windows.Forms.TextBox()
        Me.Label_3d_angle_right = New System.Windows.Forms.Label()
        Me.Button_create_hdd = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TextBox_depth = New System.Windows.Forms.TextBox()
        Me.CheckBox_pick_start_end = New System.Windows.Forms.CheckBox()
        Me.CheckBox_1DEC = New System.Windows.Forms.CheckBox()
        Me.CheckBox_0DEC = New System.Windows.Forms.CheckBox()
        Me.Panel_formating = New System.Windows.Forms.Panel()
        Me.ComboBox_layer_hdd_cl = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layer_text = New System.Windows.Forms.ComboBox()
        Me.ComboBox_text_styles = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Button_3d_curve = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel_straight_portion = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel_formating.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel_straight_portion.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_angle_left
        '
        Me.TextBox_angle_left.Location = New System.Drawing.Point(111, 3)
        Me.TextBox_angle_left.Name = "TextBox_angle_left"
        Me.TextBox_angle_left.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_angle_left.TabIndex = 1
        Me.TextBox_angle_left.Text = "15"
        '
        'TextBox_angle_right
        '
        Me.TextBox_angle_right.Location = New System.Drawing.Point(3, 3)
        Me.TextBox_angle_right.Name = "TextBox_angle_right"
        Me.TextBox_angle_right.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_angle_right.TabIndex = 2
        Me.TextBox_angle_right.Text = "15"
        '
        'TextBox_radius_Left
        '
        Me.TextBox_radius_Left.Location = New System.Drawing.Point(111, 37)
        Me.TextBox_radius_Left.Name = "TextBox_radius_Left"
        Me.TextBox_radius_Left.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_radius_Left.TabIndex = 3
        Me.TextBox_radius_Left.Text = "500"
        '
        'TextBox_straight_length
        '
        Me.TextBox_straight_length.Location = New System.Drawing.Point(154, 3)
        Me.TextBox_straight_length.Name = "TextBox_straight_length"
        Me.TextBox_straight_length.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_straight_length.TabIndex = 5
        Me.TextBox_straight_length.Text = "50"
        Me.TextBox_straight_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 14)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Depth"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(66, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 14)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Angle Right [DD]"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 14)
        Me.Label3.TabIndex = 100
        Me.Label3.Text = "Radius Left [m]"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(68, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 14)
        Me.Label4.TabIndex = 100
        Me.Label4.Text = "Radius Right [m]"
        '
        'TextBox_Radius_right
        '
        Me.TextBox_Radius_right.Location = New System.Drawing.Point(3, 37)
        Me.TextBox_Radius_right.Name = "TextBox_Radius_right"
        Me.TextBox_Radius_right.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_Radius_right.TabIndex = 4
        Me.TextBox_Radius_right.Text = "500"
        '
        'Label_STRAIGHT_LEFT
        '
        Me.Label_STRAIGHT_LEFT.AutoSize = True
        Me.Label_STRAIGHT_LEFT.Location = New System.Drawing.Point(126, 26)
        Me.Label_STRAIGHT_LEFT.Name = "Label_STRAIGHT_LEFT"
        Me.Label_STRAIGHT_LEFT.Size = New System.Drawing.Size(115, 14)
        Me.Label_STRAIGHT_LEFT.TabIndex = 100
        Me.Label_STRAIGHT_LEFT.Text = "Straight portion [m]"
        Me.Label_STRAIGHT_LEFT.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TextBox_angle_left)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.TextBox_radius_Left)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(5, 65)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(175, 67)
        Me.Panel1.TabIndex = 101
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(89, 14)
        Me.Label7.TabIndex = 100
        Me.Label7.Text = "Angle Left [DD]"
        '
        'TextBox_3d_angle_left
        '
        Me.TextBox_3d_angle_left.Location = New System.Drawing.Point(126, 9)
        Me.TextBox_3d_angle_left.Name = "TextBox_3d_angle_left"
        Me.TextBox_3d_angle_left.Size = New System.Drawing.Size(215, 20)
        Me.TextBox_3d_angle_left.TabIndex = 1
        Me.TextBox_3d_angle_left.Text = "0"
        '
        'Label_3d_angle_left
        '
        Me.Label_3d_angle_left.AutoSize = True
        Me.Label_3d_angle_left.Location = New System.Drawing.Point(5, 13)
        Me.Label_3d_angle_left.Name = "Label_3d_angle_left"
        Me.Label_3d_angle_left.Size = New System.Drawing.Size(105, 14)
        Me.Label_3d_angle_left.TabIndex = 100
        Me.Label_3d_angle_left.Text = "3D Angle Left [DD]"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_angle_right)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.TextBox_Radius_right)
        Me.Panel2.Location = New System.Drawing.Point(186, 65)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(175, 67)
        Me.Panel2.TabIndex = 101
        '
        'TextBox_3d_angle_right
        '
        Me.TextBox_3d_angle_right.Location = New System.Drawing.Point(8, 46)
        Me.TextBox_3d_angle_right.Name = "TextBox_3d_angle_right"
        Me.TextBox_3d_angle_right.Size = New System.Drawing.Size(205, 20)
        Me.TextBox_3d_angle_right.TabIndex = 2
        Me.TextBox_3d_angle_right.Text = "0"
        '
        'Label_3d_angle_right
        '
        Me.Label_3d_angle_right.AutoSize = True
        Me.Label_3d_angle_right.Location = New System.Drawing.Point(220, 50)
        Me.Label_3d_angle_right.Name = "Label_3d_angle_right"
        Me.Label_3d_angle_right.Size = New System.Drawing.Size(111, 14)
        Me.Label_3d_angle_right.TabIndex = 100
        Me.Label_3d_angle_right.Text = "3D Angle Right [DD]"
        '
        'Button_create_hdd
        '
        Me.Button_create_hdd.Location = New System.Drawing.Point(5, 297)
        Me.Button_create_hdd.Name = "Button_create_hdd"
        Me.Button_create_hdd.Size = New System.Drawing.Size(235, 39)
        Me.Button_create_hdd.TabIndex = 7
        Me.Button_create_hdd.Text = "Create"
        Me.Button_create_hdd.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.TextBox_depth)
        Me.Panel3.Location = New System.Drawing.Point(155, 3)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(72, 55)
        Me.Panel3.TabIndex = 103
        '
        'TextBox_depth
        '
        Me.TextBox_depth.Location = New System.Drawing.Point(4, 20)
        Me.TextBox_depth.Name = "TextBox_depth"
        Me.TextBox_depth.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_depth.TabIndex = 0
        Me.TextBox_depth.Text = "0"
        Me.TextBox_depth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckBox_pick_start_end
        '
        Me.CheckBox_pick_start_end.AutoSize = True
        Me.CheckBox_pick_start_end.Location = New System.Drawing.Point(5, 9)
        Me.CheckBox_pick_start_end.Name = "CheckBox_pick_start_end"
        Me.CheckBox_pick_start_end.Size = New System.Drawing.Size(102, 18)
        Me.CheckBox_pick_start_end.TabIndex = 104
        Me.CheckBox_pick_start_end.Text = "Pick Start-End"
        Me.CheckBox_pick_start_end.UseVisualStyleBackColor = True
        '
        'CheckBox_1DEC
        '
        Me.CheckBox_1DEC.AutoSize = True
        Me.CheckBox_1DEC.Location = New System.Drawing.Point(276, 318)
        Me.CheckBox_1DEC.Name = "CheckBox_1DEC"
        Me.CheckBox_1DEC.Size = New System.Drawing.Size(78, 18)
        Me.CheckBox_1DEC.TabIndex = 104
        Me.CheckBox_1DEC.Text = "1 Decimal"
        Me.CheckBox_1DEC.UseVisualStyleBackColor = True
        '
        'CheckBox_0DEC
        '
        Me.CheckBox_0DEC.AutoSize = True
        Me.CheckBox_0DEC.Checked = True
        Me.CheckBox_0DEC.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_0DEC.Location = New System.Drawing.Point(276, 297)
        Me.CheckBox_0DEC.Name = "CheckBox_0DEC"
        Me.CheckBox_0DEC.Size = New System.Drawing.Size(85, 18)
        Me.CheckBox_0DEC.TabIndex = 104
        Me.CheckBox_0DEC.Text = "0 Decimals"
        Me.CheckBox_0DEC.UseVisualStyleBackColor = True
        '
        'Panel_formating
        '
        Me.Panel_formating.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel_formating.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_formating.Controls.Add(Me.ComboBox_layer_hdd_cl)
        Me.Panel_formating.Controls.Add(Me.ComboBox_layer_text)
        Me.Panel_formating.Controls.Add(Me.ComboBox_text_styles)
        Me.Panel_formating.Controls.Add(Me.Label16)
        Me.Panel_formating.Controls.Add(Me.Label17)
        Me.Panel_formating.Controls.Add(Me.Label14)
        Me.Panel_formating.Location = New System.Drawing.Point(5, 192)
        Me.Panel_formating.Name = "Panel_formating"
        Me.Panel_formating.Size = New System.Drawing.Size(356, 99)
        Me.Panel_formating.TabIndex = 106
        '
        'ComboBox_layer_hdd_cl
        '
        Me.ComboBox_layer_hdd_cl.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer_hdd_cl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_hdd_cl.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer_hdd_cl.FormattingEnabled = True
        Me.ComboBox_layer_hdd_cl.Location = New System.Drawing.Point(167, 68)
        Me.ComboBox_layer_hdd_cl.Name = "ComboBox_layer_hdd_cl"
        Me.ComboBox_layer_hdd_cl.Size = New System.Drawing.Size(174, 22)
        Me.ComboBox_layer_hdd_cl.TabIndex = 103
        '
        'ComboBox_layer_text
        '
        Me.ComboBox_layer_text.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer_text.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_text.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer_text.FormattingEnabled = True
        Me.ComboBox_layer_text.Location = New System.Drawing.Point(167, 35)
        Me.ComboBox_layer_text.Name = "ComboBox_layer_text"
        Me.ComboBox_layer_text.Size = New System.Drawing.Size(174, 22)
        Me.ComboBox_layer_text.TabIndex = 102
        '
        'ComboBox_text_styles
        '
        Me.ComboBox_text_styles.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_text_styles.FormattingEnabled = True
        Me.ComboBox_text_styles.Location = New System.Drawing.Point(167, 3)
        Me.ComboBox_text_styles.Name = "ComboBox_text_styles"
        Me.ComboBox_text_styles.Size = New System.Drawing.Size(174, 22)
        Me.ComboBox_text_styles.TabIndex = 101
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label16.Location = New System.Drawing.Point(8, 38)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 14)
        Me.Label16.TabIndex = 100
        Me.Label16.Text = "Layer Text"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label17.Location = New System.Drawing.Point(5, 71)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(121, 14)
        Me.Label17.TabIndex = 100
        Me.Label17.Text = "Layer HDD centerline"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label14.Location = New System.Drawing.Point(5, 6)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 14)
        Me.Label14.TabIndex = 100
        Me.Label14.Text = "Text Style"
        '
        'Button_3d_curve
        '
        Me.Button_3d_curve.Location = New System.Drawing.Point(5, 476)
        Me.Button_3d_curve.Name = "Button_3d_curve"
        Me.Button_3d_curve.Size = New System.Drawing.Size(356, 29)
        Me.Button_3d_curve.TabIndex = 7
        Me.Button_3d_curve.Text = "Draft 3D compound curve"
        Me.Button_3d_curve.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Silver
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.TextBox_3d_angle_left)
        Me.Panel4.Controls.Add(Me.Label_3d_angle_left)
        Me.Panel4.Controls.Add(Me.TextBox_3d_angle_right)
        Me.Panel4.Controls.Add(Me.Label_3d_angle_right)
        Me.Panel4.Location = New System.Drawing.Point(5, 383)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(356, 78)
        Me.Panel4.TabIndex = 106
        '
        'Panel_straight_portion
        '
        Me.Panel_straight_portion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_straight_portion.Controls.Add(Me.TextBox_straight_length)
        Me.Panel_straight_portion.Controls.Add(Me.Label_STRAIGHT_LEFT)
        Me.Panel_straight_portion.Location = New System.Drawing.Point(5, 138)
        Me.Panel_straight_portion.Name = "Panel_straight_portion"
        Me.Panel_straight_portion.Size = New System.Drawing.Size(356, 48)
        Me.Panel_straight_portion.TabIndex = 101
        '
        'HDD_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(367, 340)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel_formating)
        Me.Controls.Add(Me.CheckBox_0DEC)
        Me.Controls.Add(Me.CheckBox_1DEC)
        Me.Controls.Add(Me.CheckBox_pick_start_end)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Button_3d_curve)
        Me.Controls.Add(Me.Button_create_hdd)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel_straight_portion)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "HDD_Form"
        Me.Text = "HDD"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel_formating.ResumeLayout(False)
        Me.Panel_formating.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel_straight_portion.ResumeLayout(False)
        Me.Panel_straight_portion.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_angle_left As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_angle_right As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_radius_Left As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_straight_length As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Radius_right As System.Windows.Forms.TextBox
    Friend WithEvents Label_STRAIGHT_LEFT As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button_create_hdd As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_depth As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_pick_start_end As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_1DEC As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_0DEC As System.Windows.Forms.CheckBox
    Friend WithEvents Panel_formating As System.Windows.Forms.Panel
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox_3d_angle_left As System.Windows.Forms.TextBox
    Friend WithEvents Label_3d_angle_left As System.Windows.Forms.Label
    Friend WithEvents TextBox_3d_angle_right As System.Windows.Forms.TextBox
    Friend WithEvents Label_3d_angle_right As System.Windows.Forms.Label
    Friend WithEvents Button_3d_curve As System.Windows.Forms.Button
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents ComboBox_text_styles As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_layer_text As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layer_hdd_cl As System.Windows.Forms.ComboBox
    Friend WithEvents Panel_straight_portion As System.Windows.Forms.Panel
End Class
