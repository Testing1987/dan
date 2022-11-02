<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WORKSPACE_FORM
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_WS_latime = New System.Windows.Forms.TextBox()
        Me.Button_draw_WS = New System.Windows.Forms.Button()
        Me.Label_lungime = New System.Windows.Forms.Label()
        Me.TextBox_WS_lungime = New System.Windows.Forms.TextBox()
        Me.CheckBox_Start_point = New System.Windows.Forms.CheckBox()
        Me.CheckBox_middle_point = New System.Windows.Forms.CheckBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.CheckBox_start_end = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ComboBox_units = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ComboBox_drawing_units = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.CheckBox_buffer = New System.Windows.Forms.CheckBox()
        Me.Panel_buffer = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_buffer = New System.Windows.Forms.TextBox()
        Me.Panel_options_for_start_end = New System.Windows.Forms.Panel()
        Me.CheckBox_measure_on_middle = New System.Windows.Forms.CheckBox()
        Me.CheckBox_measure_Bottom = New System.Windows.Forms.CheckBox()
        Me.CheckBox_select_two_crossing_objects = New System.Windows.Forms.CheckBox()
        Me.CheckBox_select_1crossing_object = New System.Windows.Forms.CheckBox()
        Me.Panel4.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel_buffer.SuspendLayout()
        Me.Panel_options_for_start_end.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(27, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 15)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Width"
        '
        'TextBox_WS_latime
        '
        Me.TextBox_WS_latime.Location = New System.Drawing.Point(15, 25)
        Me.TextBox_WS_latime.Name = "TextBox_WS_latime"
        Me.TextBox_WS_latime.Size = New System.Drawing.Size(73, 23)
        Me.TextBox_WS_latime.TabIndex = 1
        '
        'Button_draw_WS
        '
        Me.Button_draw_WS.BackColor = System.Drawing.Color.Green
        Me.Button_draw_WS.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_draw_WS.ForeColor = System.Drawing.Color.White
        Me.Button_draw_WS.Location = New System.Drawing.Point(14, 426)
        Me.Button_draw_WS.Name = "Button_draw_WS"
        Me.Button_draw_WS.Size = New System.Drawing.Size(206, 43)
        Me.Button_draw_WS.TabIndex = 20
        Me.Button_draw_WS.Text = "Draw"
        Me.Button_draw_WS.UseVisualStyleBackColor = False
        '
        'Label_lungime
        '
        Me.Label_lungime.AutoSize = True
        Me.Label_lungime.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label_lungime.Location = New System.Drawing.Point(117, 7)
        Me.Label_lungime.Name = "Label_lungime"
        Me.Label_lungime.Size = New System.Drawing.Size(46, 15)
        Me.Label_lungime.TabIndex = 100
        Me.Label_lungime.Text = "Length"
        '
        'TextBox_WS_lungime
        '
        Me.TextBox_WS_lungime.Location = New System.Drawing.Point(111, 25)
        Me.TextBox_WS_lungime.Name = "TextBox_WS_lungime"
        Me.TextBox_WS_lungime.Size = New System.Drawing.Size(73, 23)
        Me.TextBox_WS_lungime.TabIndex = 2
        '
        'CheckBox_Start_point
        '
        Me.CheckBox_Start_point.AutoSize = True
        Me.CheckBox_Start_point.Checked = True
        Me.CheckBox_Start_point.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_Start_point.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_Start_point.Location = New System.Drawing.Point(3, 4)
        Me.CheckBox_Start_point.Name = "CheckBox_Start_point"
        Me.CheckBox_Start_point.Size = New System.Drawing.Size(86, 19)
        Me.CheckBox_Start_point.TabIndex = 101
        Me.CheckBox_Start_point.Text = "Start Point"
        Me.CheckBox_Start_point.UseVisualStyleBackColor = True
        '
        'CheckBox_middle_point
        '
        Me.CheckBox_middle_point.AutoSize = True
        Me.CheckBox_middle_point.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_middle_point.Location = New System.Drawing.Point(3, 33)
        Me.CheckBox_middle_point.Name = "CheckBox_middle_point"
        Me.CheckBox_middle_point.Size = New System.Drawing.Size(95, 19)
        Me.CheckBox_middle_point.TabIndex = 101
        Me.CheckBox_middle_point.Text = "Middle Point"
        Me.CheckBox_middle_point.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.LightGray
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.CheckBox_buffer)
        Me.Panel4.Controls.Add(Me.CheckBox_middle_point)
        Me.Panel4.Controls.Add(Me.CheckBox_Start_point)
        Me.Panel4.Controls.Add(Me.CheckBox_start_end)
        Me.Panel4.Location = New System.Drawing.Point(14, 185)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(106, 119)
        Me.Panel4.TabIndex = 102
        '
        'CheckBox_start_end
        '
        Me.CheckBox_start_end.AutoSize = True
        Me.CheckBox_start_end.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_start_end.Location = New System.Drawing.Point(2, 61)
        Me.CheckBox_start_end.Name = "CheckBox_start_end"
        Me.CheckBox_start_end.Size = New System.Drawing.Size(85, 19)
        Me.CheckBox_start_end.TabIndex = 101
        Me.CheckBox_start_end.Text = "Start - End"
        Me.CheckBox_start_end.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightGray
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.ComboBox_units)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label_lungime)
        Me.Panel1.Controls.Add(Me.TextBox_WS_latime)
        Me.Panel1.Controls.Add(Me.TextBox_WS_lungime)
        Me.Panel1.Location = New System.Drawing.Point(14, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(206, 100)
        Me.Panel1.TabIndex = 103
        '
        'ComboBox_units
        '
        Me.ComboBox_units.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_units.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_units.FormattingEnabled = True
        Me.ComboBox_units.Items.AddRange(New Object() {"Foot", "Meter"})
        Me.ComboBox_units.Location = New System.Drawing.Point(15, 54)
        Me.ComboBox_units.Name = "ComboBox_units"
        Me.ComboBox_units.Size = New System.Drawing.Size(169, 24)
        Me.ComboBox_units.TabIndex = 104
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.LightGray
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.ComboBox_drawing_units)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Location = New System.Drawing.Point(14, 118)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(206, 61)
        Me.Panel2.TabIndex = 103
        '
        'ComboBox_drawing_units
        '
        Me.ComboBox_drawing_units.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_drawing_units.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_drawing_units.FormattingEnabled = True
        Me.ComboBox_drawing_units.Items.AddRange(New Object() {"Foot", "Meter"})
        Me.ComboBox_drawing_units.Location = New System.Drawing.Point(6, 24)
        Me.ComboBox_drawing_units.Name = "ComboBox_drawing_units"
        Me.ComboBox_drawing_units.Size = New System.Drawing.Size(181, 24)
        Me.ComboBox_drawing_units.TabIndex = 104
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(3, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 15)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Drawing units"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.PaleGreen
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Controls.Add(Me.Panel_buffer)
        Me.Panel3.Controls.Add(Me.Button_draw_WS)
        Me.Panel3.Controls.Add(Me.Panel_options_for_start_end)
        Me.Panel3.Controls.Add(Me.Panel4)
        Me.Panel3.Location = New System.Drawing.Point(12, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(234, 482)
        Me.Panel3.TabIndex = 102
        '
        'CheckBox_buffer
        '
        Me.CheckBox_buffer.AutoSize = True
        Me.CheckBox_buffer.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_buffer.Location = New System.Drawing.Point(2, 89)
        Me.CheckBox_buffer.Name = "CheckBox_buffer"
        Me.CheckBox_buffer.Size = New System.Drawing.Size(86, 19)
        Me.CheckBox_buffer.TabIndex = 101
        Me.CheckBox_buffer.Text = "Use Buffer"
        Me.CheckBox_buffer.UseVisualStyleBackColor = True
        '
        'Panel_buffer
        '
        Me.Panel_buffer.BackColor = System.Drawing.Color.LightGray
        Me.Panel_buffer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_buffer.Controls.Add(Me.Label3)
        Me.Panel_buffer.Controls.Add(Me.TextBox_buffer)
        Me.Panel_buffer.Location = New System.Drawing.Point(126, 185)
        Me.Panel_buffer.Name = "Panel_buffer"
        Me.Panel_buffer.Size = New System.Drawing.Size(94, 119)
        Me.Panel_buffer.TabIndex = 102
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(3, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 15)
        Me.Label3.TabIndex = 100
        Me.Label3.Text = "Buffer"
        '
        'TextBox_buffer
        '
        Me.TextBox_buffer.Location = New System.Drawing.Point(17, 82)
        Me.TextBox_buffer.Name = "TextBox_buffer"
        Me.TextBox_buffer.Size = New System.Drawing.Size(54, 23)
        Me.TextBox_buffer.TabIndex = 1
        Me.TextBox_buffer.Text = "50"
        Me.TextBox_buffer.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel_options_for_start_end
        '
        Me.Panel_options_for_start_end.BackColor = System.Drawing.Color.LightGray
        Me.Panel_options_for_start_end.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_options_for_start_end.Controls.Add(Me.CheckBox_measure_on_middle)
        Me.Panel_options_for_start_end.Controls.Add(Me.CheckBox_measure_Bottom)
        Me.Panel_options_for_start_end.Controls.Add(Me.CheckBox_select_two_crossing_objects)
        Me.Panel_options_for_start_end.Controls.Add(Me.CheckBox_select_1crossing_object)
        Me.Panel_options_for_start_end.Location = New System.Drawing.Point(14, 310)
        Me.Panel_options_for_start_end.Name = "Panel_options_for_start_end"
        Me.Panel_options_for_start_end.Size = New System.Drawing.Size(206, 110)
        Me.Panel_options_for_start_end.TabIndex = 102
        '
        'CheckBox_measure_on_middle
        '
        Me.CheckBox_measure_on_middle.AutoSize = True
        Me.CheckBox_measure_on_middle.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_measure_on_middle.Location = New System.Drawing.Point(3, 30)
        Me.CheckBox_measure_on_middle.Name = "CheckBox_measure_on_middle"
        Me.CheckBox_measure_on_middle.Size = New System.Drawing.Size(133, 19)
        Me.CheckBox_measure_on_middle.TabIndex = 101
        Me.CheckBox_measure_on_middle.Text = "Measure on Middle"
        Me.CheckBox_measure_on_middle.UseVisualStyleBackColor = True
        '
        'CheckBox_measure_Bottom
        '
        Me.CheckBox_measure_Bottom.AutoSize = True
        Me.CheckBox_measure_Bottom.Checked = True
        Me.CheckBox_measure_Bottom.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_measure_Bottom.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_measure_Bottom.Location = New System.Drawing.Point(3, 4)
        Me.CheckBox_measure_Bottom.Name = "CheckBox_measure_Bottom"
        Me.CheckBox_measure_Bottom.Size = New System.Drawing.Size(137, 19)
        Me.CheckBox_measure_Bottom.TabIndex = 101
        Me.CheckBox_measure_Bottom.Text = "Measure on Bottom"
        Me.CheckBox_measure_Bottom.UseVisualStyleBackColor = True
        '
        'CheckBox_select_two_crossing_objects
        '
        Me.CheckBox_select_two_crossing_objects.AutoSize = True
        Me.CheckBox_select_two_crossing_objects.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_select_two_crossing_objects.Location = New System.Drawing.Point(2, 81)
        Me.CheckBox_select_two_crossing_objects.Name = "CheckBox_select_two_crossing_objects"
        Me.CheckBox_select_two_crossing_objects.Size = New System.Drawing.Size(181, 19)
        Me.CheckBox_select_two_crossing_objects.TabIndex = 101
        Me.CheckBox_select_two_crossing_objects.Text = "Select (2) Crossing Objects"
        Me.CheckBox_select_two_crossing_objects.UseVisualStyleBackColor = True
        '
        'CheckBox_select_1crossing_object
        '
        Me.CheckBox_select_1crossing_object.AutoSize = True
        Me.CheckBox_select_1crossing_object.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_select_1crossing_object.Location = New System.Drawing.Point(2, 56)
        Me.CheckBox_select_1crossing_object.Name = "CheckBox_select_1crossing_object"
        Me.CheckBox_select_1crossing_object.Size = New System.Drawing.Size(174, 19)
        Me.CheckBox_select_1crossing_object.TabIndex = 101
        Me.CheckBox_select_1crossing_object.Text = "Select (1) Crossing Object"
        Me.CheckBox_select_1crossing_object.UseVisualStyleBackColor = True
        '
        'WORKSPACE_FORM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(259, 504)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "WORKSPACE_FORM"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "PPL Work Space"
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel_buffer.ResumeLayout(False)
        Me.Panel_buffer.PerformLayout()
        Me.Panel_options_for_start_end.ResumeLayout(False)
        Me.Panel_options_for_start_end.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_WS_latime As System.Windows.Forms.TextBox
    Friend WithEvents Button_draw_WS As System.Windows.Forms.Button
    Friend WithEvents Label_lungime As System.Windows.Forms.Label
    Friend WithEvents TextBox_WS_lungime As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_Start_point As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_middle_point As System.Windows.Forms.CheckBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_start_end As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ComboBox_units As System.Windows.Forms.ComboBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ComboBox_drawing_units As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel_options_for_start_end As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_measure_on_middle As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_measure_Bottom As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_select_1crossing_object As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_select_two_crossing_objects As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_buffer As System.Windows.Forms.CheckBox
    Friend WithEvents Panel_buffer As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_buffer As System.Windows.Forms.TextBox
End Class
