<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Alignment_w2XL_Form
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
        Me.TextBox_ROW_START_XL = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button_PICK_Water = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button_pick_pipe = New System.Windows.Forms.Button()
        Me.Button_PICK_road = New System.Windows.Forms.Button()
        Me.Button_pick_rail = New System.Windows.Forms.Button()
        Me.Button_pick_power = New System.Windows.Forms.Button()
        Me.Button_SA_start = New System.Windows.Forms.Button()
        Me.Button_SA_end = New System.Windows.Forms.Button()
        Me.Button_facility = New System.Windows.Forms.Button()
        Me.Button_ELBOW = New System.Windows.Forms.Button()
        Me.Button_COROSION = New System.Windows.Forms.Button()
        Me.Button_matchline = New System.Windows.Forms.Button()
        Me.Button_transition = New System.Windows.Forms.Button()
        Me.Button_CABLE = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_RW_start = New System.Windows.Forms.Button()
        Me.Button_RW_end = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button_SBW_start = New System.Windows.Forms.Button()
        Me.Button_SBW_End = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_ROW_START_XL
        '
        Me.TextBox_ROW_START_XL.BackColor = System.Drawing.Color.White
        Me.TextBox_ROW_START_XL.ForeColor = System.Drawing.Color.Black
        Me.TextBox_ROW_START_XL.Location = New System.Drawing.Point(84, 9)
        Me.TextBox_ROW_START_XL.Name = "TextBox_ROW_START_XL"
        Me.TextBox_ROW_START_XL.Size = New System.Drawing.Size(53, 21)
        Me.TextBox_ROW_START_XL.TabIndex = 124
        Me.TextBox_ROW_START_XL.Text = "1"
        Me.TextBox_ROW_START_XL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 15)
        Me.Label1.TabIndex = 125
        Me.Label1.Text = "Excel Row"
        '
        'Button_PICK_Water
        '
        Me.Button_PICK_Water.BackColor = System.Drawing.Color.Aqua
        Me.Button_PICK_Water.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_PICK_Water.ForeColor = System.Drawing.Color.Black
        Me.Button_PICK_Water.Location = New System.Drawing.Point(115, 56)
        Me.Button_PICK_Water.Name = "Button_PICK_Water"
        Me.Button_PICK_Water.Size = New System.Drawing.Size(128, 31)
        Me.Button_PICK_Water.TabIndex = 0
        Me.Button_PICK_Water.Text = "Water Crossing"
        Me.Button_PICK_Water.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Silver
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(249, 9)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(46, 303)
        Me.Panel1.TabIndex = 126
        '
        'Button_pick_pipe
        '
        Me.Button_pick_pipe.BackColor = System.Drawing.Color.Wheat
        Me.Button_pick_pipe.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick_pipe.ForeColor = System.Drawing.Color.Black
        Me.Button_pick_pipe.Location = New System.Drawing.Point(115, 93)
        Me.Button_pick_pipe.Name = "Button_pick_pipe"
        Me.Button_pick_pipe.Size = New System.Drawing.Size(128, 31)
        Me.Button_pick_pipe.TabIndex = 0
        Me.Button_pick_pipe.Text = "Pipeline Crossing"
        Me.Button_pick_pipe.UseVisualStyleBackColor = False
        '
        'Button_PICK_road
        '
        Me.Button_PICK_road.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Button_PICK_road.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_PICK_road.ForeColor = System.Drawing.Color.Black
        Me.Button_PICK_road.Location = New System.Drawing.Point(115, 167)
        Me.Button_PICK_road.Name = "Button_PICK_road"
        Me.Button_PICK_road.Size = New System.Drawing.Size(128, 31)
        Me.Button_PICK_road.TabIndex = 0
        Me.Button_PICK_road.Text = "Road Crossing"
        Me.Button_PICK_road.UseVisualStyleBackColor = False
        '
        'Button_pick_rail
        '
        Me.Button_pick_rail.BackColor = System.Drawing.Color.Tomato
        Me.Button_pick_rail.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick_rail.ForeColor = System.Drawing.Color.Black
        Me.Button_pick_rail.Location = New System.Drawing.Point(115, 204)
        Me.Button_pick_rail.Name = "Button_pick_rail"
        Me.Button_pick_rail.Size = New System.Drawing.Size(128, 31)
        Me.Button_pick_rail.TabIndex = 0
        Me.Button_pick_rail.Text = "Rail Crossing"
        Me.Button_pick_rail.UseVisualStyleBackColor = False
        '
        'Button_pick_power
        '
        Me.Button_pick_power.BackColor = System.Drawing.Color.PaleGreen
        Me.Button_pick_power.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick_power.ForeColor = System.Drawing.Color.Black
        Me.Button_pick_power.Location = New System.Drawing.Point(115, 241)
        Me.Button_pick_power.Name = "Button_pick_power"
        Me.Button_pick_power.Size = New System.Drawing.Size(128, 31)
        Me.Button_pick_power.TabIndex = 0
        Me.Button_pick_power.Text = "Powerline Crossing"
        Me.Button_pick_power.UseVisualStyleBackColor = False
        '
        'Button_SA_start
        '
        Me.Button_SA_start.BackColor = System.Drawing.Color.PowderBlue
        Me.Button_SA_start.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_SA_start.ForeColor = System.Drawing.Color.Black
        Me.Button_SA_start.Location = New System.Drawing.Point(3, 4)
        Me.Button_SA_start.Name = "Button_SA_start"
        Me.Button_SA_start.Size = New System.Drawing.Size(181, 31)
        Me.Button_SA_start.TabIndex = 0
        Me.Button_SA_start.Text = "Screw Anchor Start"
        Me.Button_SA_start.UseVisualStyleBackColor = False
        '
        'Button_SA_end
        '
        Me.Button_SA_end.BackColor = System.Drawing.Color.PowderBlue
        Me.Button_SA_end.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_SA_end.ForeColor = System.Drawing.Color.Black
        Me.Button_SA_end.Location = New System.Drawing.Point(3, 41)
        Me.Button_SA_end.Name = "Button_SA_end"
        Me.Button_SA_end.Size = New System.Drawing.Size(181, 31)
        Me.Button_SA_end.TabIndex = 0
        Me.Button_SA_end.Text = "Screw Anchor End"
        Me.Button_SA_end.UseVisualStyleBackColor = False
        '
        'Button_facility
        '
        Me.Button_facility.BackColor = System.Drawing.Color.Coral
        Me.Button_facility.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_facility.ForeColor = System.Drawing.Color.Black
        Me.Button_facility.Location = New System.Drawing.Point(301, 281)
        Me.Button_facility.Name = "Button_facility"
        Me.Button_facility.Size = New System.Drawing.Size(199, 31)
        Me.Button_facility.TabIndex = 0
        Me.Button_facility.Text = "Facility"
        Me.Button_facility.UseVisualStyleBackColor = False
        '
        'Button_ELBOW
        '
        Me.Button_ELBOW.BackColor = System.Drawing.Color.Thistle
        Me.Button_ELBOW.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_ELBOW.ForeColor = System.Drawing.Color.Black
        Me.Button_ELBOW.Location = New System.Drawing.Point(115, 278)
        Me.Button_ELBOW.Name = "Button_ELBOW"
        Me.Button_ELBOW.Size = New System.Drawing.Size(128, 31)
        Me.Button_ELBOW.TabIndex = 0
        Me.Button_ELBOW.Text = "Elbow"
        Me.Button_ELBOW.UseVisualStyleBackColor = False
        '
        'Button_COROSION
        '
        Me.Button_COROSION.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button_COROSION.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_COROSION.ForeColor = System.Drawing.Color.Black
        Me.Button_COROSION.Location = New System.Drawing.Point(301, 183)
        Me.Button_COROSION.Name = "Button_COROSION"
        Me.Button_COROSION.Size = New System.Drawing.Size(199, 31)
        Me.Button_COROSION.TabIndex = 0
        Me.Button_COROSION.Text = "Cathodic protection"
        Me.Button_COROSION.UseVisualStyleBackColor = False
        '
        'Button_matchline
        '
        Me.Button_matchline.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Button_matchline.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_matchline.ForeColor = System.Drawing.Color.Black
        Me.Button_matchline.Location = New System.Drawing.Point(12, 281)
        Me.Button_matchline.Name = "Button_matchline"
        Me.Button_matchline.Size = New System.Drawing.Size(90, 31)
        Me.Button_matchline.TabIndex = 0
        Me.Button_matchline.Text = "Matchline"
        Me.Button_matchline.UseVisualStyleBackColor = False
        '
        'Button_transition
        '
        Me.Button_transition.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Button_transition.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_transition.ForeColor = System.Drawing.Color.Black
        Me.Button_transition.Location = New System.Drawing.Point(12, 56)
        Me.Button_transition.Name = "Button_transition"
        Me.Button_transition.Size = New System.Drawing.Size(90, 43)
        Me.Button_transition.TabIndex = 0
        Me.Button_transition.Text = "Transition"
        Me.Button_transition.UseVisualStyleBackColor = False
        '
        'Button_CABLE
        '
        Me.Button_CABLE.BackColor = System.Drawing.Color.Khaki
        Me.Button_CABLE.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_CABLE.ForeColor = System.Drawing.Color.Black
        Me.Button_CABLE.Location = New System.Drawing.Point(115, 130)
        Me.Button_CABLE.Name = "Button_CABLE"
        Me.Button_CABLE.Size = New System.Drawing.Size(128, 31)
        Me.Button_CABLE.TabIndex = 0
        Me.Button_CABLE.Text = "Cable Crossing"
        Me.Button_CABLE.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Silver
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Button_SA_start)
        Me.Panel2.Controls.Add(Me.Button_SA_end)
        Me.Panel2.Location = New System.Drawing.Point(305, 9)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(195, 81)
        Me.Panel2.TabIndex = 126
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Button_RW_start)
        Me.Panel3.Controls.Add(Me.Button_RW_end)
        Me.Panel3.Location = New System.Drawing.Point(304, 96)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(91, 81)
        Me.Panel3.TabIndex = 126
        '
        'Button_RW_start
        '
        Me.Button_RW_start.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Button_RW_start.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_RW_start.ForeColor = System.Drawing.Color.Black
        Me.Button_RW_start.Location = New System.Drawing.Point(3, 4)
        Me.Button_RW_start.Name = "Button_RW_start"
        Me.Button_RW_start.Size = New System.Drawing.Size(79, 31)
        Me.Button_RW_start.TabIndex = 0
        Me.Button_RW_start.Text = "RW Start"
        Me.Button_RW_start.UseVisualStyleBackColor = False
        '
        'Button_RW_end
        '
        Me.Button_RW_end.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Button_RW_end.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_RW_end.ForeColor = System.Drawing.Color.Black
        Me.Button_RW_end.Location = New System.Drawing.Point(3, 41)
        Me.Button_RW_end.Name = "Button_RW_end"
        Me.Button_RW_end.Size = New System.Drawing.Size(79, 31)
        Me.Button_RW_end.TabIndex = 0
        Me.Button_RW_end.Text = "RW End"
        Me.Button_RW_end.UseVisualStyleBackColor = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Silver
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Button_SBW_start)
        Me.Panel4.Controls.Add(Me.Button_SBW_End)
        Me.Panel4.Location = New System.Drawing.Point(401, 96)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(99, 81)
        Me.Panel4.TabIndex = 126
        '
        'Button_SBW_start
        '
        Me.Button_SBW_start.BackColor = System.Drawing.Color.Aqua
        Me.Button_SBW_start.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_SBW_start.ForeColor = System.Drawing.Color.Black
        Me.Button_SBW_start.Location = New System.Drawing.Point(3, 4)
        Me.Button_SBW_start.Name = "Button_SBW_start"
        Me.Button_SBW_start.Size = New System.Drawing.Size(85, 31)
        Me.Button_SBW_start.TabIndex = 0
        Me.Button_SBW_start.Text = "SBW Start"
        Me.Button_SBW_start.UseVisualStyleBackColor = False
        '
        'Button_SBW_End
        '
        Me.Button_SBW_End.BackColor = System.Drawing.Color.Aqua
        Me.Button_SBW_End.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button_SBW_End.ForeColor = System.Drawing.Color.Black
        Me.Button_SBW_End.Location = New System.Drawing.Point(3, 41)
        Me.Button_SBW_End.Name = "Button_SBW_End"
        Me.Button_SBW_End.Size = New System.Drawing.Size(85, 31)
        Me.Button_SBW_End.TabIndex = 0
        Me.Button_SBW_End.Text = "SBW End"
        Me.Button_SBW_End.UseVisualStyleBackColor = False
        '
        'Alignment_w2XL_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(511, 322)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_ROW_START_XL)
        Me.Controls.Add(Me.Button_CABLE)
        Me.Controls.Add(Me.Button_pick_pipe)
        Me.Controls.Add(Me.Button_ELBOW)
        Me.Controls.Add(Me.Button_facility)
        Me.Controls.Add(Me.Button_pick_power)
        Me.Controls.Add(Me.Button_pick_rail)
        Me.Controls.Add(Me.Button_COROSION)
        Me.Controls.Add(Me.Button_PICK_road)
        Me.Controls.Add(Me.Button_PICK_Water)
        Me.Controls.Add(Me.Button_matchline)
        Me.Controls.Add(Me.Button_transition)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Alignment_w2XL_Form"
        Me.Text = "Alignment to Excel"
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_ROW_START_XL As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_PICK_Water As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button_pick_pipe As System.Windows.Forms.Button
    Friend WithEvents Button_PICK_road As System.Windows.Forms.Button
    Friend WithEvents Button_pick_rail As System.Windows.Forms.Button
    Friend WithEvents Button_pick_power As System.Windows.Forms.Button
    Friend WithEvents Button_SA_start As System.Windows.Forms.Button
    Friend WithEvents Button_SA_end As System.Windows.Forms.Button
    Friend WithEvents Button_facility As System.Windows.Forms.Button
    Friend WithEvents Button_ELBOW As System.Windows.Forms.Button
    Friend WithEvents Button_COROSION As System.Windows.Forms.Button
    Friend WithEvents Button_matchline As System.Windows.Forms.Button
    Friend WithEvents Button_transition As System.Windows.Forms.Button
    Friend WithEvents Button_CABLE As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button_RW_start As System.Windows.Forms.Button
    Friend WithEvents Button_RW_end As System.Windows.Forms.Button
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Button_SBW_start As System.Windows.Forms.Button
    Friend WithEvents Button_SBW_End As System.Windows.Forms.Button
End Class
