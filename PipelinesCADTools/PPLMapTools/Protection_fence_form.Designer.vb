<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Protection_fence_form
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
        Me.TextBox_buffer_dist = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button_draw = New System.Windows.Forms.Button()
        Me.TextBox_max_dist = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_offset = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CheckBox_MP = New System.Windows.Forms.CheckBox()
        Me.CheckBox_sta = New System.Windows.Forms.CheckBox()
        Me.CheckBox_lat_long = New System.Windows.Forms.CheckBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_text_size = New System.Windows.Forms.TextBox()
        Me.TextBox_arrow_size = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_buffer_dist
        '
        Me.TextBox_buffer_dist.Location = New System.Drawing.Point(3, 20)
        Me.TextBox_buffer_dist.Name = "TextBox_buffer_dist"
        Me.TextBox_buffer_dist.Size = New System.Drawing.Size(100, 21)
        Me.TextBox_buffer_dist.TabIndex = 0
        Me.TextBox_buffer_dist.Text = "100"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 2)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Distance"
        '
        'Button_draw
        '
        Me.Button_draw.Location = New System.Drawing.Point(3, 95)
        Me.Button_draw.Name = "Button_draw"
        Me.Button_draw.Size = New System.Drawing.Size(272, 49)
        Me.Button_draw.TabIndex = 2
        Me.Button_draw.Text = "Draw"
        Me.Button_draw.UseVisualStyleBackColor = True
        '
        'TextBox_max_dist
        '
        Me.TextBox_max_dist.Location = New System.Drawing.Point(132, 68)
        Me.TextBox_max_dist.Name = "TextBox_max_dist"
        Me.TextBox_max_dist.Size = New System.Drawing.Size(143, 21)
        Me.TextBox_max_dist.TabIndex = 0
        Me.TextBox_max_dist.Text = "50"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(129, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(153, 60)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Maximum distance" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "between the structure" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and the Workspace" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "that requires safety " & _
    "fence"
        '
        'TextBox_offset
        '
        Me.TextBox_offset.Location = New System.Drawing.Point(3, 68)
        Me.TextBox_offset.Name = "TextBox_offset"
        Me.TextBox_offset.Size = New System.Drawing.Size(100, 21)
        Me.TextBox_offset.TabIndex = 0
        Me.TextBox_offset.Text = "2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(0, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 15)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Offset distance"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Button_draw)
        Me.Panel1.Controls.Add(Me.TextBox_buffer_dist)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.TextBox_max_dist)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.TextBox_offset)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(292, 171)
        Me.Panel1.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.CheckBox_MP)
        Me.Panel2.Controls.Add(Me.CheckBox_sta)
        Me.Panel2.Controls.Add(Me.CheckBox_lat_long)
        Me.Panel2.Location = New System.Drawing.Point(300, 90)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(129, 84)
        Me.Panel2.TabIndex = 4
        '
        'CheckBox_MP
        '
        Me.CheckBox_MP.AutoSize = True
        Me.CheckBox_MP.Location = New System.Drawing.Point(3, 53)
        Me.CheckBox_MP.Name = "CheckBox_MP"
        Me.CheckBox_MP.Size = New System.Drawing.Size(78, 19)
        Me.CheckBox_MP.TabIndex = 0
        Me.CheckBox_MP.Text = "Mile Post"
        Me.CheckBox_MP.UseVisualStyleBackColor = True
        '
        'CheckBox_sta
        '
        Me.CheckBox_sta.AutoSize = True
        Me.CheckBox_sta.Checked = True
        Me.CheckBox_sta.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_sta.Location = New System.Drawing.Point(3, 28)
        Me.CheckBox_sta.Name = "CheckBox_sta"
        Me.CheckBox_sta.Size = New System.Drawing.Size(66, 19)
        Me.CheckBox_sta.TabIndex = 0
        Me.CheckBox_sta.Text = "Station"
        Me.CheckBox_sta.UseVisualStyleBackColor = True
        '
        'CheckBox_lat_long
        '
        Me.CheckBox_lat_long.AutoSize = True
        Me.CheckBox_lat_long.Location = New System.Drawing.Point(3, 3)
        Me.CheckBox_lat_long.Name = "CheckBox_lat_long"
        Me.CheckBox_lat_long.Size = New System.Drawing.Size(76, 19)
        Me.CheckBox_lat_long.TabIndex = 0
        Me.CheckBox_lat_long.Text = "Lat-Long"
        Me.CheckBox_lat_long.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.TextBox_text_size)
        Me.Panel3.Controls.Add(Me.TextBox_arrow_size)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Location = New System.Drawing.Point(301, 3)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(128, 81)
        Me.Panel3.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 2)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 15)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Leader Format"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 15)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Arrow size"
        '
        'TextBox_text_size
        '
        Me.TextBox_text_size.Location = New System.Drawing.Point(77, 47)
        Me.TextBox_text_size.Name = "TextBox_text_size"
        Me.TextBox_text_size.Size = New System.Drawing.Size(39, 21)
        Me.TextBox_text_size.TabIndex = 0
        Me.TextBox_text_size.Text = "5"
        '
        'TextBox_arrow_size
        '
        Me.TextBox_arrow_size.Location = New System.Drawing.Point(77, 23)
        Me.TextBox_arrow_size.Name = "TextBox_arrow_size"
        Me.TextBox_arrow_size.Size = New System.Drawing.Size(39, 21)
        Me.TextBox_arrow_size.TabIndex = 0
        Me.TextBox_arrow_size.Text = "5"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 50)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 15)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Text size"
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Panel1)
        Me.Panel4.Controls.Add(Me.Panel2)
        Me.Panel4.Controls.Add(Me.Panel3)
        Me.Panel4.Location = New System.Drawing.Point(12, 12)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(437, 183)
        Me.Panel4.TabIndex = 5
        '
        'Protection_fence_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(496, 223)
        Me.Controls.Add(Me.Panel4)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Protection_fence_form"
        Me.Text = "Protection fence"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TextBox_buffer_dist As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_draw As System.Windows.Forms.Button
    Friend WithEvents TextBox_max_dist As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_offset As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_MP As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_sta As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_lat_long As System.Windows.Forms.CheckBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_text_size As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_arrow_size As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
End Class
