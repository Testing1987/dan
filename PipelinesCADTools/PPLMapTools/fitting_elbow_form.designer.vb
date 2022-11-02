<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fitting_elbow_form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(fitting_elbow_form))
        Me.CheckBox_draw_od = New System.Windows.Forms.CheckBox()
        Me.ComboBox_nps = New System.Windows.Forms.ComboBox()
        Me.TextBox_diameter_X = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label_layer = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox_report = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_tangent_length = New System.Windows.Forms.TextBox()
        Me.TextBox_elbow_angle = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_radius1 = New System.Windows.Forms.TextBox()
        Me.Button_fitting = New System.Windows.Forms.Button()
        Me.ComboBox_layer_OD = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layer_CL = New System.Windows.Forms.ComboBox()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'CheckBox_draw_od
        '
        Me.CheckBox_draw_od.AutoSize = True
        Me.CheckBox_draw_od.Location = New System.Drawing.Point(309, 321)
        Me.CheckBox_draw_od.Name = "CheckBox_draw_od"
        Me.CheckBox_draw_od.Size = New System.Drawing.Size(72, 18)
        Me.CheckBox_draw_od.TabIndex = 18
        Me.CheckBox_draw_od.Text = "Draw OD"
        Me.CheckBox_draw_od.UseVisualStyleBackColor = True
        '
        'ComboBox_nps
        '
        Me.ComboBox_nps.BackColor = System.Drawing.Color.White
        Me.ComboBox_nps.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_nps.FormattingEnabled = True
        Me.ComboBox_nps.Location = New System.Drawing.Point(99, 317)
        Me.ComboBox_nps.Name = "ComboBox_nps"
        Me.ComboBox_nps.Size = New System.Drawing.Size(103, 22)
        Me.ComboBox_nps.TabIndex = 17
        '
        'TextBox_diameter_X
        '
        Me.TextBox_diameter_X.Location = New System.Drawing.Point(29, 318)
        Me.TextBox_diameter_X.Name = "TextBox_diameter_X"
        Me.TextBox_diameter_X.Size = New System.Drawing.Size(32, 20)
        Me.TextBox_diameter_X.TabIndex = 14
        Me.TextBox_diameter_X.Text = "3"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(294, 269)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 14)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Layer OD"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(59, 319)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 19)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "D x"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_layer
        '
        Me.Label_layer.AutoSize = True
        Me.Label_layer.Location = New System.Drawing.Point(50, 269)
        Me.Label_layer.Name = "Label_layer"
        Me.Label_layer.Size = New System.Drawing.Size(56, 14)
        Me.Label_layer.TabIndex = 12
        Me.Label_layer.Text = "Layer CL"
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_report)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.TextBox_tangent_length)
        Me.Panel2.Controls.Add(Me.TextBox_elbow_angle)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.TextBox_radius1)
        Me.Panel2.Location = New System.Drawing.Point(12, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(377, 248)
        Me.Panel2.TabIndex = 13
        '
        'TextBox_report
        '
        Me.TextBox_report.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_report.ForeColor = System.Drawing.Color.Black
        Me.TextBox_report.Location = New System.Drawing.Point(202, 193)
        Me.TextBox_report.Multiline = True
        Me.TextBox_report.Name = "TextBox_report"
        Me.TextBox_report.ReadOnly = True
        Me.TextBox_report.Size = New System.Drawing.Size(173, 53)
        Me.TextBox_report.TabIndex = 22
        Me.TextBox_report.Text = "line1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "line2" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "line3"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(72, 36)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(139, 14)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Straight Tangent Length"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_tangent_length
        '
        Me.TextBox_tangent_length.Location = New System.Drawing.Point(217, 33)
        Me.TextBox_tangent_length.Name = "TextBox_tangent_length"
        Me.TextBox_tangent_length.Size = New System.Drawing.Size(54, 20)
        Me.TextBox_tangent_length.TabIndex = 20
        Me.TextBox_tangent_length.Text = "1"
        '
        'TextBox_elbow_angle
        '
        Me.TextBox_elbow_angle.Location = New System.Drawing.Point(25, 136)
        Me.TextBox_elbow_angle.Name = "TextBox_elbow_angle"
        Me.TextBox_elbow_angle.Size = New System.Drawing.Size(54, 20)
        Me.TextBox_elbow_angle.TabIndex = 19
        Me.TextBox_elbow_angle.Text = "30"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(33, 117)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 14)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Angle"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(301, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 14)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Radius"
        '
        'TextBox_radius1
        '
        Me.TextBox_radius1.Location = New System.Drawing.Point(298, 97)
        Me.TextBox_radius1.Name = "TextBox_radius1"
        Me.TextBox_radius1.Size = New System.Drawing.Size(54, 20)
        Me.TextBox_radius1.TabIndex = 0
        Me.TextBox_radius1.Text = "43.65"
        '
        'Button_fitting
        '
        Me.Button_fitting.Location = New System.Drawing.Point(12, 345)
        Me.Button_fitting.Name = "Button_fitting"
        Me.Button_fitting.Size = New System.Drawing.Size(377, 47)
        Me.Button_fitting.TabIndex = 21
        Me.Button_fitting.Text = "Draw fitting elbow"
        Me.Button_fitting.UseVisualStyleBackColor = True
        '
        'ComboBox_layer_OD
        '
        Me.ComboBox_layer_OD.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer_OD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_OD.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer_OD.FormattingEnabled = True
        Me.ComboBox_layer_OD.Location = New System.Drawing.Point(253, 286)
        Me.ComboBox_layer_OD.Name = "ComboBox_layer_OD"
        Me.ComboBox_layer_OD.Size = New System.Drawing.Size(130, 22)
        Me.ComboBox_layer_OD.TabIndex = 22
        '
        'ComboBox_layer_CL
        '
        Me.ComboBox_layer_CL.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer_CL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_CL.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer_CL.FormattingEnabled = True
        Me.ComboBox_layer_CL.Location = New System.Drawing.Point(12, 286)
        Me.ComboBox_layer_CL.Name = "ComboBox_layer_CL"
        Me.ComboBox_layer_CL.Size = New System.Drawing.Size(130, 22)
        Me.ComboBox_layer_CL.TabIndex = 23
        '
        'fitting_elbow_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(395, 395)
        Me.Controls.Add(Me.ComboBox_layer_OD)
        Me.Controls.Add(Me.ComboBox_layer_CL)
        Me.Controls.Add(Me.Button_fitting)
        Me.Controls.Add(Me.CheckBox_draw_od)
        Me.Controls.Add(Me.ComboBox_nps)
        Me.Controls.Add(Me.TextBox_diameter_X)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label_layer)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = false
        Me.Name = "fitting_elbow_form"
        Me.Text = "Dual fitting elbow"
        Me.Panel2.ResumeLayout(false)
        Me.Panel2.PerformLayout
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents CheckBox_draw_od As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox_nps As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox_diameter_X As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label_layer As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_radius1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_tangent_length As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_elbow_angle As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button_fitting As System.Windows.Forms.Button
    Friend WithEvents TextBox_report As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox_layer_OD As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layer_CL As System.Windows.Forms.ComboBox
End Class
