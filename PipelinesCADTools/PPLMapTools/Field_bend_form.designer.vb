<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Field_bend_form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Field_bend_form))
        Me.Button_draw = New System.Windows.Forms.Button()
        Me.TextBox_radius1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_radius2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label_layer = New System.Windows.Forms.Label()
        Me.ComboBox_nps = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_diameter_X = New System.Windows.Forms.TextBox()
        Me.CheckBox_draw_od = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBox_layer_CL = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layer_OD = New System.Windows.Forms.ComboBox()
        Me.TextBox_report = New System.Windows.Forms.TextBox()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_draw
        '
        Me.Button_draw.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_draw.Location = New System.Drawing.Point(9, 274)
        Me.Button_draw.Name = "Button_draw"
        Me.Button_draw.Size = New System.Drawing.Size(323, 34)
        Me.Button_draw.TabIndex = 0
        Me.Button_draw.Text = "Create"
        Me.Button_draw.UseVisualStyleBackColor = True
        '
        'TextBox_radius1
        '
        Me.TextBox_radius1.Location = New System.Drawing.Point(236, 93)
        Me.TextBox_radius1.Name = "TextBox_radius1"
        Me.TextBox_radius1.Size = New System.Drawing.Size(54, 20)
        Me.TextBox_radius1.TabIndex = 0
        Me.TextBox_radius1.Text = "43.65"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(237, 75)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 14)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Radius 1"
        '
        'TextBox_radius2
        '
        Me.TextBox_radius2.Location = New System.Drawing.Point(44, 69)
        Me.TextBox_radius2.Name = "TextBox_radius2"
        Me.TextBox_radius2.Size = New System.Drawing.Size(54, 20)
        Me.TextBox_radius2.TabIndex = 1
        Me.TextBox_radius2.Text = "43.65"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(45, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 14)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Radius 2"
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_report)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.TextBox_radius2)
        Me.Panel2.Controls.Add(Me.TextBox_radius1)
        Me.Panel2.Location = New System.Drawing.Point(9, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(323, 179)
        Me.Panel2.TabIndex = 3
        '
        'Label_layer
        '
        Me.Label_layer.AutoSize = True
        Me.Label_layer.Location = New System.Drawing.Point(9, 198)
        Me.Label_layer.Name = "Label_layer"
        Me.Label_layer.Size = New System.Drawing.Size(56, 14)
        Me.Label_layer.TabIndex = 2
        Me.Label_layer.Text = "Layer CL"
        '
        'ComboBox_nps
        '
        Me.ComboBox_nps.BackColor = System.Drawing.Color.White
        Me.ComboBox_nps.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_nps.FormattingEnabled = True
        Me.ComboBox_nps.Location = New System.Drawing.Point(82, 246)
        Me.ComboBox_nps.Name = "ComboBox_nps"
        Me.ComboBox_nps.Size = New System.Drawing.Size(103, 22)
        Me.ComboBox_nps.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(42, 248)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 19)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "D x"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_diameter_X
        '
        Me.TextBox_diameter_X.Location = New System.Drawing.Point(12, 247)
        Me.TextBox_diameter_X.Name = "TextBox_diameter_X"
        Me.TextBox_diameter_X.Size = New System.Drawing.Size(32, 20)
        Me.TextBox_diameter_X.TabIndex = 6
        Me.TextBox_diameter_X.Text = "10"
        '
        'CheckBox_draw_od
        '
        Me.CheckBox_draw_od.AutoSize = True
        Me.CheckBox_draw_od.Location = New System.Drawing.Point(248, 250)
        Me.CheckBox_draw_od.Name = "CheckBox_draw_od"
        Me.CheckBox_draw_od.Size = New System.Drawing.Size(72, 18)
        Me.CheckBox_draw_od.TabIndex = 8
        Me.CheckBox_draw_od.Text = "Draw OD"
        Me.CheckBox_draw_od.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(245, 198)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 14)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Layer OD"
        '
        'ComboBox_layer_CL
        '
        Me.ComboBox_layer_CL.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer_CL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_CL.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer_CL.FormattingEnabled = True
        Me.ComboBox_layer_CL.Location = New System.Drawing.Point(9, 213)
        Me.ComboBox_layer_CL.Name = "ComboBox_layer_CL"
        Me.ComboBox_layer_CL.Size = New System.Drawing.Size(130, 22)
        Me.ComboBox_layer_CL.TabIndex = 9
        '
        'ComboBox_layer_OD
        '
        Me.ComboBox_layer_OD.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer_OD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_OD.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer_OD.FormattingEnabled = True
        Me.ComboBox_layer_OD.Location = New System.Drawing.Point(202, 213)
        Me.ComboBox_layer_OD.Name = "ComboBox_layer_OD"
        Me.ComboBox_layer_OD.Size = New System.Drawing.Size(130, 22)
        Me.ComboBox_layer_OD.TabIndex = 9
        '
        'TextBox_report
        '
        Me.TextBox_report.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_report.ForeColor = System.Drawing.Color.Black
        Me.TextBox_report.Location = New System.Drawing.Point(3, 119)
        Me.TextBox_report.Multiline = True
        Me.TextBox_report.Name = "TextBox_report"
        Me.TextBox_report.ReadOnly = True
        Me.TextBox_report.Size = New System.Drawing.Size(152, 53)
        Me.TextBox_report.TabIndex = 23
        Me.TextBox_report.Text = "line1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "line2" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "line3"
        '
        'Field_bend_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(342, 315)
        Me.Controls.Add(Me.ComboBox_layer_OD)
        Me.Controls.Add(Me.ComboBox_layer_CL)
        Me.Controls.Add(Me.CheckBox_draw_od)
        Me.Controls.Add(Me.ComboBox_nps)
        Me.Controls.Add(Me.TextBox_diameter_X)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label_layer)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Button_draw)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Field_bend_form"
        Me.Text = "DUAL field bend"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_draw As System.Windows.Forms.Button
    Friend WithEvents TextBox_radius1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_radius2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label_layer As System.Windows.Forms.Label
    Friend WithEvents ComboBox_nps As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_diameter_X As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_draw_od As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_layer_CL As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layer_OD As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox_report As System.Windows.Forms.TextBox
End Class
