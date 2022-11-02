<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _3d_clamp_Form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_3d_clamp_Form))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox_plate_thickness_inches = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_od_mm = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_half_of_length_in = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_clamp_width_inches = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_plate_separation_inches = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_Bolt_dist_fromCL_inches = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox_bolt_nps = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(154, 36)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Draw"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox_plate_thickness_inches
        '
        Me.TextBox_plate_thickness_inches.Location = New System.Drawing.Point(55, 283)
        Me.TextBox_plate_thickness_inches.Name = "TextBox_plate_thickness_inches"
        Me.TextBox_plate_thickness_inches.Size = New System.Drawing.Size(70, 20)
        Me.TextBox_plate_thickness_inches.TabIndex = 1
        Me.TextBox_plate_thickness_inches.Text = "0.25"
        Me.TextBox_plate_thickness_inches.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Gainsboro
        Me.Label1.Location = New System.Drawing.Point(42, 252)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 28)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Plate thickness" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[in]"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_od_mm
        '
        Me.TextBox_od_mm.Location = New System.Drawing.Point(281, 363)
        Me.TextBox_od_mm.Name = "TextBox_od_mm"
        Me.TextBox_od_mm.Size = New System.Drawing.Size(70, 20)
        Me.TextBox_od_mm.TabIndex = 1
        Me.TextBox_od_mm.Text = "88.9"
        Me.TextBox_od_mm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Gainsboro
        Me.Label2.Location = New System.Drawing.Point(292, 332)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 28)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "OD" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[mm]"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Gainsboro
        Me.Label3.Location = New System.Drawing.Point(141, 533)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 28)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Bolt Diameter" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[in]"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_half_of_length_in
        '
        Me.TextBox_half_of_length_in.Location = New System.Drawing.Point(478, 317)
        Me.TextBox_half_of_length_in.Name = "TextBox_half_of_length_in"
        Me.TextBox_half_of_length_in.Size = New System.Drawing.Size(70, 20)
        Me.TextBox_half_of_length_in.TabIndex = 1
        Me.TextBox_half_of_length_in.Text = "4"
        Me.TextBox_half_of_length_in.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Gainsboro
        Me.Label4.Location = New System.Drawing.Point(475, 340)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 42)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Half of" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Total Length" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[in]"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_clamp_width_inches
        '
        Me.TextBox_clamp_width_inches.Location = New System.Drawing.Point(863, 78)
        Me.TextBox_clamp_width_inches.Name = "TextBox_clamp_width_inches"
        Me.TextBox_clamp_width_inches.Size = New System.Drawing.Size(70, 20)
        Me.TextBox_clamp_width_inches.TabIndex = 1
        Me.TextBox_clamp_width_inches.Text = "2"
        Me.TextBox_clamp_width_inches.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Gainsboro
        Me.Label5.Location = New System.Drawing.Point(860, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 28)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Clamp Width" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[in]"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_plate_separation_inches
        '
        Me.TextBox_plate_separation_inches.Location = New System.Drawing.Point(993, 363)
        Me.TextBox_plate_separation_inches.Name = "TextBox_plate_separation_inches"
        Me.TextBox_plate_separation_inches.Size = New System.Drawing.Size(70, 20)
        Me.TextBox_plate_separation_inches.TabIndex = 1
        Me.TextBox_plate_separation_inches.Text = "1"
        Me.TextBox_plate_separation_inches.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Gainsboro
        Me.Label7.Location = New System.Drawing.Point(978, 331)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 28)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Plate Separation" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[in]"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Bolt_dist_fromCL_inches
        '
        Me.TextBox_Bolt_dist_fromCL_inches.Location = New System.Drawing.Point(675, 591)
        Me.TextBox_Bolt_dist_fromCL_inches.Name = "TextBox_Bolt_dist_fromCL_inches"
        Me.TextBox_Bolt_dist_fromCL_inches.Size = New System.Drawing.Size(70, 20)
        Me.TextBox_Bolt_dist_fromCL_inches.TabIndex = 1
        Me.TextBox_Bolt_dist_fromCL_inches.Text = "2.9375"
        Me.TextBox_Bolt_dist_fromCL_inches.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Gainsboro
        Me.Label8.Location = New System.Drawing.Point(660, 614)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(103, 42)
        Me.Label8.TabIndex = 2
        Me.Label8.Text = "Distance from" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Bolt to Centerline" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[in]"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox_bolt_nps
        '
        Me.ComboBox_bolt_nps.BackColor = System.Drawing.Color.White
        Me.ComboBox_bolt_nps.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_bolt_nps.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_bolt_nps.FormattingEnabled = True
        Me.ComboBox_bolt_nps.Location = New System.Drawing.Point(144, 564)
        Me.ComboBox_bolt_nps.Name = "ComboBox_bolt_nps"
        Me.ComboBox_bolt_nps.Size = New System.Drawing.Size(74, 22)
        Me.ComboBox_bolt_nps.TabIndex = 4
        '
        '_3d_clamp_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(1099, 690)
        Me.Controls.Add(Me.ComboBox_bolt_nps)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_od_mm)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TextBox_Bolt_dist_fromCL_inches)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox_half_of_length_in)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TextBox_plate_separation_inches)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBox_clamp_width_inches)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_plate_thickness_inches)
        Me.Controls.Add(Me.Button1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "_3d_clamp_Form"
        Me.Text = "3D CLAMP CREATOR"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox_plate_thickness_inches As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_od_mm As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_half_of_length_in As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_clamp_width_inches As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_plate_separation_inches As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Bolt_dist_fromCL_inches As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_bolt_nps As System.Windows.Forms.ComboBox
End Class
