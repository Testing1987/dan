<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _3d_Ubolt_form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_3d_Ubolt_form))
        Me.TextBox_stud_diam_mm = New System.Windows.Forms.TextBox()
        Me.TextBox_GAP_mm = New System.Windows.Forms.TextBox()
        Me.TextBox_extend_mm = New System.Windows.Forms.TextBox()
        Me.TextBox_plate_thickness = New System.Windows.Forms.TextBox()
        Me.ComboBox_NPS = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox_stud_diam_mm
        '
        Me.TextBox_stud_diam_mm.Location = New System.Drawing.Point(582, 571)
        Me.TextBox_stud_diam_mm.Name = "TextBox_stud_diam_mm"
        Me.TextBox_stud_diam_mm.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_stud_diam_mm.TabIndex = 101
        Me.TextBox_stud_diam_mm.Text = "10"
        Me.TextBox_stud_diam_mm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_GAP_mm
        '
        Me.TextBox_GAP_mm.Location = New System.Drawing.Point(670, 68)
        Me.TextBox_GAP_mm.Name = "TextBox_GAP_mm"
        Me.TextBox_GAP_mm.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_GAP_mm.TabIndex = 101
        Me.TextBox_GAP_mm.Text = "5"
        Me.TextBox_GAP_mm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_extend_mm
        '
        Me.TextBox_extend_mm.Location = New System.Drawing.Point(45, 437)
        Me.TextBox_extend_mm.Name = "TextBox_extend_mm"
        Me.TextBox_extend_mm.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_extend_mm.TabIndex = 101
        Me.TextBox_extend_mm.Text = "53"
        Me.TextBox_extend_mm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_plate_thickness
        '
        Me.TextBox_plate_thickness.Location = New System.Drawing.Point(849, 334)
        Me.TextBox_plate_thickness.Name = "TextBox_plate_thickness"
        Me.TextBox_plate_thickness.Size = New System.Drawing.Size(57, 20)
        Me.TextBox_plate_thickness.TabIndex = 101
        Me.TextBox_plate_thickness.Text = "6"
        Me.TextBox_plate_thickness.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ComboBox_NPS
        '
        Me.ComboBox_NPS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_NPS.FormattingEnabled = True
        Me.ComboBox_NPS.Location = New System.Drawing.Point(409, 193)
        Me.ComboBox_NPS.Name = "ComboBox_NPS"
        Me.ComboBox_NPS.Size = New System.Drawing.Size(53, 22)
        Me.ComboBox_NPS.TabIndex = 103
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(413, 176)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 14)
        Me.Label4.TabIndex = 102
        Me.Label4.Text = "NPS [in]"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(815, 571)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(91, 33)
        Me.Button1.TabIndex = 104
        Me.Button1.Text = "Draw"
        Me.Button1.UseVisualStyleBackColor = True
        '
        '_3d_Ubolt_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(930, 616)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ComboBox_NPS)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox_plate_thickness)
        Me.Controls.Add(Me.TextBox_extend_mm)
        Me.Controls.Add(Me.TextBox_GAP_mm)
        Me.Controls.Add(Me.TextBox_stud_diam_mm)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "_3d_Ubolt_form"
        Me.Text = "Ubolt builder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_stud_diam_mm As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_GAP_mm As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_extend_mm As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_plate_thickness As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox_NPS As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
