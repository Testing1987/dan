<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class profiler3d_form
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
        Me.TextBox_Hscale = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_Vscale = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_Hincr = New System.Windows.Forms.TextBox()
        Me.TextBox_Vincr = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_H_Elevation = New System.Windows.Forms.TextBox()
        Me.TextBox_L_elevation = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button_draw = New System.Windows.Forms.Button()
        Me.TextBox_Ground_Chainage = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_3d_length = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CheckBox_PICK_ZERO = New System.Windows.Forms.CheckBox()
        Me.TextBox_text_height = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TextBox_color_index_grid_lines = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.ComboBox_text_styles = New System.Windows.Forms.ComboBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ComboBox_LAYER_GRIDLINES = New System.Windows.Forms.ComboBox()
        Me.ComboBox_LAYER_POLYLINE = New System.Windows.Forms.ComboBox()
        Me.ComboBox_LAYER_TEXT = New System.Windows.Forms.ComboBox()
        Me.ComboBox_LINETYPE = New System.Windows.Forms.ComboBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_Hscale
        '
        Me.TextBox_Hscale.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Hscale.Location = New System.Drawing.Point(202, 282)
        Me.TextBox_Hscale.Name = "TextBox_Hscale"
        Me.TextBox_Hscale.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_Hscale.TabIndex = 7
        Me.TextBox_Hscale.Text = "1000"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(20, 288)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 14)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Horizontal Scale"
        '
        'TextBox_Vscale
        '
        Me.TextBox_Vscale.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Vscale.Location = New System.Drawing.Point(202, 312)
        Me.TextBox_Vscale.Name = "TextBox_Vscale"
        Me.TextBox_Vscale.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_Vscale.TabIndex = 8
        Me.TextBox_Vscale.Text = "1000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(20, 318)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 14)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Vertical Scale"
        '
        'TextBox_Hincr
        '
        Me.TextBox_Hincr.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Hincr.Location = New System.Drawing.Point(202, 344)
        Me.TextBox_Hincr.Name = "TextBox_Hincr"
        Me.TextBox_Hincr.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_Hincr.TabIndex = 9
        Me.TextBox_Hincr.Text = "100"
        '
        'TextBox_Vincr
        '
        Me.TextBox_Vincr.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Vincr.Location = New System.Drawing.Point(202, 374)
        Me.TextBox_Vincr.Name = "TextBox_Vincr"
        Me.TextBox_Vincr.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_Vincr.TabIndex = 10
        Me.TextBox_Vincr.Text = "50"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(20, 349)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 14)
        Me.Label3.TabIndex = 100
        Me.Label3.Text = "Horizontal Increment"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(20, 379)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 14)
        Me.Label4.TabIndex = 100
        Me.Label4.Text = "Vertical Increment"
        '
        'TextBox_H_Elevation
        '
        Me.TextBox_H_Elevation.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_H_Elevation.Location = New System.Drawing.Point(202, 434)
        Me.TextBox_H_Elevation.Name = "TextBox_H_Elevation"
        Me.TextBox_H_Elevation.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_H_Elevation.TabIndex = 12
        '
        'TextBox_L_elevation
        '
        Me.TextBox_L_elevation.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_L_elevation.Location = New System.Drawing.Point(202, 404)
        Me.TextBox_L_elevation.Name = "TextBox_L_elevation"
        Me.TextBox_L_elevation.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_L_elevation.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(20, 409)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(101, 14)
        Me.Label5.TabIndex = 100
        Me.Label5.Text = "Lowest Elevation"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(20, 439)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(101, 14)
        Me.Label6.TabIndex = 100
        Me.Label6.Text = "Highest Elevation"
        '
        'Button_draw
        '
        Me.Button_draw.BackColor = System.Drawing.Color.Lime
        Me.Button_draw.ForeColor = System.Drawing.Color.Black
        Me.Button_draw.Location = New System.Drawing.Point(12, 558)
        Me.Button_draw.Name = "Button_draw"
        Me.Button_draw.Size = New System.Drawing.Size(295, 27)
        Me.Button_draw.TabIndex = 14
        Me.Button_draw.Text = "DRAW!"
        Me.Button_draw.UseVisualStyleBackColor = False
        '
        'TextBox_Ground_Chainage
        '
        Me.TextBox_Ground_Chainage.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Ground_Chainage.Location = New System.Drawing.Point(202, 464)
        Me.TextBox_Ground_Chainage.Name = "TextBox_Ground_Chainage"
        Me.TextBox_Ground_Chainage.ReadOnly = True
        Me.TextBox_Ground_Chainage.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_Ground_Chainage.TabIndex = 70
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(20, 470)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(106, 14)
        Me.Label7.TabIndex = 100
        Me.Label7.Text = "Ground 2D Length"
        '
        'TextBox_3d_length
        '
        Me.TextBox_3d_length.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_3d_length.Location = New System.Drawing.Point(202, 494)
        Me.TextBox_3d_length.Name = "TextBox_3d_length"
        Me.TextBox_3d_length.ReadOnly = True
        Me.TextBox_3d_length.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_3d_length.TabIndex = 70
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label9.Location = New System.Drawing.Point(20, 500)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(106, 14)
        Me.Label9.TabIndex = 100
        Me.Label9.Text = "Ground 3D Length"
        '
        'CheckBox_PICK_ZERO
        '
        Me.CheckBox_PICK_ZERO.AutoSize = True
        Me.CheckBox_PICK_ZERO.Location = New System.Drawing.Point(23, 534)
        Me.CheckBox_PICK_ZERO.Name = "CheckBox_PICK_ZERO"
        Me.CheckBox_PICK_ZERO.Size = New System.Drawing.Size(230, 18)
        Me.CheckBox_PICK_ZERO.TabIndex = 13
        Me.CheckBox_PICK_ZERO.Text = "Pick the 0+000 position on 3D Polyline"
        Me.CheckBox_PICK_ZERO.UseVisualStyleBackColor = True
        '
        'TextBox_text_height
        '
        Me.TextBox_text_height.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_text_height.Location = New System.Drawing.Point(188, 219)
        Me.TextBox_text_height.Name = "TextBox_text_height"
        Me.TextBox_text_height.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_text_height.TabIndex = 4
        Me.TextBox_text_height.Text = "2.5"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label12.Location = New System.Drawing.Point(3, 224)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(68, 14)
        Me.Label12.TabIndex = 100
        Me.Label12.Text = "Text Height"
        '
        'TextBox_color_index_grid_lines
        '
        Me.TextBox_color_index_grid_lines.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_color_index_grid_lines.Location = New System.Drawing.Point(188, 188)
        Me.TextBox_color_index_grid_lines.Name = "TextBox_color_index_grid_lines"
        Me.TextBox_color_index_grid_lines.Size = New System.Drawing.Size(86, 22)
        Me.TextBox_color_index_grid_lines.TabIndex = 3
        Me.TextBox_color_index_grid_lines.Text = "8"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label13.Location = New System.Drawing.Point(3, 194)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(130, 14)
        Me.Label13.TabIndex = 100
        Me.Label13.Text = "Color Index Grid Lines"
        '
        'ComboBox_text_styles
        '
        Me.ComboBox_text_styles.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_text_styles.FormattingEnabled = True
        Me.ComboBox_text_styles.Location = New System.Drawing.Point(73, 29)
        Me.ComboBox_text_styles.Name = "ComboBox_text_styles"
        Me.ComboBox_text_styles.Size = New System.Drawing.Size(198, 22)
        Me.ComboBox_text_styles.TabIndex = 101
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label14.Location = New System.Drawing.Point(3, 32)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 14)
        Me.Label14.TabIndex = 100
        Me.Label14.Text = "Text Style"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label15.Location = New System.Drawing.Point(3, 68)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(98, 14)
        Me.Label15.TabIndex = 100
        Me.Label15.Text = "Layer Grid Lines"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label16.Location = New System.Drawing.Point(3, 128)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 14)
        Me.Label16.TabIndex = 100
        Me.Label16.Text = "Layer Text"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label17.Location = New System.Drawing.Point(3, 164)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(123, 14)
        Me.Label17.TabIndex = 100
        Me.Label17.Text = "Layer Profile Polyline"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.ComboBox_LAYER_GRIDLINES)
        Me.Panel1.Controls.Add(Me.ComboBox_LAYER_POLYLINE)
        Me.Panel1.Controls.Add(Me.ComboBox_LAYER_TEXT)
        Me.Panel1.Controls.Add(Me.ComboBox_LINETYPE)
        Me.Panel1.Controls.Add(Me.ComboBox_text_styles)
        Me.Panel1.Controls.Add(Me.TextBox_text_height)
        Me.Panel1.Controls.Add(Me.TextBox_color_index_grid_lines)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Location = New System.Drawing.Point(12, 13)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(295, 256)
        Me.Panel1.TabIndex = 102
        '
        'ComboBox_LAYER_GRIDLINES
        '
        Me.ComboBox_LAYER_GRIDLINES.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_LAYER_GRIDLINES.FormattingEnabled = True
        Me.ComboBox_LAYER_GRIDLINES.Location = New System.Drawing.Point(136, 59)
        Me.ComboBox_LAYER_GRIDLINES.Name = "ComboBox_LAYER_GRIDLINES"
        Me.ComboBox_LAYER_GRIDLINES.Size = New System.Drawing.Size(138, 22)
        Me.ComboBox_LAYER_GRIDLINES.TabIndex = 101
        '
        'ComboBox_LAYER_POLYLINE
        '
        Me.ComboBox_LAYER_POLYLINE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_LAYER_POLYLINE.FormattingEnabled = True
        Me.ComboBox_LAYER_POLYLINE.Location = New System.Drawing.Point(136, 159)
        Me.ComboBox_LAYER_POLYLINE.Name = "ComboBox_LAYER_POLYLINE"
        Me.ComboBox_LAYER_POLYLINE.Size = New System.Drawing.Size(138, 22)
        Me.ComboBox_LAYER_POLYLINE.TabIndex = 101
        '
        'ComboBox_LAYER_TEXT
        '
        Me.ComboBox_LAYER_TEXT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_LAYER_TEXT.FormattingEnabled = True
        Me.ComboBox_LAYER_TEXT.Location = New System.Drawing.Point(113, 122)
        Me.ComboBox_LAYER_TEXT.Name = "ComboBox_LAYER_TEXT"
        Me.ComboBox_LAYER_TEXT.Size = New System.Drawing.Size(161, 22)
        Me.ComboBox_LAYER_TEXT.TabIndex = 101
        '
        'ComboBox_LINETYPE
        '
        Me.ComboBox_LINETYPE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_LINETYPE.FormattingEnabled = True
        Me.ComboBox_LINETYPE.Location = New System.Drawing.Point(113, 93)
        Me.ComboBox_LINETYPE.Name = "ComboBox_LINETYPE"
        Me.ComboBox_LINETYPE.Size = New System.Drawing.Size(161, 22)
        Me.ComboBox_LINETYPE.TabIndex = 101
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label19.Location = New System.Drawing.Point(3, 98)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(100, 14)
        Me.Label19.TabIndex = 100
        Me.Label19.Text = "Linetype for Grid"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label18.Location = New System.Drawing.Point(6, 5)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(148, 14)
        Me.Label18.TabIndex = 100
        Me.Label18.Text = "Graph format parameters"
        '
        'profiler3d_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(322, 597)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CheckBox_PICK_ZERO)
        Me.Controls.Add(Me.Button_draw)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox_L_elevation)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_Vincr)
        Me.Controls.Add(Me.TextBox_3d_length)
        Me.Controls.Add(Me.TextBox_Ground_Chainage)
        Me.Controls.Add(Me.TextBox_H_Elevation)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_Hincr)
        Me.Controls.Add(Me.TextBox_Vscale)
        Me.Controls.Add(Me.TextBox_Hscale)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "profiler3d_form"
        Me.Text = "Profiler Creator"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_Hscale As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Vscale As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Hincr As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Vincr As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_H_Elevation As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_L_elevation As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button_draw As System.Windows.Forms.Button
    Friend WithEvents TextBox_Ground_Chainage As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_3d_length As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_PICK_ZERO As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox_text_height As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextBox_color_index_grid_lines As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_text_styles As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_LINETYPE As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_LAYER_GRIDLINES As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_LAYER_POLYLINE As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_LAYER_TEXT As System.Windows.Forms.ComboBox
End Class
