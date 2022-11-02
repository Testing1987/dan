<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Point_insertor_form
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
        Me.TextBox_message = New System.Windows.Forms.TextBox()
        Me.Panel_points = New System.Windows.Forms.Panel()
        Me.ComboBox_poly_layer = New System.Windows.Forms.ComboBox()
        Me.RadioButton_number_description_elevation = New System.Windows.Forms.RadioButton()
        Me.RadioButton_INSERT_Leader = New System.Windows.Forms.RadioButton()
        Me.RadioButton_polyline_only = New System.Windows.Forms.RadioButton()
        Me.RadioButton_point_number_and_elevation = New System.Windows.Forms.RadioButton()
        Me.RadioButton_point_number_and_description = New System.Windows.Forms.RadioButton()
        Me.Button_2D_3D = New System.Windows.Forms.Button()
        Me.RadioButton_point_number_only = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Points_elevation_only = New System.Windows.Forms.RadioButton()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Button_insert_points_to_acad = New System.Windows.Forms.Button()
        Me.Panel_COLUMNS = New System.Windows.Forms.Panel()
        Me.Panel_BLOCKS = New System.Windows.Forms.Panel()
        Me.ComboBox_Layer_for_blocks = New System.Windows.Forms.ComboBox()
        Me.TextBox_Atribut_name2 = New System.Windows.Forms.TextBox()
        Me.TextBox_block_scale = New System.Windows.Forms.TextBox()
        Me.Label_block_scale = New System.Windows.Forms.Label()
        Me.Label_atribut_name = New System.Windows.Forms.Label()
        Me.TextBox_block_name = New System.Windows.Forms.TextBox()
        Me.Label_Block_name = New System.Windows.Forms.Label()
        Me.TextBox_atribut_value1 = New System.Windows.Forms.TextBox()
        Me.Label_block_layer = New System.Windows.Forms.Label()
        Me.TextBox_Atribut_name1 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBox_atribut_value2 = New System.Windows.Forms.TextBox()
        Me.Label_atribut_value = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button_remove_points = New System.Windows.Forms.Button()
        Me.CheckBox_insert_blocks = New System.Windows.Forms.CheckBox()
        Me.CheckBox_line_code = New System.Windows.Forms.CheckBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.TextBox_description = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.TextBox_ln = New System.Windows.Forms.TextBox()
        Me.TextBox_extra2 = New System.Windows.Forms.TextBox()
        Me.TextBox_row_end = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TextBox_NORTH = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.TextBox_layer_prefix = New System.Windows.Forms.TextBox()
        Me.TextBox_decimals = New System.Windows.Forms.TextBox()
        Me.TextBox_row_start = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_extra1 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox_East = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.TextBox_elevation = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_Point_name = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel_points.SuspendLayout()
        Me.Panel_COLUMNS.SuspendLayout()
        Me.Panel_BLOCKS.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_message
        '
        Me.TextBox_message.BackColor = System.Drawing.Color.MidnightBlue
        Me.TextBox_message.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_message.ForeColor = System.Drawing.Color.PeachPuff
        Me.TextBox_message.Location = New System.Drawing.Point(12, 10)
        Me.TextBox_message.Name = "TextBox_message"
        Me.TextBox_message.Size = New System.Drawing.Size(336, 22)
        Me.TextBox_message.TabIndex = 13
        Me.TextBox_message.Text = "Status text box"
        '
        'Panel_points
        '
        Me.Panel_points.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_points.Controls.Add(Me.ComboBox_poly_layer)
        Me.Panel_points.Controls.Add(Me.RadioButton_number_description_elevation)
        Me.Panel_points.Controls.Add(Me.RadioButton_INSERT_Leader)
        Me.Panel_points.Controls.Add(Me.RadioButton_polyline_only)
        Me.Panel_points.Controls.Add(Me.RadioButton_point_number_and_elevation)
        Me.Panel_points.Controls.Add(Me.RadioButton_point_number_and_description)
        Me.Panel_points.Controls.Add(Me.Button_2D_3D)
        Me.Panel_points.Controls.Add(Me.RadioButton_point_number_only)
        Me.Panel_points.Controls.Add(Me.RadioButton_Points_elevation_only)
        Me.Panel_points.Controls.Add(Me.Label9)
        Me.Panel_points.Location = New System.Drawing.Point(12, 445)
        Me.Panel_points.Name = "Panel_points"
        Me.Panel_points.Size = New System.Drawing.Size(336, 171)
        Me.Panel_points.TabIndex = 12
        '
        'ComboBox_poly_layer
        '
        Me.ComboBox_poly_layer.BackColor = System.Drawing.Color.White
        Me.ComboBox_poly_layer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_poly_layer.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_poly_layer.FormattingEnabled = True
        Me.ComboBox_poly_layer.Location = New System.Drawing.Point(224, 139)
        Me.ComboBox_poly_layer.Name = "ComboBox_poly_layer"
        Me.ComboBox_poly_layer.Size = New System.Drawing.Size(106, 21)
        Me.ComboBox_poly_layer.TabIndex = 16
        '
        'RadioButton_number_description_elevation
        '
        Me.RadioButton_number_description_elevation.AutoSize = True
        Me.RadioButton_number_description_elevation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_number_description_elevation.Location = New System.Drawing.Point(3, 95)
        Me.RadioButton_number_description_elevation.Name = "RadioButton_number_description_elevation"
        Me.RadioButton_number_description_elevation.Size = New System.Drawing.Size(250, 17)
        Me.RadioButton_number_description_elevation.TabIndex = 6
        Me.RadioButton_number_description_elevation.Text = "Point number, description and elevation"
        Me.RadioButton_number_description_elevation.UseVisualStyleBackColor = True
        '
        'RadioButton_INSERT_Leader
        '
        Me.RadioButton_INSERT_Leader.AutoSize = True
        Me.RadioButton_INSERT_Leader.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_INSERT_Leader.Location = New System.Drawing.Point(3, 118)
        Me.RadioButton_INSERT_Leader.Name = "RadioButton_INSERT_Leader"
        Me.RadioButton_INSERT_Leader.Size = New System.Drawing.Size(123, 17)
        Me.RadioButton_INSERT_Leader.TabIndex = 6
        Me.RadioButton_INSERT_Leader.Text = "Add Mleader only"
        Me.RadioButton_INSERT_Leader.UseVisualStyleBackColor = True
        '
        'RadioButton_polyline_only
        '
        Me.RadioButton_polyline_only.AutoSize = True
        Me.RadioButton_polyline_only.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_polyline_only.Location = New System.Drawing.Point(4, 140)
        Me.RadioButton_polyline_only.Name = "RadioButton_polyline_only"
        Me.RadioButton_polyline_only.Size = New System.Drawing.Size(131, 17)
        Me.RadioButton_polyline_only.TabIndex = 6
        Me.RadioButton_polyline_only.Text = "Draw Polyline Only"
        Me.RadioButton_polyline_only.UseVisualStyleBackColor = True
        '
        'RadioButton_point_number_and_elevation
        '
        Me.RadioButton_point_number_and_elevation.AutoSize = True
        Me.RadioButton_point_number_and_elevation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_point_number_and_elevation.Location = New System.Drawing.Point(3, 72)
        Me.RadioButton_point_number_and_elevation.Name = "RadioButton_point_number_and_elevation"
        Me.RadioButton_point_number_and_elevation.Size = New System.Drawing.Size(180, 17)
        Me.RadioButton_point_number_and_elevation.TabIndex = 6
        Me.RadioButton_point_number_and_elevation.Text = "Point number and elevation"
        Me.RadioButton_point_number_and_elevation.UseVisualStyleBackColor = True
        '
        'RadioButton_point_number_and_description
        '
        Me.RadioButton_point_number_and_description.AutoSize = True
        Me.RadioButton_point_number_and_description.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_point_number_and_description.Location = New System.Drawing.Point(3, 49)
        Me.RadioButton_point_number_and_description.Name = "RadioButton_point_number_and_description"
        Me.RadioButton_point_number_and_description.Size = New System.Drawing.Size(190, 17)
        Me.RadioButton_point_number_and_description.TabIndex = 6
        Me.RadioButton_point_number_and_description.Text = "Point number and description"
        Me.RadioButton_point_number_and_description.UseVisualStyleBackColor = True
        '
        'Button_2D_3D
        '
        Me.Button_2D_3D.BackColor = System.Drawing.Color.Magenta
        Me.Button_2D_3D.Font = New System.Drawing.Font("Bookman Old Style", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_2D_3D.ForeColor = System.Drawing.Color.Yellow
        Me.Button_2D_3D.Location = New System.Drawing.Point(293, 3)
        Me.Button_2D_3D.Name = "Button_2D_3D"
        Me.Button_2D_3D.Size = New System.Drawing.Size(36, 31)
        Me.Button_2D_3D.TabIndex = 5
        Me.Button_2D_3D.Text = "3D"
        Me.Button_2D_3D.UseVisualStyleBackColor = False
        '
        'RadioButton_point_number_only
        '
        Me.RadioButton_point_number_only.AutoSize = True
        Me.RadioButton_point_number_only.Checked = True
        Me.RadioButton_point_number_only.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_point_number_only.Location = New System.Drawing.Point(3, 26)
        Me.RadioButton_point_number_only.Name = "RadioButton_point_number_only"
        Me.RadioButton_point_number_only.Size = New System.Drawing.Size(99, 17)
        Me.RadioButton_point_number_only.TabIndex = 6
        Me.RadioButton_point_number_only.TabStop = True
        Me.RadioButton_point_number_only.Text = "Point number"
        Me.RadioButton_point_number_only.UseVisualStyleBackColor = True
        '
        'RadioButton_Points_elevation_only
        '
        Me.RadioButton_Points_elevation_only.AutoSize = True
        Me.RadioButton_Points_elevation_only.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton_Points_elevation_only.Location = New System.Drawing.Point(3, 3)
        Me.RadioButton_Points_elevation_only.Name = "RadioButton_Points_elevation_only"
        Me.RadioButton_Points_elevation_only.Size = New System.Drawing.Size(105, 17)
        Me.RadioButton_Points_elevation_only.TabIndex = 6
        Me.RadioButton_Points_elevation_only.Text = "Elevation only"
        Me.RadioButton_Points_elevation_only.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(142, 142)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(66, 13)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Poly Layer"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button_insert_points_to_acad
        '
        Me.Button_insert_points_to_acad.BackColor = System.Drawing.Color.Blue
        Me.Button_insert_points_to_acad.Font = New System.Drawing.Font("Bookman Old Style", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_insert_points_to_acad.ForeColor = System.Drawing.Color.White
        Me.Button_insert_points_to_acad.Location = New System.Drawing.Point(12, 622)
        Me.Button_insert_points_to_acad.Name = "Button_insert_points_to_acad"
        Me.Button_insert_points_to_acad.Size = New System.Drawing.Size(118, 27)
        Me.Button_insert_points_to_acad.TabIndex = 10
        Me.Button_insert_points_to_acad.Text = "EXCEL->ACAD"
        Me.Button_insert_points_to_acad.UseVisualStyleBackColor = False
        '
        'Panel_COLUMNS
        '
        Me.Panel_COLUMNS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_COLUMNS.Controls.Add(Me.Panel_BLOCKS)
        Me.Panel_COLUMNS.Controls.Add(Me.Button_remove_points)
        Me.Panel_COLUMNS.Controls.Add(Me.CheckBox_insert_blocks)
        Me.Panel_COLUMNS.Controls.Add(Me.CheckBox_line_code)
        Me.Panel_COLUMNS.Controls.Add(Me.Label8)
        Me.Panel_COLUMNS.Controls.Add(Me.Label15)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_description)
        Me.Panel_COLUMNS.Controls.Add(Me.Label1)
        Me.Panel_COLUMNS.Controls.Add(Me.Label20)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_ln)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_extra2)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_row_end)
        Me.Panel_COLUMNS.Controls.Add(Me.Label16)
        Me.Panel_COLUMNS.Controls.Add(Me.Label14)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_NORTH)
        Me.Panel_COLUMNS.Controls.Add(Me.Label13)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_layer_prefix)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_decimals)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_row_start)
        Me.Panel_COLUMNS.Controls.Add(Me.Label7)
        Me.Panel_COLUMNS.Controls.Add(Me.Label4)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_extra1)
        Me.Panel_COLUMNS.Controls.Add(Me.Label6)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_East)
        Me.Panel_COLUMNS.Controls.Add(Me.Label18)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_elevation)
        Me.Panel_COLUMNS.Controls.Add(Me.Label3)
        Me.Panel_COLUMNS.Controls.Add(Me.Label5)
        Me.Panel_COLUMNS.Controls.Add(Me.TextBox_Point_name)
        Me.Panel_COLUMNS.Controls.Add(Me.Label2)
        Me.Panel_COLUMNS.Location = New System.Drawing.Point(12, 38)
        Me.Panel_COLUMNS.Name = "Panel_COLUMNS"
        Me.Panel_COLUMNS.Size = New System.Drawing.Size(336, 401)
        Me.Panel_COLUMNS.TabIndex = 11
        '
        'Panel_BLOCKS
        '
        Me.Panel_BLOCKS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_BLOCKS.Controls.Add(Me.ComboBox_Layer_for_blocks)
        Me.Panel_BLOCKS.Controls.Add(Me.TextBox_Atribut_name2)
        Me.Panel_BLOCKS.Controls.Add(Me.TextBox_block_scale)
        Me.Panel_BLOCKS.Controls.Add(Me.Label_block_scale)
        Me.Panel_BLOCKS.Controls.Add(Me.Label_atribut_name)
        Me.Panel_BLOCKS.Controls.Add(Me.TextBox_block_name)
        Me.Panel_BLOCKS.Controls.Add(Me.Label_Block_name)
        Me.Panel_BLOCKS.Controls.Add(Me.TextBox_atribut_value1)
        Me.Panel_BLOCKS.Controls.Add(Me.Label_block_layer)
        Me.Panel_BLOCKS.Controls.Add(Me.TextBox_Atribut_name1)
        Me.Panel_BLOCKS.Controls.Add(Me.Label10)
        Me.Panel_BLOCKS.Controls.Add(Me.TextBox_atribut_value2)
        Me.Panel_BLOCKS.Controls.Add(Me.Label_atribut_value)
        Me.Panel_BLOCKS.Controls.Add(Me.Label11)
        Me.Panel_BLOCKS.Location = New System.Drawing.Point(3, 104)
        Me.Panel_BLOCKS.Name = "Panel_BLOCKS"
        Me.Panel_BLOCKS.Size = New System.Drawing.Size(323, 165)
        Me.Panel_BLOCKS.TabIndex = 14
        '
        'ComboBox_Layer_for_blocks
        '
        Me.ComboBox_Layer_for_blocks.BackColor = System.Drawing.Color.White
        Me.ComboBox_Layer_for_blocks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Layer_for_blocks.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_Layer_for_blocks.FormattingEnabled = True
        Me.ComboBox_Layer_for_blocks.Location = New System.Drawing.Point(3, 126)
        Me.ComboBox_Layer_for_blocks.Name = "ComboBox_Layer_for_blocks"
        Me.ComboBox_Layer_for_blocks.Size = New System.Drawing.Size(172, 21)
        Me.ComboBox_Layer_for_blocks.TabIndex = 16
        '
        'TextBox_Atribut_name2
        '
        Me.TextBox_Atribut_name2.BackColor = System.Drawing.Color.White
        Me.TextBox_Atribut_name2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Atribut_name2.Location = New System.Drawing.Point(139, 79)
        Me.TextBox_Atribut_name2.Name = "TextBox_Atribut_name2"
        Me.TextBox_Atribut_name2.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_Atribut_name2.TabIndex = 10
        Me.TextBox_Atribut_name2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_block_scale
        '
        Me.TextBox_block_scale.BackColor = System.Drawing.Color.White
        Me.TextBox_block_scale.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_block_scale.Location = New System.Drawing.Point(275, 131)
        Me.TextBox_block_scale.Name = "TextBox_block_scale"
        Me.TextBox_block_scale.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_block_scale.TabIndex = 10
        Me.TextBox_block_scale.Text = "1"
        Me.TextBox_block_scale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_block_scale
        '
        Me.Label_block_scale.AutoSize = True
        Me.Label_block_scale.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_block_scale.ForeColor = System.Drawing.Color.White
        Me.Label_block_scale.Location = New System.Drawing.Point(194, 134)
        Me.Label_block_scale.Name = "Label_block_scale"
        Me.Label_block_scale.Size = New System.Drawing.Size(75, 13)
        Me.Label_block_scale.TabIndex = 3
        Me.Label_block_scale.Text = "Block Scale"
        Me.Label_block_scale.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_atribut_name
        '
        Me.Label_atribut_name.AutoSize = True
        Me.Label_atribut_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_atribut_name.ForeColor = System.Drawing.Color.White
        Me.Label_atribut_name.Location = New System.Drawing.Point(0, 26)
        Me.Label_atribut_name.Name = "Label_atribut_name"
        Me.Label_atribut_name.Size = New System.Drawing.Size(59, 39)
        Me.Label_atribut_name.TabIndex = 3
        Me.Label_atribut_name.Text = "Attribute " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Name " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(TAG)1"
        Me.Label_atribut_name.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_block_name
        '
        Me.TextBox_block_name.BackColor = System.Drawing.Color.White
        Me.TextBox_block_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_block_name.Location = New System.Drawing.Point(3, 3)
        Me.TextBox_block_name.Name = "TextBox_block_name"
        Me.TextBox_block_name.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_block_name.TabIndex = 10
        Me.TextBox_block_name.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_Block_name
        '
        Me.Label_Block_name.AutoSize = True
        Me.Label_Block_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Block_name.ForeColor = System.Drawing.Color.White
        Me.Label_Block_name.Location = New System.Drawing.Point(42, 6)
        Me.Label_Block_name.Name = "Label_Block_name"
        Me.Label_Block_name.Size = New System.Drawing.Size(119, 13)
        Me.Label_Block_name.TabIndex = 3
        Me.Label_Block_name.Text = "Block Name column"
        Me.Label_Block_name.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_atribut_value1
        '
        Me.TextBox_atribut_value1.BackColor = System.Drawing.Color.White
        Me.TextBox_atribut_value1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_atribut_value1.Location = New System.Drawing.Point(75, 79)
        Me.TextBox_atribut_value1.Name = "TextBox_atribut_value1"
        Me.TextBox_atribut_value1.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_atribut_value1.TabIndex = 10
        Me.TextBox_atribut_value1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_block_layer
        '
        Me.Label_block_layer.AutoSize = True
        Me.Label_block_layer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_block_layer.ForeColor = System.Drawing.Color.White
        Me.Label_block_layer.Location = New System.Drawing.Point(0, 109)
        Me.Label_block_layer.Name = "Label_block_layer"
        Me.Label_block_layer.Size = New System.Drawing.Size(74, 13)
        Me.Label_block_layer.TabIndex = 3
        Me.Label_block_layer.Text = "Block Layer"
        Me.Label_block_layer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Atribut_name1
        '
        Me.TextBox_Atribut_name1.BackColor = System.Drawing.Color.White
        Me.TextBox_Atribut_name1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Atribut_name1.Location = New System.Drawing.Point(15, 79)
        Me.TextBox_Atribut_name1.Name = "TextBox_Atribut_name1"
        Me.TextBox_Atribut_name1.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_Atribut_name1.TabIndex = 10
        Me.TextBox_Atribut_name1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(123, 30)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(59, 39)
        Me.Label10.TabIndex = 3
        Me.Label10.Text = "Attribute " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Name " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(TAG)2"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_atribut_value2
        '
        Me.TextBox_atribut_value2.BackColor = System.Drawing.Color.White
        Me.TextBox_atribut_value2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_atribut_value2.Location = New System.Drawing.Point(199, 78)
        Me.TextBox_atribut_value2.Name = "TextBox_atribut_value2"
        Me.TextBox_atribut_value2.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_atribut_value2.TabIndex = 10
        Me.TextBox_atribut_value2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_atribut_value
        '
        Me.Label_atribut_value.AutoSize = True
        Me.Label_atribut_value.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_atribut_value.ForeColor = System.Drawing.Color.White
        Me.Label_atribut_value.Location = New System.Drawing.Point(63, 46)
        Me.Label_atribut_value.Name = "Label_atribut_value"
        Me.Label_atribut_value.Size = New System.Drawing.Size(55, 26)
        Me.Label_atribut_value.TabIndex = 3
        Me.Label_atribut_value.Text = "Attribute" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Value1"
        Me.Label_atribut_value.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(187, 46)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(55, 26)
        Me.Label11.TabIndex = 3
        Me.Label11.Text = "Attribute" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Value2"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button_remove_points
        '
        Me.Button_remove_points.BackColor = System.Drawing.Color.Red
        Me.Button_remove_points.Font = New System.Drawing.Font("Bookman Old Style", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_remove_points.ForeColor = System.Drawing.Color.White
        Me.Button_remove_points.Location = New System.Drawing.Point(250, 322)
        Me.Button_remove_points.Name = "Button_remove_points"
        Me.Button_remove_points.Size = New System.Drawing.Size(79, 62)
        Me.Button_remove_points.TabIndex = 15
        Me.Button_remove_points.Text = "Remove all  data"
        Me.Button_remove_points.UseVisualStyleBackColor = False
        '
        'CheckBox_insert_blocks
        '
        Me.CheckBox_insert_blocks.AutoSize = True
        Me.CheckBox_insert_blocks.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_insert_blocks.Location = New System.Drawing.Point(129, 328)
        Me.CheckBox_insert_blocks.Name = "CheckBox_insert_blocks"
        Me.CheckBox_insert_blocks.Size = New System.Drawing.Size(64, 17)
        Me.CheckBox_insert_blocks.TabIndex = 10
        Me.CheckBox_insert_blocks.Text = "Blocks"
        Me.CheckBox_insert_blocks.UseVisualStyleBackColor = True
        '
        'CheckBox_line_code
        '
        Me.CheckBox_line_code.AutoSize = True
        Me.CheckBox_line_code.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_line_code.Location = New System.Drawing.Point(129, 308)
        Me.CheckBox_line_code.Name = "CheckBox_line_code"
        Me.CheckBox_line_code.Size = New System.Drawing.Size(89, 17)
        Me.CheckBox_line_code.TabIndex = 10
        Me.CheckBox_line_code.Text = "Line Codes"
        Me.CheckBox_line_code.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(178, 47)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 13)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "Descr"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(83, 325)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(29, 13)
        Me.Label15.TabIndex = 9
        Me.Label15.Text = "End"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_description
        '
        Me.TextBox_description.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_description.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_description.Location = New System.Drawing.Point(182, 24)
        Me.TextBox_description.Name = "TextBox_description"
        Me.TextBox_description.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_description.TabIndex = 5
        Me.TextBox_description.Text = "E"
        Me.TextBox_description.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(2, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "COLUMNS IN EXCEL"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Black
        Me.Label20.Location = New System.Drawing.Point(249, 82)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(72, 13)
        Me.Label20.TabIndex = 3
        Me.Label20.Text = "Layer name"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_ln
        '
        Me.TextBox_ln.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_ln.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_ln.Location = New System.Drawing.Point(269, 63)
        Me.TextBox_ln.Name = "TextBox_ln"
        Me.TextBox_ln.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_ln.TabIndex = 8
        Me.TextBox_ln.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_extra2
        '
        Me.TextBox_extra2.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_extra2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_extra2.Location = New System.Drawing.Point(269, 24)
        Me.TextBox_extra2.Name = "TextBox_extra2"
        Me.TextBox_extra2.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_extra2.TabIndex = 8
        Me.TextBox_extra2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_row_end
        '
        Me.TextBox_row_end.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_row_end.Location = New System.Drawing.Point(8, 322)
        Me.TextBox_row_end.Name = "TextBox_row_end"
        Me.TextBox_row_end.Size = New System.Drawing.Size(69, 20)
        Me.TextBox_row_end.TabIndex = 1
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(158, 288)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(58, 13)
        Me.Label16.TabIndex = 7
        Me.Label16.Text = "Decimals"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(84, 299)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(34, 13)
        Me.Label14.TabIndex = 7
        Me.Label14.Text = "Start"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_NORTH
        '
        Me.TextBox_NORTH.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_NORTH.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_NORTH.Location = New System.Drawing.Point(93, 24)
        Me.TextBox_NORTH.Name = "TextBox_NORTH"
        Me.TextBox_NORTH.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_NORTH.TabIndex = 4
        Me.TextBox_NORTH.Text = "D"
        Me.TextBox_NORTH.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(9, 272)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(105, 13)
        Me.Label13.TabIndex = 6
        Me.Label13.Text = "ROWS IN EXCEL"
        '
        'TextBox_layer_prefix
        '
        Me.TextBox_layer_prefix.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_layer_prefix.Location = New System.Drawing.Point(86, 361)
        Me.TextBox_layer_prefix.Name = "TextBox_layer_prefix"
        Me.TextBox_layer_prefix.Size = New System.Drawing.Size(130, 20)
        Me.TextBox_layer_prefix.TabIndex = 5
        '
        'TextBox_decimals
        '
        Me.TextBox_decimals.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_decimals.Location = New System.Drawing.Point(129, 285)
        Me.TextBox_decimals.Name = "TextBox_decimals"
        Me.TextBox_decimals.Size = New System.Drawing.Size(23, 20)
        Me.TextBox_decimals.TabIndex = 5
        Me.TextBox_decimals.Text = "3"
        '
        'TextBox_row_start
        '
        Me.TextBox_row_start.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_row_start.Location = New System.Drawing.Point(9, 296)
        Me.TextBox_row_start.Name = "TextBox_row_start"
        Me.TextBox_row_start.Size = New System.Drawing.Size(69, 20)
        Me.TextBox_row_start.TabIndex = 0
        Me.TextBox_row_start.Text = "10"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(269, 47)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(42, 13)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "extra2"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(99, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "N"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_extra1
        '
        Me.TextBox_extra1.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_extra1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_extra1.Location = New System.Drawing.Point(225, 24)
        Me.TextBox_extra1.Name = "TextBox_extra1"
        Me.TextBox_extra1.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_extra1.TabIndex = 7
        Me.TextBox_extra1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(230, 47)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(42, 13)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "extra1"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_East
        '
        Me.TextBox_East.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_East.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_East.Location = New System.Drawing.Point(49, 24)
        Me.TextBox_East.Name = "TextBox_East"
        Me.TextBox_East.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_East.TabIndex = 3
        Me.TextBox_East.Text = "C"
        Me.TextBox_East.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Black
        Me.Label18.Location = New System.Drawing.Point(7, 364)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(73, 13)
        Me.Label18.TabIndex = 3
        Me.Label18.Text = "Layer prefix"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_elevation
        '
        Me.TextBox_elevation.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_elevation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_elevation.Location = New System.Drawing.Point(138, 24)
        Me.TextBox_elevation.Name = "TextBox_elevation"
        Me.TextBox_elevation.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_elevation.TabIndex = 6
        Me.TextBox_elevation.Text = "E"
        Me.TextBox_elevation.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(52, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(15, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "E"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(143, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(22, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "EL"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Point_name
        '
        Me.TextBox_Point_name.BackColor = System.Drawing.Color.Khaki
        Me.TextBox_Point_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Point_name.Location = New System.Drawing.Point(5, 24)
        Me.TextBox_Point_name.Name = "TextBox_Point_name"
        Me.TextBox_Point_name.Size = New System.Drawing.Size(33, 20)
        Me.TextBox_Point_name.TabIndex = 2
        Me.TextBox_Point_name.Text = "G"
        Me.TextBox_Point_name.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(6, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "PN"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Point_insertor_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(360, 661)
        Me.Controls.Add(Me.TextBox_message)
        Me.Controls.Add(Me.Panel_points)
        Me.Controls.Add(Me.Button_insert_points_to_acad)
        Me.Controls.Add(Me.Panel_COLUMNS)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Point_insertor_form"
        Me.Text = "Point Insertor"
        Me.Panel_points.ResumeLayout(False)
        Me.Panel_points.PerformLayout()
        Me.Panel_COLUMNS.ResumeLayout(False)
        Me.Panel_COLUMNS.PerformLayout()
        Me.Panel_BLOCKS.ResumeLayout(False)
        Me.Panel_BLOCKS.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_message As System.Windows.Forms.TextBox
    Friend WithEvents Panel_points As System.Windows.Forms.Panel
    Friend WithEvents RadioButton_number_description_elevation As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_point_number_and_elevation As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_point_number_and_description As System.Windows.Forms.RadioButton
    Friend WithEvents Button_2D_3D As System.Windows.Forms.Button
    Friend WithEvents RadioButton_point_number_only As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Points_elevation_only As System.Windows.Forms.RadioButton
    Friend WithEvents Button_insert_points_to_acad As System.Windows.Forms.Button
    Friend WithEvents Panel_COLUMNS As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_insert_blocks As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_line_code As System.Windows.Forms.CheckBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TextBox_extra2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_row_end As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox_NORTH As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TextBox_layer_prefix As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_decimals As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_row_start As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_extra1 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox_East As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TextBox_elevation As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_description As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Point_name As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_remove_points As System.Windows.Forms.Button
    Friend WithEvents Label_Block_name As System.Windows.Forms.Label
    Friend WithEvents TextBox_block_name As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton_polyline_only As System.Windows.Forms.RadioButton
    Friend WithEvents ComboBox_Layer_for_blocks As System.Windows.Forms.ComboBox
    Friend WithEvents Label_block_layer As System.Windows.Forms.Label
    Friend WithEvents ComboBox_poly_layer As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label_atribut_value As System.Windows.Forms.Label
    Friend WithEvents Label_atribut_name As System.Windows.Forms.Label
    Friend WithEvents TextBox_atribut_value1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Atribut_name1 As System.Windows.Forms.TextBox
    Friend WithEvents Label_block_scale As System.Windows.Forms.Label
    Friend WithEvents TextBox_block_scale As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton_INSERT_Leader As System.Windows.Forms.RadioButton
    Friend WithEvents Panel_BLOCKS As System.Windows.Forms.Panel
    Friend WithEvents TextBox_Atribut_name2 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBox_atribut_value2 As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox_ln As Windows.Forms.TextBox
    Friend WithEvents Label20 As Windows.Forms.Label
End Class
