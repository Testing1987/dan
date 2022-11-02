<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Intersection_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CheckBox_Line_direction = New System.Windows.Forms.CheckBox()
        Me.CheckBox_select_multiple_CL = New System.Windows.Forms.CheckBox()
        Me.CheckBox_output_layers = New System.Windows.Forms.CheckBox()
        Me.CheckBox_Object_data = New System.Windows.Forms.CheckBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Button_calc_start_end = New System.Windows.Forms.Button()
        Me.Button_calculate_int = New System.Windows.Forms.Button()
        Me.ButtonTotal_length_of_reroutes = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.CheckBox_scan_points = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox_buffer = New System.Windows.Forms.TextBox()
        Me.Button_poly_3d_face_scanning = New System.Windows.Forms.Button()
        Me.Button_residence_scanning = New System.Windows.Forms.Button()
        Me.Button_scan_single_point_segment = New System.Windows.Forms.Button()
        Me.Button_Scan_segments = New System.Windows.Forms.Button()
        Me.Button_output_offset = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.CheckBox_zero_decimals = New System.Windows.Forms.CheckBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_LOAD_LAYER_NAMES = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox_start_row = New System.Windows.Forms.TextBox()
        Me.TextBox_layer_description = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_end_row = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_layer_name = New System.Windows.Forms.TextBox()
        Me.CheckBox_US_station = New System.Windows.Forms.CheckBox()
        Me.CheckBox_no_CSF = New System.Windows.Forms.CheckBox()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_buffer_for_excelPT = New System.Windows.Forms.TextBox()
        Me.TextBox_offset = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBox_station = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_start = New System.Windows.Forms.TextBox()
        Me.TextBox_end = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_description = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_Point_name = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.TextBox_east_intersection = New System.Windows.Forms.TextBox()
        Me.TextBox_north_intersection = New System.Windows.Forms.TextBox()
        Me.TextBox_elevation_INTERSECTION = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Button_station_to_pointUSA = New System.Windows.Forms.Button()
        Me.Button_point_to_Station_usa = New System.Windows.Forms.Button()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.Button_load_equations_from_excel = New System.Windows.Forms.Button()
        Me.Panel18 = New System.Windows.Forms.Panel()
        Me.TextBox_Row_End_eq = New System.Windows.Forms.TextBox()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.TextBox_Row_Start_eq = New System.Windows.Forms.TextBox()
        Me.CheckBox_use_equation = New System.Windows.Forms.CheckBox()
        Me.Panel19 = New System.Windows.Forms.Panel()
        Me.TextBox_col_statation_ahead = New System.Windows.Forms.TextBox()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.TextBox_col_station_back = New System.Windows.Forms.TextBox()
        Me.Button_cl_2d_crossing_3D = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        Me.Panel18.SuspendLayout()
        Me.Panel19.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.CheckBox_Line_direction)
        Me.Panel2.Controls.Add(Me.CheckBox_select_multiple_CL)
        Me.Panel2.Controls.Add(Me.CheckBox_output_layers)
        Me.Panel2.Controls.Add(Me.CheckBox_Object_data)
        Me.Panel2.Controls.Add(Me.Label20)
        Me.Panel2.Controls.Add(Me.Button_cl_2d_crossing_3D)
        Me.Panel2.Controls.Add(Me.Button_calculate_int)
        Me.Panel2.Location = New System.Drawing.Point(6, 6)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(405, 209)
        Me.Panel2.TabIndex = 104
        '
        'CheckBox_Line_direction
        '
        Me.CheckBox_Line_direction.AutoSize = True
        Me.CheckBox_Line_direction.Location = New System.Drawing.Point(8, 78)
        Me.CheckBox_Line_direction.Name = "CheckBox_Line_direction"
        Me.CheckBox_Line_direction.Size = New System.Drawing.Size(134, 19)
        Me.CheckBox_Line_direction.TabIndex = 110
        Me.CheckBox_Line_direction.Text = "Show line direction"
        Me.CheckBox_Line_direction.UseVisualStyleBackColor = True
        '
        'CheckBox_select_multiple_CL
        '
        Me.CheckBox_select_multiple_CL.AutoSize = True
        Me.CheckBox_select_multiple_CL.Location = New System.Drawing.Point(8, 3)
        Me.CheckBox_select_multiple_CL.Name = "CheckBox_select_multiple_CL"
        Me.CheckBox_select_multiple_CL.Size = New System.Drawing.Size(197, 19)
        Me.CheckBox_select_multiple_CL.TabIndex = 110
        Me.CheckBox_select_multiple_CL.Text = "Multiple centerlines (reroutes)"
        Me.CheckBox_select_multiple_CL.UseVisualStyleBackColor = True
        '
        'CheckBox_output_layers
        '
        Me.CheckBox_output_layers.AutoSize = True
        Me.CheckBox_output_layers.Checked = True
        Me.CheckBox_output_layers.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_output_layers.Location = New System.Drawing.Point(8, 28)
        Me.CheckBox_output_layers.Name = "CheckBox_output_layers"
        Me.CheckBox_output_layers.Size = New System.Drawing.Size(142, 19)
        Me.CheckBox_output_layers.TabIndex = 111
        Me.CheckBox_output_layers.Text = "Output Layer Names"
        Me.CheckBox_output_layers.UseVisualStyleBackColor = True
        '
        'CheckBox_Object_data
        '
        Me.CheckBox_Object_data.AutoSize = True
        Me.CheckBox_Object_data.Checked = True
        Me.CheckBox_Object_data.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_Object_data.Location = New System.Drawing.Point(8, 53)
        Me.CheckBox_Object_data.Name = "CheckBox_Object_data"
        Me.CheckBox_Object_data.Size = New System.Drawing.Size(130, 19)
        Me.CheckBox_Object_data.TabIndex = 110
        Me.CheckBox_Object_data.Text = "Output object data"
        Me.CheckBox_Object_data.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(250, 115)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(148, 75)
        Me.Label20.TabIndex = 6
        Me.Label20.Text = "Station" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of the intersection points" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "between a 3D or 2D " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "polyline and 2D polylin" &
    "es" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Button_calc_start_end
        '
        Me.Button_calc_start_end.Location = New System.Drawing.Point(22, 73)
        Me.Button_calc_start_end.Name = "Button_calc_start_end"
        Me.Button_calc_start_end.Size = New System.Drawing.Size(149, 75)
        Me.Button_calc_start_end.TabIndex = 109
        Me.Button_calc_start_end.Text = "output all start-end points of segments on top of centerline"
        Me.Button_calc_start_end.UseVisualStyleBackColor = True
        '
        'Button_calculate_int
        '
        Me.Button_calculate_int.Location = New System.Drawing.Point(8, 115)
        Me.Button_calculate_int.Name = "Button_calculate_int"
        Me.Button_calculate_int.Size = New System.Drawing.Size(217, 75)
        Me.Button_calculate_int.TabIndex = 109
        Me.Button_calculate_int.Text = "Output all intersection points for each polyline intersecting the centerline"
        Me.Button_calculate_int.UseVisualStyleBackColor = True
        '
        'ButtonTotal_length_of_reroutes
        '
        Me.ButtonTotal_length_of_reroutes.Location = New System.Drawing.Point(279, 3)
        Me.ButtonTotal_length_of_reroutes.Name = "ButtonTotal_length_of_reroutes"
        Me.ButtonTotal_length_of_reroutes.Size = New System.Drawing.Size(122, 44)
        Me.ButtonTotal_length_of_reroutes.TabIndex = 112
        Me.ButtonTotal_length_of_reroutes.Text = "REROUTE LENGTH (grid)"
        Me.ButtonTotal_length_of_reroutes.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(424, 250)
        Me.TabControl1.TabIndex = 105
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage1.Controls.Add(Me.Panel2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 27)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(416, 219)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Intersect"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage2.Controls.Add(Me.CheckBox_scan_points)
        Me.TabPage2.Controls.Add(Me.Button_calc_start_end)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.TextBox_buffer)
        Me.TabPage2.Controls.Add(Me.Button_poly_3d_face_scanning)
        Me.TabPage2.Controls.Add(Me.Button_residence_scanning)
        Me.TabPage2.Controls.Add(Me.Button_scan_single_point_segment)
        Me.TabPage2.Controls.Add(Me.Button_Scan_segments)
        Me.TabPage2.Controls.Add(Me.Button_output_offset)
        Me.TabPage2.Location = New System.Drawing.Point(4, 27)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(416, 219)
        Me.TabPage2.TabIndex = 2
        Me.TabPage2.Text = "Point Features"
        '
        'CheckBox_scan_points
        '
        Me.CheckBox_scan_points.AutoSize = True
        Me.CheckBox_scan_points.Location = New System.Drawing.Point(22, 16)
        Me.CheckBox_scan_points.Name = "CheckBox_scan_points"
        Me.CheckBox_scan_points.Size = New System.Drawing.Size(189, 19)
        Me.CheckBox_scan_points.TabIndex = 123
        Me.CheckBox_scan_points.Text = "Add Points objects to search"
        Me.CheckBox_scan_points.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(19, 46)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(42, 15)
        Me.Label6.TabIndex = 117
        Me.Label6.Text = "Buffer"
        '
        'TextBox_buffer
        '
        Me.TextBox_buffer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_buffer.Location = New System.Drawing.Point(67, 44)
        Me.TextBox_buffer.Name = "TextBox_buffer"
        Me.TextBox_buffer.Size = New System.Drawing.Size(66, 20)
        Me.TextBox_buffer.TabIndex = 118
        Me.TextBox_buffer.Text = "200"
        '
        'Button_poly_3d_face_scanning
        '
        Me.Button_poly_3d_face_scanning.Location = New System.Drawing.Point(249, 16)
        Me.Button_poly_3d_face_scanning.Name = "Button_poly_3d_face_scanning"
        Me.Button_poly_3d_face_scanning.Size = New System.Drawing.Size(156, 45)
        Me.Button_poly_3d_face_scanning.TabIndex = 110
        Me.Button_poly_3d_face_scanning.Text = "Polylines N point"
        Me.Button_poly_3d_face_scanning.UseVisualStyleBackColor = True
        '
        'Button_residence_scanning
        '
        Me.Button_residence_scanning.Location = New System.Drawing.Point(249, 67)
        Me.Button_residence_scanning.Name = "Button_residence_scanning"
        Me.Button_residence_scanning.Size = New System.Drawing.Size(156, 32)
        Me.Button_residence_scanning.TabIndex = 110
        Me.Button_residence_scanning.Text = "Residence Scanning"
        Me.Button_residence_scanning.UseVisualStyleBackColor = True
        '
        'Button_scan_single_point_segment
        '
        Me.Button_scan_single_point_segment.Location = New System.Drawing.Point(249, 156)
        Me.Button_scan_single_point_segment.Name = "Button_scan_single_point_segment"
        Me.Button_scan_single_point_segment.Size = New System.Drawing.Size(156, 47)
        Me.Button_scan_single_point_segment.TabIndex = 110
        Me.Button_scan_single_point_segment.Text = "Scan  for segments" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Single point)"
        Me.Button_scan_single_point_segment.UseVisualStyleBackColor = True
        '
        'Button_Scan_segments
        '
        Me.Button_Scan_segments.Location = New System.Drawing.Point(249, 103)
        Me.Button_Scan_segments.Name = "Button_Scan_segments"
        Me.Button_Scan_segments.Size = New System.Drawing.Size(156, 47)
        Me.Button_Scan_segments.TabIndex = 110
        Me.Button_Scan_segments.Text = "Scan  for segments" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Start-End)"
        Me.Button_Scan_segments.UseVisualStyleBackColor = True
        '
        'Button_output_offset
        '
        Me.Button_output_offset.Location = New System.Drawing.Point(22, 154)
        Me.Button_output_offset.Name = "Button_output_offset"
        Me.Button_output_offset.Size = New System.Drawing.Size(189, 51)
        Me.Button_output_offset.TabIndex = 110
        Me.Button_output_offset.Text = "Output all positions of blocks in relation to the centerline"
        Me.Button_output_offset.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage3.Controls.Add(Me.Panel1)
        Me.TabPage3.Location = New System.Drawing.Point(4, 27)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(416, 219)
        Me.TabPage3.TabIndex = 1
        Me.TabPage3.Text = "Settings"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.ButtonTotal_length_of_reroutes)
        Me.Panel1.Controls.Add(Me.CheckBox_zero_decimals)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.CheckBox_US_station)
        Me.Panel1.Controls.Add(Me.CheckBox_no_CSF)
        Me.Panel1.Location = New System.Drawing.Point(6, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(408, 210)
        Me.Panel1.TabIndex = 0
        '
        'CheckBox_zero_decimals
        '
        Me.CheckBox_zero_decimals.AutoSize = True
        Me.CheckBox_zero_decimals.Location = New System.Drawing.Point(3, 53)
        Me.CheckBox_zero_decimals.Name = "CheckBox_zero_decimals"
        Me.CheckBox_zero_decimals.Size = New System.Drawing.Size(215, 19)
        Me.CheckBox_zero_decimals.TabIndex = 122
        Me.CheckBox_zero_decimals.Text = "Round to nearest integer (0 decs)"
        Me.CheckBox_zero_decimals.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Button_LOAD_LAYER_NAMES)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.TextBox_start_row)
        Me.Panel3.Controls.Add(Me.TextBox_layer_description)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.TextBox_end_row)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.TextBox_layer_name)
        Me.Panel3.Location = New System.Drawing.Point(3, 93)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(313, 96)
        Me.Panel3.TabIndex = 121
        '
        'Button_LOAD_LAYER_NAMES
        '
        Me.Button_LOAD_LAYER_NAMES.Location = New System.Drawing.Point(6, 59)
        Me.Button_LOAD_LAYER_NAMES.Name = "Button_LOAD_LAYER_NAMES"
        Me.Button_LOAD_LAYER_NAMES.Size = New System.Drawing.Size(142, 28)
        Me.Button_LOAD_LAYER_NAMES.TabIndex = 115
        Me.Button_LOAD_LAYER_NAMES.Text = "Load Layer names"
        Me.Button_LOAD_LAYER_NAMES.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(189, 37)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 15)
        Me.Label9.TabIndex = 117
        Me.Label9.Text = "End Row"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 15)
        Me.Label8.TabIndex = 118
        Me.Label8.Text = "Layer column"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_start_row
        '
        Me.TextBox_start_row.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_start_row.Location = New System.Drawing.Point(259, 7)
        Me.TextBox_start_row.Name = "TextBox_start_row"
        Me.TextBox_start_row.Size = New System.Drawing.Size(39, 20)
        Me.TextBox_start_row.TabIndex = 119
        Me.TextBox_start_row.Text = "2"
        '
        'TextBox_layer_description
        '
        Me.TextBox_layer_description.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_layer_description.Location = New System.Drawing.Point(130, 33)
        Me.TextBox_layer_description.Name = "TextBox_layer_description"
        Me.TextBox_layer_description.Size = New System.Drawing.Size(39, 20)
        Me.TextBox_layer_description.TabIndex = 120
        Me.TextBox_layer_description.Text = "B"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 37)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(118, 15)
        Me.Label7.TabIndex = 117
        Me.Label7.Text = "Description Column"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_end_row
        '
        Me.TextBox_end_row.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_end_row.Location = New System.Drawing.Point(259, 33)
        Me.TextBox_end_row.Name = "TextBox_end_row"
        Me.TextBox_end_row.Size = New System.Drawing.Size(39, 20)
        Me.TextBox_end_row.TabIndex = 120
        Me.TextBox_end_row.Text = "63"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(189, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 15)
        Me.Label2.TabIndex = 118
        Me.Label2.Text = "Start Row"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_layer_name
        '
        Me.TextBox_layer_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_layer_name.Location = New System.Drawing.Point(130, 7)
        Me.TextBox_layer_name.Name = "TextBox_layer_name"
        Me.TextBox_layer_name.Size = New System.Drawing.Size(39, 20)
        Me.TextBox_layer_name.TabIndex = 119
        Me.TextBox_layer_name.Text = "A"
        '
        'CheckBox_US_station
        '
        Me.CheckBox_US_station.AutoSize = True
        Me.CheckBox_US_station.Checked = True
        Me.CheckBox_US_station.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_US_station.Location = New System.Drawing.Point(3, 28)
        Me.CheckBox_US_station.Name = "CheckBox_US_station"
        Me.CheckBox_US_station.Size = New System.Drawing.Size(166, 19)
        Me.CheckBox_US_station.TabIndex = 111
        Me.CheckBox_US_station.Text = "Display Stations US style"
        Me.CheckBox_US_station.UseVisualStyleBackColor = True
        '
        'CheckBox_no_CSF
        '
        Me.CheckBox_no_CSF.AutoSize = True
        Me.CheckBox_no_CSF.Checked = True
        Me.CheckBox_no_CSF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_no_CSF.Location = New System.Drawing.Point(3, 3)
        Me.CheckBox_no_CSF.Name = "CheckBox_no_CSF"
        Me.CheckBox_no_CSF.Size = New System.Drawing.Size(68, 19)
        Me.CheckBox_no_CSF.TabIndex = 112
        Me.CheckBox_no_CSF.Text = "NO CSF"
        Me.CheckBox_no_CSF.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage4.Controls.Add(Me.Label11)
        Me.TabPage4.Controls.Add(Me.TextBox_buffer_for_excelPT)
        Me.TabPage4.Controls.Add(Me.TextBox_offset)
        Me.TabPage4.Controls.Add(Me.Label10)
        Me.TabPage4.Controls.Add(Me.TextBox_station)
        Me.TabPage4.Controls.Add(Me.Label21)
        Me.TabPage4.Controls.Add(Me.Label3)
        Me.TabPage4.Controls.Add(Me.TextBox_start)
        Me.TabPage4.Controls.Add(Me.TextBox_end)
        Me.TabPage4.Controls.Add(Me.Label4)
        Me.TabPage4.Controls.Add(Me.TextBox_description)
        Me.TabPage4.Controls.Add(Me.Label1)
        Me.TabPage4.Controls.Add(Me.TextBox_Point_name)
        Me.TabPage4.Controls.Add(Me.Label5)
        Me.TabPage4.Controls.Add(Me.Label25)
        Me.TabPage4.Controls.Add(Me.TextBox_east_intersection)
        Me.TabPage4.Controls.Add(Me.TextBox_north_intersection)
        Me.TabPage4.Controls.Add(Me.TextBox_elevation_INTERSECTION)
        Me.TabPage4.Controls.Add(Me.Label26)
        Me.TabPage4.Controls.Add(Me.Label22)
        Me.TabPage4.Controls.Add(Me.Button_station_to_pointUSA)
        Me.TabPage4.Controls.Add(Me.Button_point_to_Station_usa)
        Me.TabPage4.Location = New System.Drawing.Point(4, 27)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(416, 219)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Others"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(252, 103)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(42, 15)
        Me.Label11.TabIndex = 127
        Me.Label11.Text = "Buffer"
        '
        'TextBox_buffer_for_excelPT
        '
        Me.TextBox_buffer_for_excelPT.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_buffer_for_excelPT.Location = New System.Drawing.Point(300, 101)
        Me.TextBox_buffer_for_excelPT.Name = "TextBox_buffer_for_excelPT"
        Me.TextBox_buffer_for_excelPT.Size = New System.Drawing.Size(66, 20)
        Me.TextBox_buffer_for_excelPT.TabIndex = 128
        Me.TextBox_buffer_for_excelPT.Text = "200"
        '
        'TextBox_offset
        '
        Me.TextBox_offset.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_offset.Location = New System.Drawing.Point(332, 75)
        Me.TextBox_offset.Name = "TextBox_offset"
        Me.TextBox_offset.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_offset.TabIndex = 126
        Me.TextBox_offset.Text = "G"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label10.Location = New System.Drawing.Point(326, 55)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(44, 17)
        Me.Label10.TabIndex = 125
        Me.Label10.Text = "Offset"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_station
        '
        Me.TextBox_station.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_station.Location = New System.Drawing.Point(332, 24)
        Me.TextBox_station.Name = "TextBox_station"
        Me.TextBox_station.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_station.TabIndex = 126
        Me.TextBox_station.Text = "F"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label21.Location = New System.Drawing.Point(326, 4)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(49, 17)
        Me.Label21.TabIndex = 125
        Me.Label21.Text = "Station"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Location = New System.Drawing.Point(10, 90)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 17)
        Me.Label3.TabIndex = 121
        Me.Label3.Text = "End Row"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_start
        '
        Me.TextBox_start.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_start.Location = New System.Drawing.Point(79, 61)
        Me.TextBox_start.Name = "TextBox_start"
        Me.TextBox_start.Size = New System.Drawing.Size(61, 20)
        Me.TextBox_start.TabIndex = 123
        Me.TextBox_start.Text = "1"
        '
        'TextBox_end
        '
        Me.TextBox_end.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_end.Location = New System.Drawing.Point(79, 88)
        Me.TextBox_end.Name = "TextBox_end"
        Me.TextBox_end.Size = New System.Drawing.Size(61, 20)
        Me.TextBox_end.TabIndex = 124
        Me.TextBox_end.Text = "2"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label4.Location = New System.Drawing.Point(10, 63)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(65, 17)
        Me.Label4.TabIndex = 122
        Me.Label4.Text = "Start Row"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_description
        '
        Me.TextBox_description.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_description.Location = New System.Drawing.Point(266, 24)
        Me.TextBox_description.Name = "TextBox_description"
        Me.TextBox_description.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_description.TabIndex = 119
        Me.TextBox_description.Text = "E"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(246, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 17)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "Description"
        '
        'TextBox_Point_name
        '
        Me.TextBox_Point_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Point_name.Location = New System.Drawing.Point(24, 24)
        Me.TextBox_Point_name.Name = "TextBox_Point_name"
        Me.TextBox_Point_name.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_Point_name.TabIndex = 120
        Me.TextBox_Point_name.Text = "A"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label5.Location = New System.Drawing.Point(4, 4)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 17)
        Me.Label5.TabIndex = 118
        Me.Label5.Text = "Point Name"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label25.Location = New System.Drawing.Point(128, 4)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(40, 17)
        Me.Label25.TabIndex = 111
        Me.Label25.Text = "North"
        '
        'TextBox_east_intersection
        '
        Me.TextBox_east_intersection.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_east_intersection.Location = New System.Drawing.Point(85, 24)
        Me.TextBox_east_intersection.Name = "TextBox_east_intersection"
        Me.TextBox_east_intersection.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_east_intersection.TabIndex = 114
        Me.TextBox_east_intersection.Text = "B"
        '
        'TextBox_north_intersection
        '
        Me.TextBox_north_intersection.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_north_intersection.Location = New System.Drawing.Point(131, 24)
        Me.TextBox_north_intersection.Name = "TextBox_north_intersection"
        Me.TextBox_north_intersection.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_north_intersection.TabIndex = 115
        Me.TextBox_north_intersection.Text = "C"
        '
        'TextBox_elevation_INTERSECTION
        '
        Me.TextBox_elevation_INTERSECTION.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_elevation_INTERSECTION.Location = New System.Drawing.Point(183, 24)
        Me.TextBox_elevation_INTERSECTION.Name = "TextBox_elevation_INTERSECTION"
        Me.TextBox_elevation_INTERSECTION.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_elevation_INTERSECTION.TabIndex = 116
        Me.TextBox_elevation_INTERSECTION.Text = "D"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label26.Location = New System.Drawing.Point(82, 4)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(34, 17)
        Me.Label26.TabIndex = 112
        Me.Label26.Text = "East"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label22.Location = New System.Drawing.Point(180, 4)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(60, 17)
        Me.Label22.TabIndex = 113
        Me.Label22.Text = "Elevation"
        '
        'Button_station_to_pointUSA
        '
        Me.Button_station_to_pointUSA.Location = New System.Drawing.Point(212, 164)
        Me.Button_station_to_pointUSA.Name = "Button_station_to_pointUSA"
        Me.Button_station_to_pointUSA.Size = New System.Drawing.Size(176, 46)
        Me.Button_station_to_pointUSA.TabIndex = 110
        Me.Button_station_to_pointUSA.Text = "Calculates coordinates of a Station from Excel"
        Me.Button_station_to_pointUSA.UseVisualStyleBackColor = True
        '
        'Button_point_to_Station_usa
        '
        Me.Button_point_to_Station_usa.Location = New System.Drawing.Point(10, 126)
        Me.Button_point_to_Station_usa.Name = "Button_point_to_Station_usa"
        Me.Button_point_to_Station_usa.Size = New System.Drawing.Size(176, 46)
        Me.Button_point_to_Station_usa.TabIndex = 110
        Me.Button_point_to_Station_usa.Text = "Calculates the Station" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of a point from Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Button_point_to_Station_usa.UseVisualStyleBackColor = True
        '
        'TabPage6
        '
        Me.TabPage6.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage6.Controls.Add(Me.Button_load_equations_from_excel)
        Me.TabPage6.Controls.Add(Me.Panel18)
        Me.TabPage6.Controls.Add(Me.CheckBox_use_equation)
        Me.TabPage6.Controls.Add(Me.Panel19)
        Me.TabPage6.Location = New System.Drawing.Point(4, 27)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(416, 219)
        Me.TabPage6.TabIndex = 6
        Me.TabPage6.Text = "Equations"
        '
        'Button_load_equations_from_excel
        '
        Me.Button_load_equations_from_excel.Location = New System.Drawing.Point(182, 129)
        Me.Button_load_equations_from_excel.Name = "Button_load_equations_from_excel"
        Me.Button_load_equations_from_excel.Size = New System.Drawing.Size(211, 36)
        Me.Button_load_equations_from_excel.TabIndex = 2003
        Me.Button_load_equations_from_excel.Text = "Load From Excel"
        Me.Button_load_equations_from_excel.UseVisualStyleBackColor = True
        '
        'Panel18
        '
        Me.Panel18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel18.Controls.Add(Me.TextBox_Row_End_eq)
        Me.Panel18.Controls.Add(Me.Label75)
        Me.Panel18.Controls.Add(Me.Label76)
        Me.Panel18.Controls.Add(Me.TextBox_Row_Start_eq)
        Me.Panel18.Location = New System.Drawing.Point(8, 95)
        Me.Panel18.Name = "Panel18"
        Me.Panel18.Size = New System.Drawing.Size(143, 70)
        Me.Panel18.TabIndex = 2001
        '
        'TextBox_Row_End_eq
        '
        Me.TextBox_Row_End_eq.BackColor = System.Drawing.Color.White
        Me.TextBox_Row_End_eq.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Row_End_eq.Location = New System.Drawing.Point(77, 36)
        Me.TextBox_Row_End_eq.Name = "TextBox_Row_End_eq"
        Me.TextBox_Row_End_eq.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_Row_End_eq.TabIndex = 3
        Me.TextBox_Row_End_eq.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label75
        '
        Me.Label75.AutoSize = True
        Me.Label75.Location = New System.Drawing.Point(5, 8)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(63, 15)
        Me.Label75.TabIndex = 300
        Me.Label75.Text = "Row Start"
        '
        'Label76
        '
        Me.Label76.AutoSize = True
        Me.Label76.Location = New System.Drawing.Point(11, 39)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(56, 15)
        Me.Label76.TabIndex = 300
        Me.Label76.Text = "Row End"
        '
        'TextBox_Row_Start_eq
        '
        Me.TextBox_Row_Start_eq.BackColor = System.Drawing.Color.White
        Me.TextBox_Row_Start_eq.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Row_Start_eq.Location = New System.Drawing.Point(77, 5)
        Me.TextBox_Row_Start_eq.Name = "TextBox_Row_Start_eq"
        Me.TextBox_Row_Start_eq.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_Row_Start_eq.TabIndex = 2
        Me.TextBox_Row_Start_eq.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckBox_use_equation
        '
        Me.CheckBox_use_equation.AutoSize = True
        Me.CheckBox_use_equation.Location = New System.Drawing.Point(8, 8)
        Me.CheckBox_use_equation.Name = "CheckBox_use_equation"
        Me.CheckBox_use_equation.Size = New System.Drawing.Size(150, 19)
        Me.CheckBox_use_equation.TabIndex = 7
        Me.CheckBox_use_equation.Text = "Use Station Equations"
        Me.CheckBox_use_equation.UseVisualStyleBackColor = True
        '
        'Panel19
        '
        Me.Panel19.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel19.Controls.Add(Me.TextBox_col_statation_ahead)
        Me.Panel19.Controls.Add(Me.Label77)
        Me.Panel19.Controls.Add(Me.Label78)
        Me.Panel19.Controls.Add(Me.Label79)
        Me.Panel19.Controls.Add(Me.TextBox_col_station_back)
        Me.Panel19.Location = New System.Drawing.Point(8, 33)
        Me.Panel19.Name = "Panel19"
        Me.Panel19.Size = New System.Drawing.Size(407, 56)
        Me.Panel19.TabIndex = 2002
        '
        'TextBox_col_statation_ahead
        '
        Me.TextBox_col_statation_ahead.BackColor = System.Drawing.Color.White
        Me.TextBox_col_statation_ahead.ForeColor = System.Drawing.Color.Black
        Me.TextBox_col_statation_ahead.Location = New System.Drawing.Point(172, 28)
        Me.TextBox_col_statation_ahead.Name = "TextBox_col_statation_ahead"
        Me.TextBox_col_statation_ahead.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_col_statation_ahead.TabIndex = 1
        Me.TextBox_col_statation_ahead.Text = "C"
        Me.TextBox_col_statation_ahead.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label77
        '
        Me.Label77.AutoSize = True
        Me.Label77.Location = New System.Drawing.Point(3, 0)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(57, 15)
        Me.Label77.TabIndex = 300
        Me.Label77.Text = "Columns"
        '
        'Label78
        '
        Me.Label78.AutoSize = True
        Me.Label78.Location = New System.Drawing.Point(71, 10)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(79, 15)
        Me.Label78.TabIndex = 300
        Me.Label78.Text = "Station Back"
        '
        'Label79
        '
        Me.Label79.AutoSize = True
        Me.Label79.Location = New System.Drawing.Point(159, 10)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(86, 15)
        Me.Label79.TabIndex = 300
        Me.Label79.Text = "Station Ahead"
        '
        'TextBox_col_station_back
        '
        Me.TextBox_col_station_back.BackColor = System.Drawing.Color.White
        Me.TextBox_col_station_back.ForeColor = System.Drawing.Color.Black
        Me.TextBox_col_station_back.Location = New System.Drawing.Point(69, 28)
        Me.TextBox_col_station_back.Name = "TextBox_col_station_back"
        Me.TextBox_col_station_back.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_col_station_back.TabIndex = 0
        Me.TextBox_col_station_back.Text = "B"
        Me.TextBox_col_station_back.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_cl_2d_crossing_3D
        '
        Me.Button_cl_2d_crossing_3D.Location = New System.Drawing.Point(186, 22)
        Me.Button_cl_2d_crossing_3D.Name = "Button_cl_2d_crossing_3D"
        Me.Button_cl_2d_crossing_3D.Size = New System.Drawing.Size(217, 75)
        Me.Button_cl_2d_crossing_3D.TabIndex = 109
        Me.Button_cl_2d_crossing_3D.Text = "Intersect 2D cl with 3D  Polyline Crossings"
        Me.Button_cl_2d_crossing_3D.UseVisualStyleBackColor = True
        '
        'Intersection_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(442, 263)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Intersection_Form"
        Me.Text = "Centerline Intersect "
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage6.ResumeLayout(False)
        Me.TabPage6.PerformLayout()
        Me.Panel18.ResumeLayout(False)
        Me.Panel18.PerformLayout()
        Me.Panel19.ResumeLayout(False)
        Me.Panel19.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Button_calculate_int As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_output_layers As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_no_CSF As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_US_station As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Button_output_offset As System.Windows.Forms.Button
    Friend WithEvents CheckBox_Object_data As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox_buffer As System.Windows.Forms.TextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button_LOAD_LAYER_NAMES As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox_start_row As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_layer_description As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_end_row As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_layer_name As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_zero_decimals As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_scan_points As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents CheckBox_select_multiple_CL As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonTotal_length_of_reroutes As System.Windows.Forms.Button
    Friend WithEvents Button_Scan_segments As System.Windows.Forms.Button
    Friend WithEvents Button_residence_scanning As System.Windows.Forms.Button
    Friend WithEvents Button_poly_3d_face_scanning As System.Windows.Forms.Button
    Friend WithEvents Button_scan_single_point_segment As System.Windows.Forms.Button
    Friend WithEvents Button_point_to_Station_usa As System.Windows.Forms.Button
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents TextBox_east_intersection As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_north_intersection As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_elevation_INTERSECTION As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TextBox_description As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Point_name As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_start As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_end As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_station As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TextBox_offset As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox_buffer_for_excelPT As System.Windows.Forms.TextBox
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents Button_load_equations_from_excel As System.Windows.Forms.Button
    Friend WithEvents Panel18 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_Row_End_eq As System.Windows.Forms.TextBox
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents Label76 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Row_Start_eq As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_use_equation As System.Windows.Forms.CheckBox
    Friend WithEvents Panel19 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_col_statation_ahead As System.Windows.Forms.TextBox
    Friend WithEvents Label77 As System.Windows.Forms.Label
    Friend WithEvents Label78 As System.Windows.Forms.Label
    Friend WithEvents Label79 As System.Windows.Forms.Label
    Friend WithEvents TextBox_col_station_back As System.Windows.Forms.TextBox
    Friend WithEvents Button_station_to_pointUSA As System.Windows.Forms.Button
    Friend WithEvents Button_calc_start_end As Windows.Forms.Button
    Friend WithEvents CheckBox_Line_direction As Windows.Forms.CheckBox
    Friend WithEvents Button_cl_2d_crossing_3D As Windows.Forms.Button
End Class
