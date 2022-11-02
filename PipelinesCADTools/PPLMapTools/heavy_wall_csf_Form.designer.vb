<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class heavy_wall_csf_Form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(heavy_wall_csf_Form))
        Me.Button_2dgrid = New System.Windows.Forms.Button()
        Me.Button_CSF_TO_AUTOCAD = New System.Windows.Forms.Button()
        Me.Button_pt_to_CSF = New System.Windows.Forms.Button()
        Me.Button_output_to_excel_parallel_middle = New System.Windows.Forms.Button()
        Me.Button_output_to_excel_top = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.TextBox_end = New System.Windows.Forms.TextBox()
        Me.TextBox_start = New System.Windows.Forms.TextBox()
        Me.TextBox_chainage = New System.Windows.Forms.TextBox()
        Me.TextBox_CSF_Chainage = New System.Windows.Forms.TextBox()
        Me.TextBox_elevation = New System.Windows.Forms.TextBox()
        Me.TextBox_north_intersection = New System.Windows.Forms.TextBox()
        Me.TextBox_east_intersection = New System.Windows.Forms.TextBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button_output_to_excel_parallel_ends = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Button_3d_2d = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Button_start_end_without_CSF = New System.Windows.Forms.Button()
        Me.Button_rerouteMP = New System.Windows.Forms.Button()
        Me.Button_read_sta_wr_2xl = New System.Windows.Forms.Button()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Button_read_write_csf = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Button_reroute_projection = New System.Windows.Forms.Button()
        Me.TextBox_s1 = New System.Windows.Forms.TextBox()
        Me.TextBoxx1 = New System.Windows.Forms.TextBox()
        Me.TextBox_e1 = New System.Windows.Forms.TextBox()
        Me.TextBoxy1 = New System.Windows.Forms.TextBox()
        Me.TextBox_X = New System.Windows.Forms.TextBox()
        Me.TextBoxs = New System.Windows.Forms.TextBox()
        Me.TextBoxs1 = New System.Windows.Forms.TextBox()
        Me.TextBoxz1 = New System.Windows.Forms.TextBox()
        Me.TextBox_Y = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBox_Z = New System.Windows.Forms.TextBox()
        Me.TextBoxCSF2 = New System.Windows.Forms.TextBox()
        Me.TextBoxcsf1 = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox_CSF_STA = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.Button_gen_station = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_2dgrid
        '
        Me.Button_2dgrid.Location = New System.Drawing.Point(262, -2)
        Me.Button_2dgrid.Name = "Button_2dgrid"
        Me.Button_2dgrid.Size = New System.Drawing.Size(125, 113)
        Me.Button_2dgrid.TabIndex = 109
        Me.Button_2dgrid.Text = "Reads from Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "grid 2d chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "calculates" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "csf 3d chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and X,Y,Z"
        Me.Button_2dgrid.UseVisualStyleBackColor = True
        '
        'Button_CSF_TO_AUTOCAD
        '
        Me.Button_CSF_TO_AUTOCAD.Location = New System.Drawing.Point(136, -1)
        Me.Button_CSF_TO_AUTOCAD.Name = "Button_CSF_TO_AUTOCAD"
        Me.Button_CSF_TO_AUTOCAD.Size = New System.Drawing.Size(125, 113)
        Me.Button_CSF_TO_AUTOCAD.TabIndex = 109
        Me.Button_CSF_TO_AUTOCAD.Text = "Reads Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "CSF Chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "writes back to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Excel X,Y,Z"
        Me.Button_CSF_TO_AUTOCAD.UseVisualStyleBackColor = True
        '
        'Button_pt_to_CSF
        '
        Me.Button_pt_to_CSF.Location = New System.Drawing.Point(9, 131)
        Me.Button_pt_to_CSF.Name = "Button_pt_to_CSF"
        Me.Button_pt_to_CSF.Size = New System.Drawing.Size(128, 78)
        Me.Button_pt_to_CSF.TabIndex = 109
        Me.Button_pt_to_CSF.Text = "Output to Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "CSF Chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of a point" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "from Excel"
        Me.Button_pt_to_CSF.UseVisualStyleBackColor = True
        '
        'Button_output_to_excel_parallel_middle
        '
        Me.Button_output_to_excel_parallel_middle.Location = New System.Drawing.Point(6, 7)
        Me.Button_output_to_excel_parallel_middle.Name = "Button_output_to_excel_parallel_middle"
        Me.Button_output_to_excel_parallel_middle.Size = New System.Drawing.Size(362, 109)
        Me.Button_output_to_excel_parallel_middle.TabIndex = 109
        Me.Button_output_to_excel_parallel_middle.Text = resources.GetString("Button_output_to_excel_parallel_middle.Text")
        Me.Button_output_to_excel_parallel_middle.UseVisualStyleBackColor = True
        '
        'Button_output_to_excel_top
        '
        Me.Button_output_to_excel_top.Location = New System.Drawing.Point(2, 1)
        Me.Button_output_to_excel_top.Name = "Button_output_to_excel_top"
        Me.Button_output_to_excel_top.Size = New System.Drawing.Size(128, 109)
        Me.Button_output_to_excel_top.TabIndex = 109
        Me.Button_output_to_excel_top.Text = "Output to Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Grid 3D Chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of a segment " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(start - end)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "drafted on top" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) &
    "of the CL"
        Me.Button_output_to_excel_top.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 15)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "End Row"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 15)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Start Row"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(0, 21)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(60, 15)
        Me.Label21.TabIndex = 6
        Me.Label21.Text = "Chainage"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(1, -1)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(57, 15)
        Me.Label24.TabIndex = 6
        Me.Label24.Text = "COLUMN"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(216, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 15)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "CSF CHAINAGE"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(156, 21)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(58, 15)
        Me.Label22.TabIndex = 6
        Me.Label22.Text = "Elevation"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(104, 21)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(38, 15)
        Me.Label25.TabIndex = 6
        Me.Label25.Text = "North"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(58, 21)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(32, 15)
        Me.Label26.TabIndex = 6
        Me.Label26.Text = "East"
        '
        'TextBox_end
        '
        Me.TextBox_end.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_end.Location = New System.Drawing.Point(87, 36)
        Me.TextBox_end.Name = "TextBox_end"
        Me.TextBox_end.Size = New System.Drawing.Size(61, 20)
        Me.TextBox_end.TabIndex = 104
        Me.TextBox_end.Text = "2"
        '
        'TextBox_start
        '
        Me.TextBox_start.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_start.Location = New System.Drawing.Point(87, 9)
        Me.TextBox_start.Name = "TextBox_start"
        Me.TextBox_start.Size = New System.Drawing.Size(61, 20)
        Me.TextBox_start.TabIndex = 104
        Me.TextBox_start.Text = "2"
        '
        'TextBox_chainage
        '
        Me.TextBox_chainage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_chainage.Location = New System.Drawing.Point(3, 37)
        Me.TextBox_chainage.Name = "TextBox_chainage"
        Me.TextBox_chainage.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_chainage.TabIndex = 104
        Me.TextBox_chainage.Text = "A"
        '
        'TextBox_CSF_Chainage
        '
        Me.TextBox_CSF_Chainage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_CSF_Chainage.Location = New System.Drawing.Point(233, 38)
        Me.TextBox_CSF_Chainage.Name = "TextBox_CSF_Chainage"
        Me.TextBox_CSF_Chainage.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_CSF_Chainage.TabIndex = 104
        Me.TextBox_CSF_Chainage.Text = "F"
        '
        'TextBox_elevation
        '
        Me.TextBox_elevation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_elevation.Location = New System.Drawing.Point(156, 38)
        Me.TextBox_elevation.Name = "TextBox_elevation"
        Me.TextBox_elevation.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_elevation.TabIndex = 104
        Me.TextBox_elevation.Text = "E"
        '
        'TextBox_north_intersection
        '
        Me.TextBox_north_intersection.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_north_intersection.Location = New System.Drawing.Point(104, 38)
        Me.TextBox_north_intersection.Name = "TextBox_north_intersection"
        Me.TextBox_north_intersection.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_north_intersection.TabIndex = 104
        Me.TextBox_north_intersection.Text = "D"
        '
        'TextBox_east_intersection
        '
        Me.TextBox_east_intersection.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_east_intersection.Location = New System.Drawing.Point(58, 38)
        Me.TextBox_east_intersection.Name = "TextBox_east_intersection"
        Me.TextBox_east_intersection.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_east_intersection.TabIndex = 104
        Me.TextBox_east_intersection.Text = "C"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Location = New System.Drawing.Point(8, 93)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(400, 292)
        Me.TabControl1.TabIndex = 106
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.LightGray
        Me.TabPage1.Controls.Add(Me.Button_output_to_excel_parallel_ends)
        Me.TabPage1.Controls.Add(Me.Button_output_to_excel_parallel_middle)
        Me.TabPage1.Location = New System.Drawing.Point(4, 24)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(392, 264)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Heavy Wall Calcs"
        '
        'Button_output_to_excel_parallel_ends
        '
        Me.Button_output_to_excel_parallel_ends.Location = New System.Drawing.Point(6, 122)
        Me.Button_output_to_excel_parallel_ends.Name = "Button_output_to_excel_parallel_ends"
        Me.Button_output_to_excel_parallel_ends.Size = New System.Drawing.Size(362, 82)
        Me.Button_output_to_excel_parallel_ends.TabIndex = 109
        Me.Button_output_to_excel_parallel_ends.Text = "Output to Excel Grid 3D Chainage CSF Chainage " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of a segment (start - end) drafte" &
    "d parallel with the CL" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Calculates by keeping the ends of the segment " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(result " &
    "is not matching the segment length)"
        Me.Button_output_to_excel_parallel_ends.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.LightGray
        Me.TabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage2.Controls.Add(Me.Button_2dgrid)
        Me.TabPage2.Controls.Add(Me.Button_output_to_excel_top)
        Me.TabPage2.Controls.Add(Me.Button_CSF_TO_AUTOCAD)
        Me.TabPage2.Controls.Add(Me.Button_3d_2d)
        Me.TabPage2.Controls.Add(Me.Button_pt_to_CSF)
        Me.TabPage2.Location = New System.Drawing.Point(4, 24)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(392, 264)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Rest"
        '
        'Button_3d_2d
        '
        Me.Button_3d_2d.Location = New System.Drawing.Point(208, 163)
        Me.Button_3d_2d.Name = "Button_3d_2d"
        Me.Button_3d_2d.Size = New System.Drawing.Size(174, 91)
        Me.Button_3d_2d.TabIndex = 109
        Me.Button_3d_2d.Text = "Output to Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3d station" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of a 2D point" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "from Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(no CSF)"
        Me.Button_3d_2d.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.LightGray
        Me.TabPage3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage3.Controls.Add(Me.Button_start_end_without_CSF)
        Me.TabPage3.Controls.Add(Me.Button_rerouteMP)
        Me.TabPage3.Controls.Add(Me.Button_read_sta_wr_2xl)
        Me.TabPage3.Location = New System.Drawing.Point(4, 24)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(392, 264)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Rest2"
        '
        'Button_start_end_without_CSF
        '
        Me.Button_start_end_without_CSF.Location = New System.Drawing.Point(8, 6)
        Me.Button_start_end_without_CSF.Name = "Button_start_end_without_CSF"
        Me.Button_start_end_without_CSF.Size = New System.Drawing.Size(210, 82)
        Me.Button_start_end_without_CSF.TabIndex = 109
        Me.Button_start_end_without_CSF.Text = "Output to Excel Station " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "of a segment (start - end) " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "drafted parallel with the " &
    "CL" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Button_start_end_without_CSF.UseVisualStyleBackColor = True
        '
        'Button_rerouteMP
        '
        Me.Button_rerouteMP.Location = New System.Drawing.Point(21, 124)
        Me.Button_rerouteMP.Name = "Button_rerouteMP"
        Me.Button_rerouteMP.Size = New System.Drawing.Size(125, 113)
        Me.Button_rerouteMP.TabIndex = 109
        Me.Button_rerouteMP.Text = "REROUTE MP" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "writes back to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Excel X,Y,Z" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "each 0.1 mp" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "along reroute"
        Me.Button_rerouteMP.UseVisualStyleBackColor = True
        '
        'Button_read_sta_wr_2xl
        '
        Me.Button_read_sta_wr_2xl.Location = New System.Drawing.Point(257, 6)
        Me.Button_read_sta_wr_2xl.Name = "Button_read_sta_wr_2xl"
        Me.Button_read_sta_wr_2xl.Size = New System.Drawing.Size(125, 113)
        Me.Button_read_sta_wr_2xl.TabIndex = 109
        Me.Button_read_sta_wr_2xl.Text = "Reads Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "station" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "writes back to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Excel X,Y,Z"
        Me.Button_read_sta_wr_2xl.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.LightGray
        Me.TabPage4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage4.Controls.Add(Me.Label15)
        Me.TabPage4.Controls.Add(Me.Button_read_write_csf)
        Me.TabPage4.Controls.Add(Me.Label11)
        Me.TabPage4.Controls.Add(Me.Label8)
        Me.TabPage4.Controls.Add(Me.Label14)
        Me.TabPage4.Controls.Add(Me.Button_reroute_projection)
        Me.TabPage4.Controls.Add(Me.TextBox_s1)
        Me.TabPage4.Controls.Add(Me.TextBoxx1)
        Me.TabPage4.Controls.Add(Me.TextBox_e1)
        Me.TabPage4.Controls.Add(Me.TextBoxy1)
        Me.TabPage4.Controls.Add(Me.TextBox_X)
        Me.TabPage4.Controls.Add(Me.TextBoxs)
        Me.TabPage4.Controls.Add(Me.TextBoxs1)
        Me.TabPage4.Controls.Add(Me.TextBoxz1)
        Me.TabPage4.Controls.Add(Me.TextBox_Y)
        Me.TabPage4.Controls.Add(Me.Label10)
        Me.TabPage4.Controls.Add(Me.TextBox_Z)
        Me.TabPage4.Controls.Add(Me.TextBoxCSF2)
        Me.TabPage4.Controls.Add(Me.TextBoxcsf1)
        Me.TabPage4.Controls.Add(Me.Label13)
        Me.TabPage4.Controls.Add(Me.Label12)
        Me.TabPage4.Controls.Add(Me.Label16)
        Me.TabPage4.Controls.Add(Me.Label4)
        Me.TabPage4.Controls.Add(Me.Label9)
        Me.TabPage4.Controls.Add(Me.TextBox_CSF_STA)
        Me.TabPage4.Controls.Add(Me.Label6)
        Me.TabPage4.Controls.Add(Me.Label5)
        Me.TabPage4.Controls.Add(Me.Label7)
        Me.TabPage4.Location = New System.Drawing.Point(4, 24)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(392, 264)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Rest3"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(277, 77)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(56, 15)
        Me.Label15.TabIndex = 6
        Me.Label15.Text = "End Row"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button_read_write_csf
        '
        Me.Button_read_write_csf.Location = New System.Drawing.Point(6, 206)
        Me.Button_read_write_csf.Name = "Button_read_write_csf"
        Me.Button_read_write_csf.Size = New System.Drawing.Size(265, 48)
        Me.Button_read_write_csf.TabIndex = 110
        Me.Button_read_write_csf.Text = "read CSF from sheet below and " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "write csf in sheet above"
        Me.Button_read_write_csf.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(189, 163)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(13, 15)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "y"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(129, 6)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(13, 15)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "y"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(267, 51)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(63, 15)
        Me.Label14.TabIndex = 6
        Me.Label14.Text = "Start Row"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button_reroute_projection
        '
        Me.Button_reroute_projection.Location = New System.Drawing.Point(8, 49)
        Me.Button_reroute_projection.Name = "Button_reroute_projection"
        Me.Button_reroute_projection.Size = New System.Drawing.Size(210, 110)
        Me.Button_reroute_projection.TabIndex = 109
        Me.Button_reroute_projection.Text = "Reads Excel point" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "projects it to a 2d poly" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and keeps the elevation, " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "calculate" &
    "s the intermmediary" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "nodes elevation for the 2d poly" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "writes results back to exc" &
    "el"
        Me.Button_reroute_projection.UseVisualStyleBackColor = True
        '
        'TextBox_s1
        '
        Me.TextBox_s1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_s1.Location = New System.Drawing.Point(339, 49)
        Me.TextBox_s1.Name = "TextBox_s1"
        Me.TextBox_s1.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_s1.TabIndex = 104
        Me.TextBox_s1.Text = "2"
        '
        'TextBoxx1
        '
        Me.TextBoxx1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxx1.Location = New System.Drawing.Point(124, 180)
        Me.TextBoxx1.Name = "TextBoxx1"
        Me.TextBoxx1.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxx1.TabIndex = 104
        Me.TextBoxx1.Text = "E"
        Me.TextBoxx1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_e1
        '
        Me.TextBox_e1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_e1.Location = New System.Drawing.Point(339, 75)
        Me.TextBox_e1.Name = "TextBox_e1"
        Me.TextBox_e1.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_e1.TabIndex = 104
        Me.TextBox_e1.Text = "2"
        '
        'TextBoxy1
        '
        Me.TextBoxy1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxy1.Location = New System.Drawing.Point(181, 180)
        Me.TextBoxy1.Name = "TextBoxy1"
        Me.TextBoxy1.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxy1.TabIndex = 104
        Me.TextBoxy1.Text = "F"
        Me.TextBoxy1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_X
        '
        Me.TextBox_X.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_X.Location = New System.Drawing.Point(64, 23)
        Me.TextBox_X.Name = "TextBox_X"
        Me.TextBox_X.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_X.TabIndex = 104
        Me.TextBox_X.Text = "B"
        Me.TextBox_X.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxs
        '
        Me.TextBoxs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxs.Location = New System.Drawing.Point(296, 185)
        Me.TextBoxs.Name = "TextBoxs"
        Me.TextBoxs.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxs.TabIndex = 104
        Me.TextBoxs.Text = "4"
        Me.TextBoxs.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxs1
        '
        Me.TextBoxs1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxs1.Location = New System.Drawing.Point(296, 237)
        Me.TextBoxs1.Name = "TextBoxs1"
        Me.TextBoxs1.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxs1.TabIndex = 104
        Me.TextBoxs1.Text = "3"
        Me.TextBoxs1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxz1
        '
        Me.TextBoxz1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxz1.Location = New System.Drawing.Point(238, 180)
        Me.TextBoxz1.Name = "TextBoxz1"
        Me.TextBoxz1.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxz1.TabIndex = 104
        Me.TextBoxz1.Text = "G"
        Me.TextBoxz1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_Y
        '
        Me.TextBox_Y.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Y.Location = New System.Drawing.Point(121, 23)
        Me.TextBox_Y.Name = "TextBox_Y"
        Me.TextBox_Y.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_Y.TabIndex = 104
        Me.TextBox_Y.Text = "C"
        Me.TextBox_Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(247, 163)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(13, 15)
        Me.Label10.TabIndex = 6
        Me.Label10.Text = "z"
        '
        'TextBox_Z
        '
        Me.TextBox_Z.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Z.Location = New System.Drawing.Point(178, 23)
        Me.TextBox_Z.Name = "TextBox_Z"
        Me.TextBox_Z.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_Z.TabIndex = 104
        Me.TextBox_Z.Text = "D"
        Me.TextBox_Z.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxCSF2
        '
        Me.TextBoxCSF2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxCSF2.Location = New System.Drawing.Point(62, 180)
        Me.TextBoxCSF2.Name = "TextBoxCSF2"
        Me.TextBoxCSF2.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxCSF2.TabIndex = 104
        Me.TextBoxCSF2.Text = "C"
        Me.TextBoxCSF2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxcsf1
        '
        Me.TextBoxcsf1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxcsf1.Location = New System.Drawing.Point(8, 180)
        Me.TextBoxcsf1.Name = "TextBoxcsf1"
        Me.TextBoxcsf1.Size = New System.Drawing.Size(34, 20)
        Me.TextBoxcsf1.TabIndex = 104
        Me.TextBoxcsf1.Text = "A"
        Me.TextBoxcsf1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(279, 219)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 15)
        Me.Label13.TabIndex = 6
        Me.Label13.Text = "sheet no (bellow)"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(282, 167)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(101, 15)
        Me.Label12.TabIndex = 6
        Me.Label12.Text = "sheet no (above)"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(59, 163)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(36, 15)
        Me.Label16.TabIndex = 6
        Me.Label16.Text = "CSF2"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(187, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(13, 15)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "z"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 163)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(36, 15)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "CSF1"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_CSF_STA
        '
        Me.TextBox_CSF_STA.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_CSF_STA.Location = New System.Drawing.Point(7, 23)
        Me.TextBox_CSF_STA.Name = "TextBox_CSF_STA"
        Me.TextBox_CSF_STA.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_CSF_STA.TabIndex = 104
        Me.TextBox_CSF_STA.Text = "A"
        Me.TextBox_CSF_STA.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(135, 163)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(14, 15)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "x"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 6)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(54, 15)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "CSF STA"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(75, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(14, 15)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "x"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.TextBox_start)
        Me.Panel1.Controls.Add(Me.TextBox_end)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 391)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(400, 68)
        Me.Panel1.TabIndex = 105
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Label25)
        Me.Panel3.Controls.Add(Me.TextBox_east_intersection)
        Me.Panel3.Controls.Add(Me.TextBox_north_intersection)
        Me.Panel3.Controls.Add(Me.TextBox_elevation)
        Me.Panel3.Controls.Add(Me.TextBox_chainage)
        Me.Panel3.Controls.Add(Me.Label26)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.Label21)
        Me.Panel3.Controls.Add(Me.Label22)
        Me.Panel3.Controls.Add(Me.TextBox_CSF_Chainage)
        Me.Panel3.Controls.Add(Me.Label24)
        Me.Panel3.Location = New System.Drawing.Point(16, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(392, 75)
        Me.Panel3.TabIndex = 105
        '
        'TabPage5
        '
        Me.TabPage5.BackColor = System.Drawing.Color.LightGray
        Me.TabPage5.Controls.Add(Me.Button_gen_station)
        Me.TabPage5.Location = New System.Drawing.Point(4, 24)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(392, 264)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Rest4"
        '
        'Button_gen_station
        '
        Me.Button_gen_station.Location = New System.Drawing.Point(6, 6)
        Me.Button_gen_station.Name = "Button_gen_station"
        Me.Button_gen_station.Size = New System.Drawing.Size(173, 57)
        Me.Button_gen_station.TabIndex = 0
        Me.Button_gen_station.Text = "Reads 3d point FROM XL" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Select CL" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Output chainage TO XL"
        Me.Button_gen_station.UseVisualStyleBackColor = True
        '
        'heavy_wall_csf_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(416, 470)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "heavy_wall_csf_Form"
        Me.Text = "Heavy wall_positions on CL"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button_output_to_excel_top As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TextBox_chainage As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_elevation As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_north_intersection As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_east_intersection As System.Windows.Forms.TextBox
    Friend WithEvents Button_output_to_excel_parallel_middle As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_CSF_Chainage As System.Windows.Forms.TextBox
    Friend WithEvents Button_CSF_TO_AUTOCAD As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_end As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_start As System.Windows.Forms.TextBox
    Friend WithEvents Button_pt_to_CSF As System.Windows.Forms.Button
    Friend WithEvents Button_2dgrid As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Button_output_to_excel_parallel_ends As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button_3d_2d As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Button_start_end_without_CSF As System.Windows.Forms.Button
    Friend WithEvents Button_read_sta_wr_2xl As Windows.Forms.Button
    Friend WithEvents Button_rerouteMP As Windows.Forms.Button
    Friend WithEvents TabPage4 As Windows.Forms.TabPage
    Friend WithEvents Button_reroute_projection As Windows.Forms.Button
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents TextBox_X As Windows.Forms.TextBox
    Friend WithEvents TextBox_Y As Windows.Forms.TextBox
    Friend WithEvents TextBox_Z As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents TextBox_CSF_STA As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents Button_read_write_csf As Windows.Forms.Button
    Friend WithEvents Label11 As Windows.Forms.Label
    Friend WithEvents TextBoxx1 As Windows.Forms.TextBox
    Friend WithEvents TextBoxy1 As Windows.Forms.TextBox
    Friend WithEvents TextBoxs As Windows.Forms.TextBox
    Friend WithEvents TextBoxs1 As Windows.Forms.TextBox
    Friend WithEvents TextBoxz1 As Windows.Forms.TextBox
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents TextBoxcsf1 As Windows.Forms.TextBox
    Friend WithEvents Label13 As Windows.Forms.Label
    Friend WithEvents Label12 As Windows.Forms.Label
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Label15 As Windows.Forms.Label
    Friend WithEvents TextBox_s1 As Windows.Forms.TextBox
    Friend WithEvents TextBox_e1 As Windows.Forms.TextBox
    Friend WithEvents Label14 As Windows.Forms.Label
    Friend WithEvents TextBoxCSF2 As Windows.Forms.TextBox
    Friend WithEvents Label16 As Windows.Forms.Label
    Friend WithEvents TabPage5 As Windows.Forms.TabPage
    Friend WithEvents Button_gen_station As Windows.Forms.Button
End Class
