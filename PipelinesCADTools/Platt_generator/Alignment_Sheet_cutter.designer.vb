<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Alignment_Sheet_cutter
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
        Me.TabPage_drawing_setup = New System.Windows.Forms.TabControl()
        Me.TabPage_templates = New System.Windows.Forms.TabPage()
        Me.Button_templates_drawing_template = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Button_templates_Generate_sheet = New System.Windows.Forms.Button()
        Me.TextBox_OBJECT_DATA_FIELD_NAME = New System.Windows.Forms.TextBox()
        Me.Button_read_templates = New System.Windows.Forms.Button()
        Me.TextBox_templates_file_prefix = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_template_viewport_scale = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBox_TEMPLATES_main_viewport_center_X = New System.Windows.Forms.TextBox()
        Me.TextBox_TEMPLATES_main_viewport_center_y = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Button_templates_output_folder = New System.Windows.Forms.Button()
        Me.TextBox_templates_dwt_template = New System.Windows.Forms.TextBox()
        Me.TextBox_templates_Output_Directory = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TabPage_rectangle_to_viewport = New System.Windows.Forms.TabPage()
        Me.Button_adjust_viewport = New System.Windows.Forms.Button()
        Me.Button_read_rectangle = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.TextBox_adjust_viewport_scale = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TabPageViewports = New System.Windows.Forms.TabPage()
        Me.Button_dwt_template = New System.Windows.Forms.Button()
        Me.Label_Output_Directory = New System.Windows.Forms.Label()
        Me.Button_browse_Output_Directory = New System.Windows.Forms.Button()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.TextBox_NEW_NAME_PREFIX = New System.Windows.Forms.TextBox()
        Me.TextBox_start_number = New System.Windows.Forms.TextBox()
        Me.TextBox_blockScale = New System.Windows.Forms.TextBox()
        Me.TextBox_north_arrow = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_Output_Directory = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TextBox_dwt_template = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.TextBox_north_arrow_Big_X = New System.Windows.Forms.TextBox()
        Me.TextBox_north_arrow_Big_y = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.TabPage_dwg_setup = New System.Windows.Forms.TabPage()
        Me.TextBox_matchline_length = New System.Windows.Forms.TextBox()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_main_viewport_center_X = New System.Windows.Forms.TextBox()
        Me.TextBox_main_viewport_center_Y = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_main_viewport_height = New System.Windows.Forms.TextBox()
        Me.TextBox_main_viewport_width = New System.Windows.Forms.TextBox()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Button_rename_layout = New System.Windows.Forms.Button()
        Me.Button_adjust_rectangle = New System.Windows.Forms.Button()
        Me.Button_rectangles_2Pts = New System.Windows.Forms.Button()
        Me.Button_PLACE_VIEWPORTS = New System.Windows.Forms.Button()
        Me.Button_generate_Platt = New System.Windows.Forms.Button()
        Me.TabPage_drawing_setup.SuspendLayout()
        Me.TabPage_templates.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.TabPage_rectangle_to_viewport.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.TabPageViewports.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabPage_dwg_setup.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel10.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage_drawing_setup
        '
        Me.TabPage_drawing_setup.Controls.Add(Me.TabPage_templates)
        Me.TabPage_drawing_setup.Controls.Add(Me.TabPage_rectangle_to_viewport)
        Me.TabPage_drawing_setup.Controls.Add(Me.TabPageViewports)
        Me.TabPage_drawing_setup.Controls.Add(Me.TabPage_dwg_setup)
        Me.TabPage_drawing_setup.Location = New System.Drawing.Point(3, 12)
        Me.TabPage_drawing_setup.Name = "TabPage_drawing_setup"
        Me.TabPage_drawing_setup.SelectedIndex = 0
        Me.TabPage_drawing_setup.Size = New System.Drawing.Size(444, 422)
        Me.TabPage_drawing_setup.TabIndex = 0
        '
        'TabPage_templates
        '
        Me.TabPage_templates.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage_templates.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage_templates.Controls.Add(Me.Button_templates_drawing_template)
        Me.TabPage_templates.Controls.Add(Me.Label8)
        Me.TabPage_templates.Controls.Add(Me.Label18)
        Me.TabPage_templates.Controls.Add(Me.Button_templates_Generate_sheet)
        Me.TabPage_templates.Controls.Add(Me.TextBox_OBJECT_DATA_FIELD_NAME)
        Me.TabPage_templates.Controls.Add(Me.Button_read_templates)
        Me.TabPage_templates.Controls.Add(Me.TextBox_templates_file_prefix)
        Me.TabPage_templates.Controls.Add(Me.Label16)
        Me.TabPage_templates.Controls.Add(Me.Panel1)
        Me.TabPage_templates.Controls.Add(Me.Panel3)
        Me.TabPage_templates.Controls.Add(Me.Button_templates_output_folder)
        Me.TabPage_templates.Controls.Add(Me.TextBox_templates_dwt_template)
        Me.TabPage_templates.Controls.Add(Me.TextBox_templates_Output_Directory)
        Me.TabPage_templates.Controls.Add(Me.Label14)
        Me.TabPage_templates.Location = New System.Drawing.Point(4, 24)
        Me.TabPage_templates.Name = "TabPage_templates"
        Me.TabPage_templates.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_templates.Size = New System.Drawing.Size(436, 394)
        Me.TabPage_templates.TabIndex = 4
        Me.TabPage_templates.Text = "Templates"
        '
        'Button_templates_drawing_template
        '
        Me.Button_templates_drawing_template.Location = New System.Drawing.Point(5, 276)
        Me.Button_templates_drawing_template.Name = "Button_templates_drawing_template"
        Me.Button_templates_drawing_template.Size = New System.Drawing.Size(37, 38)
        Me.Button_templates_drawing_template.TabIndex = 2
        Me.Button_templates_drawing_template.Text = ". ."
        Me.Button_templates_drawing_template.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label8.Location = New System.Drawing.Point(153, 143)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(266, 17)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "OBJECT DATA FIELD NAME FOR DWG NUMBER"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label18.Location = New System.Drawing.Point(6, 107)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(49, 17)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "PREFIX"
        '
        'Button_templates_Generate_sheet
        '
        Me.Button_templates_Generate_sheet.BackColor = System.Drawing.Color.Beige
        Me.Button_templates_Generate_sheet.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_templates_Generate_sheet.Location = New System.Drawing.Point(6, 336)
        Me.Button_templates_Generate_sheet.Name = "Button_templates_Generate_sheet"
        Me.Button_templates_Generate_sheet.Size = New System.Drawing.Size(183, 48)
        Me.Button_templates_Generate_sheet.TabIndex = 4
        Me.Button_templates_Generate_sheet.Text = "GENERATE SHEETS"
        Me.Button_templates_Generate_sheet.UseVisualStyleBackColor = False
        '
        'TextBox_OBJECT_DATA_FIELD_NAME
        '
        Me.TextBox_OBJECT_DATA_FIELD_NAME.Location = New System.Drawing.Point(152, 163)
        Me.TextBox_OBJECT_DATA_FIELD_NAME.Name = "TextBox_OBJECT_DATA_FIELD_NAME"
        Me.TextBox_OBJECT_DATA_FIELD_NAME.Size = New System.Drawing.Size(123, 21)
        Me.TextBox_OBJECT_DATA_FIELD_NAME.TabIndex = 11
        Me.TextBox_OBJECT_DATA_FIELD_NAME.Text = "OBJECTID"
        '
        'Button_read_templates
        '
        Me.Button_read_templates.BackColor = System.Drawing.Color.Beige
        Me.Button_read_templates.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_read_templates.Location = New System.Drawing.Point(6, 6)
        Me.Button_read_templates.Name = "Button_read_templates"
        Me.Button_read_templates.Size = New System.Drawing.Size(183, 48)
        Me.Button_read_templates.TabIndex = 4
        Me.Button_read_templates.Text = "Read Templates"
        Me.Button_read_templates.UseVisualStyleBackColor = False
        '
        'TextBox_templates_file_prefix
        '
        Me.TextBox_templates_file_prefix.Location = New System.Drawing.Point(5, 127)
        Me.TextBox_templates_file_prefix.Name = "TextBox_templates_file_prefix"
        Me.TextBox_templates_file_prefix.Size = New System.Drawing.Size(123, 21)
        Me.TextBox_templates_file_prefix.TabIndex = 11
        Me.TextBox_templates_file_prefix.Text = "000-03-51-"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label16.Location = New System.Drawing.Point(5, 183)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(102, 17)
        Me.Label16.TabIndex = 1
        Me.Label16.Text = "Output Directory"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.TextBox_template_viewport_scale)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Location = New System.Drawing.Point(195, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(105, 73)
        Me.Panel1.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Silver
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label7.Location = New System.Drawing.Point(5, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 17)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = " Viewport Scale" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_template_viewport_scale
        '
        Me.TextBox_template_viewport_scale.Location = New System.Drawing.Point(3, 42)
        Me.TextBox_template_viewport_scale.Name = "TextBox_template_viewport_scale"
        Me.TextBox_template_viewport_scale.Size = New System.Drawing.Size(92, 21)
        Me.TextBox_template_viewport_scale.TabIndex = 8
        Me.TextBox_template_viewport_scale.Text = "50"
        Me.TextBox_template_viewport_scale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label9.Location = New System.Drawing.Point(28, 22)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(36, 17)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "1"" to"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.TextBox_TEMPLATES_main_viewport_center_X)
        Me.Panel3.Controls.Add(Me.TextBox_TEMPLATES_main_viewport_center_y)
        Me.Panel3.Controls.Add(Me.Label12)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Location = New System.Drawing.Point(306, 6)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(112, 134)
        Me.Panel3.TabIndex = 10
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Silver
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label10.Location = New System.Drawing.Point(6, 7)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(78, 47)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Viewport " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Bottom Left" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Coordinates"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_TEMPLATES_main_viewport_center_X
        '
        Me.TextBox_TEMPLATES_main_viewport_center_X.Location = New System.Drawing.Point(25, 72)
        Me.TextBox_TEMPLATES_main_viewport_center_X.Name = "TextBox_TEMPLATES_main_viewport_center_X"
        Me.TextBox_TEMPLATES_main_viewport_center_X.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_TEMPLATES_main_viewport_center_X.TabIndex = 8
        Me.TextBox_TEMPLATES_main_viewport_center_X.Text = "239.9960"
        Me.TextBox_TEMPLATES_main_viewport_center_X.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_TEMPLATES_main_viewport_center_y
        '
        Me.TextBox_TEMPLATES_main_viewport_center_y.Location = New System.Drawing.Point(25, 99)
        Me.TextBox_TEMPLATES_main_viewport_center_y.Name = "TextBox_TEMPLATES_main_viewport_center_y"
        Me.TextBox_TEMPLATES_main_viewport_center_y.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_TEMPLATES_main_viewport_center_y.TabIndex = 8
        Me.TextBox_TEMPLATES_main_viewport_center_y.Text = "2300.0020"
        Me.TextBox_TEMPLATES_main_viewport_center_y.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.Silver
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label12.Location = New System.Drawing.Point(3, 75)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(17, 17)
        Me.Label12.TabIndex = 9
        Me.Label12.Text = "X"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Silver
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label13.Location = New System.Drawing.Point(4, 102)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(16, 17)
        Me.Label13.TabIndex = 9
        Me.Label13.Text = "Y"
        '
        'Button_templates_output_folder
        '
        Me.Button_templates_output_folder.Location = New System.Drawing.Point(3, 203)
        Me.Button_templates_output_folder.Name = "Button_templates_output_folder"
        Me.Button_templates_output_folder.Size = New System.Drawing.Size(37, 38)
        Me.Button_templates_output_folder.TabIndex = 2
        Me.Button_templates_output_folder.Text = ". ."
        Me.Button_templates_output_folder.UseVisualStyleBackColor = True
        '
        'TextBox_templates_dwt_template
        '
        Me.TextBox_templates_dwt_template.BackColor = System.Drawing.Color.White
        Me.TextBox_templates_dwt_template.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_templates_dwt_template.ForeColor = System.Drawing.Color.Black
        Me.TextBox_templates_dwt_template.Location = New System.Drawing.Point(45, 276)
        Me.TextBox_templates_dwt_template.Multiline = True
        Me.TextBox_templates_dwt_template.Name = "TextBox_templates_dwt_template"
        Me.TextBox_templates_dwt_template.Size = New System.Drawing.Size(386, 38)
        Me.TextBox_templates_dwt_template.TabIndex = 0
        Me.TextBox_templates_dwt_template.Text = "C:\Users\pop70694\Documents\Work Files\2017-04-12 agen\OUTPUT\TEMPLATE_CUTTER.dwg" & _
    ""
        '
        'TextBox_templates_Output_Directory
        '
        Me.TextBox_templates_Output_Directory.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_templates_Output_Directory.Location = New System.Drawing.Point(45, 203)
        Me.TextBox_templates_Output_Directory.Multiline = True
        Me.TextBox_templates_Output_Directory.Name = "TextBox_templates_Output_Directory"
        Me.TextBox_templates_Output_Directory.Size = New System.Drawing.Size(386, 38)
        Me.TextBox_templates_Output_Directory.TabIndex = 0
        Me.TextBox_templates_Output_Directory.Text = "C:\Users\pop70694\Documents\Work Files\2017-04-12 agen\OUTPUT\"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label14.Location = New System.Drawing.Point(3, 256)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(113, 17)
        Me.Label14.TabIndex = 1
        Me.Label14.Text = " Template drawing"
        '
        'TabPage_rectangle_to_viewport
        '
        Me.TabPage_rectangle_to_viewport.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage_rectangle_to_viewport.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage_rectangle_to_viewport.Controls.Add(Me.Button_adjust_viewport)
        Me.TabPage_rectangle_to_viewport.Controls.Add(Me.Button_read_rectangle)
        Me.TabPage_rectangle_to_viewport.Controls.Add(Me.Panel4)
        Me.TabPage_rectangle_to_viewport.Location = New System.Drawing.Point(4, 24)
        Me.TabPage_rectangle_to_viewport.Name = "TabPage_rectangle_to_viewport"
        Me.TabPage_rectangle_to_viewport.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_rectangle_to_viewport.Size = New System.Drawing.Size(436, 394)
        Me.TabPage_rectangle_to_viewport.TabIndex = 5
        Me.TabPage_rectangle_to_viewport.Text = "Template2Viewport"
        '
        'Button_adjust_viewport
        '
        Me.Button_adjust_viewport.BackColor = System.Drawing.Color.Beige
        Me.Button_adjust_viewport.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_adjust_viewport.Location = New System.Drawing.Point(6, 336)
        Me.Button_adjust_viewport.Name = "Button_adjust_viewport"
        Me.Button_adjust_viewport.Size = New System.Drawing.Size(183, 48)
        Me.Button_adjust_viewport.TabIndex = 4
        Me.Button_adjust_viewport.Text = "Adjust Viewport"
        Me.Button_adjust_viewport.UseVisualStyleBackColor = False
        '
        'Button_read_rectangle
        '
        Me.Button_read_rectangle.BackColor = System.Drawing.Color.Beige
        Me.Button_read_rectangle.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_read_rectangle.Location = New System.Drawing.Point(6, 6)
        Me.Button_read_rectangle.Name = "Button_read_rectangle"
        Me.Button_read_rectangle.Size = New System.Drawing.Size(183, 48)
        Me.Button_read_rectangle.TabIndex = 4
        Me.Button_read_rectangle.Text = "Read Template rectangle"
        Me.Button_read_rectangle.UseVisualStyleBackColor = False
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Label22)
        Me.Panel4.Controls.Add(Me.TextBox_adjust_viewport_scale)
        Me.Panel4.Controls.Add(Me.Label19)
        Me.Panel4.Controls.Add(Me.Label23)
        Me.Panel4.Location = New System.Drawing.Point(195, 6)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(206, 61)
        Me.Panel4.TabIndex = 10
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.BackColor = System.Drawing.Color.Silver
        Me.Label22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label22.Location = New System.Drawing.Point(5, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(98, 17)
        Me.Label22.TabIndex = 9
        Me.Label22.Text = " Viewport Scale" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_adjust_viewport_scale
        '
        Me.TextBox_adjust_viewport_scale.Location = New System.Drawing.Point(47, 21)
        Me.TextBox_adjust_viewport_scale.Name = "TextBox_adjust_viewport_scale"
        Me.TextBox_adjust_viewport_scale.Size = New System.Drawing.Size(56, 21)
        Me.TextBox_adjust_viewport_scale.TabIndex = 8
        Me.TextBox_adjust_viewport_scale.Text = "100"
        Me.TextBox_adjust_viewport_scale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.White
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label19.Location = New System.Drawing.Point(109, 24)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(31, 17)
        Me.Label19.TabIndex = 9
        Me.Label19.Text = "feet"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.BackColor = System.Drawing.Color.White
        Me.Label23.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label23.Location = New System.Drawing.Point(5, 24)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(36, 17)
        Me.Label23.TabIndex = 9
        Me.Label23.Text = "1"" to"
        '
        'TabPageViewports
        '
        Me.TabPageViewports.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPageViewports.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPageViewports.Controls.Add(Me.Button_dwt_template)
        Me.TabPageViewports.Controls.Add(Me.Label_Output_Directory)
        Me.TabPageViewports.Controls.Add(Me.Button_browse_Output_Directory)
        Me.TabPageViewports.Controls.Add(Me.Label42)
        Me.TabPageViewports.Controls.Add(Me.TextBox_NEW_NAME_PREFIX)
        Me.TabPageViewports.Controls.Add(Me.TextBox_start_number)
        Me.TabPageViewports.Controls.Add(Me.TextBox_blockScale)
        Me.TabPageViewports.Controls.Add(Me.TextBox_north_arrow)
        Me.TabPageViewports.Controls.Add(Me.Label2)
        Me.TabPageViewports.Controls.Add(Me.TextBox_Output_Directory)
        Me.TabPageViewports.Controls.Add(Me.Label3)
        Me.TabPageViewports.Controls.Add(Me.Label1)
        Me.TabPageViewports.Controls.Add(Me.Label17)
        Me.TabPageViewports.Controls.Add(Me.TextBox_dwt_template)
        Me.TabPageViewports.Controls.Add(Me.Panel2)
        Me.TabPageViewports.Location = New System.Drawing.Point(4, 24)
        Me.TabPageViewports.Name = "TabPageViewports"
        Me.TabPageViewports.Size = New System.Drawing.Size(436, 394)
        Me.TabPageViewports.TabIndex = 3
        Me.TabPageViewports.Text = "Viewports & North Arrows"
        '
        'Button_dwt_template
        '
        Me.Button_dwt_template.Location = New System.Drawing.Point(3, 81)
        Me.Button_dwt_template.Name = "Button_dwt_template"
        Me.Button_dwt_template.Size = New System.Drawing.Size(37, 38)
        Me.Button_dwt_template.TabIndex = 2
        Me.Button_dwt_template.Text = ". ."
        Me.Button_dwt_template.UseVisualStyleBackColor = True
        '
        'Label_Output_Directory
        '
        Me.Label_Output_Directory.AutoSize = True
        Me.Label_Output_Directory.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label_Output_Directory.Location = New System.Drawing.Point(3, 0)
        Me.Label_Output_Directory.Name = "Label_Output_Directory"
        Me.Label_Output_Directory.Size = New System.Drawing.Size(102, 17)
        Me.Label_Output_Directory.TabIndex = 1
        Me.Label_Output_Directory.Text = "Output Directory"
        '
        'Button_browse_Output_Directory
        '
        Me.Button_browse_Output_Directory.Location = New System.Drawing.Point(1, 20)
        Me.Button_browse_Output_Directory.Name = "Button_browse_Output_Directory"
        Me.Button_browse_Output_Directory.Size = New System.Drawing.Size(37, 38)
        Me.Button_browse_Output_Directory.TabIndex = 2
        Me.Button_browse_Output_Directory.Text = ". ."
        Me.Button_browse_Output_Directory.UseVisualStyleBackColor = True
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.BackColor = System.Drawing.Color.Gainsboro
        Me.Label42.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label42.Location = New System.Drawing.Point(6, 322)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(149, 17)
        Me.Label42.TabIndex = 12
        Me.Label42.Text = "North Arrow Block Name"
        '
        'TextBox_NEW_NAME_PREFIX
        '
        Me.TextBox_NEW_NAME_PREFIX.Location = New System.Drawing.Point(6, 142)
        Me.TextBox_NEW_NAME_PREFIX.Name = "TextBox_NEW_NAME_PREFIX"
        Me.TextBox_NEW_NAME_PREFIX.Size = New System.Drawing.Size(270, 21)
        Me.TextBox_NEW_NAME_PREFIX.TabIndex = 11
        Me.TextBox_NEW_NAME_PREFIX.Text = "P3-"
        '
        'TextBox_start_number
        '
        Me.TextBox_start_number.Location = New System.Drawing.Point(317, 142)
        Me.TextBox_start_number.Name = "TextBox_start_number"
        Me.TextBox_start_number.Size = New System.Drawing.Size(53, 21)
        Me.TextBox_start_number.TabIndex = 11
        Me.TextBox_start_number.Text = "7001"
        Me.TextBox_start_number.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_blockScale
        '
        Me.TextBox_blockScale.Location = New System.Drawing.Point(161, 348)
        Me.TextBox_blockScale.Name = "TextBox_blockScale"
        Me.TextBox_blockScale.Size = New System.Drawing.Size(53, 21)
        Me.TextBox_blockScale.TabIndex = 11
        Me.TextBox_blockScale.Text = "1"
        Me.TextBox_blockScale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_north_arrow
        '
        Me.TextBox_north_arrow.Location = New System.Drawing.Point(161, 319)
        Me.TextBox_north_arrow.Name = "TextBox_north_arrow"
        Me.TextBox_north_arrow.Size = New System.Drawing.Size(161, 21)
        Me.TextBox_north_arrow.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Location = New System.Drawing.Point(6, 122)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "PREFIX"
        '
        'TextBox_Output_Directory
        '
        Me.TextBox_Output_Directory.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Output_Directory.Location = New System.Drawing.Point(43, 20)
        Me.TextBox_Output_Directory.Multiline = True
        Me.TextBox_Output_Directory.Name = "TextBox_Output_Directory"
        Me.TextBox_Output_Directory.Size = New System.Drawing.Size(386, 38)
        Me.TextBox_Output_Directory.TabIndex = 0
        Me.TextBox_Output_Directory.Text = "G:\Spectra\357602_AccessNortheast\Drafting\016_C_Typicals\Residential Site Specif" & _
    "ics\02-Stony Point\"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Location = New System.Drawing.Point(319, 122)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 17)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Index"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(79, 348)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Block Scale"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label17.Location = New System.Drawing.Point(1, 61)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(113, 17)
        Me.Label17.TabIndex = 1
        Me.Label17.Text = " Template drawing"
        '
        'TextBox_dwt_template
        '
        Me.TextBox_dwt_template.BackColor = System.Drawing.Color.White
        Me.TextBox_dwt_template.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_dwt_template.ForeColor = System.Drawing.Color.Black
        Me.TextBox_dwt_template.Location = New System.Drawing.Point(43, 81)
        Me.TextBox_dwt_template.Multiline = True
        Me.TextBox_dwt_template.Name = "TextBox_dwt_template"
        Me.TextBox_dwt_template.Size = New System.Drawing.Size(386, 38)
        Me.TextBox_dwt_template.TabIndex = 0
        Me.TextBox_dwt_template.Text = "G:\Spectra\357602_AccessNortheast\Drafting\010_C_Geobase\Working\FOR DAN\2016-03-" & _
    "31\Residential Site Specific.dwt"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label26)
        Me.Panel2.Controls.Add(Me.TextBox_north_arrow_Big_X)
        Me.Panel2.Controls.Add(Me.TextBox_north_arrow_Big_y)
        Me.Panel2.Controls.Add(Me.Label27)
        Me.Panel2.Controls.Add(Me.Label28)
        Me.Panel2.Location = New System.Drawing.Point(9, 178)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(105, 134)
        Me.Panel2.TabIndex = 10
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.BackColor = System.Drawing.Color.Silver
        Me.Label26.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label26.Location = New System.Drawing.Point(5, 10)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(90, 47)
        Me.Label26.TabIndex = 9
        Me.Label26.Text = "North Arrow " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Main Viewport" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Coordinates"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_north_arrow_Big_X
        '
        Me.TextBox_north_arrow_Big_X.Location = New System.Drawing.Point(25, 72)
        Me.TextBox_north_arrow_Big_X.Name = "TextBox_north_arrow_Big_X"
        Me.TextBox_north_arrow_Big_X.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_north_arrow_Big_X.TabIndex = 8
        Me.TextBox_north_arrow_Big_X.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_north_arrow_Big_y
        '
        Me.TextBox_north_arrow_Big_y.Location = New System.Drawing.Point(25, 99)
        Me.TextBox_north_arrow_Big_y.Name = "TextBox_north_arrow_Big_y"
        Me.TextBox_north_arrow_Big_y.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_north_arrow_Big_y.TabIndex = 8
        Me.TextBox_north_arrow_Big_y.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.BackColor = System.Drawing.Color.Silver
        Me.Label27.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label27.Location = New System.Drawing.Point(3, 75)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(17, 17)
        Me.Label27.TabIndex = 9
        Me.Label27.Text = "X"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.BackColor = System.Drawing.Color.Silver
        Me.Label28.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label28.Location = New System.Drawing.Point(4, 102)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(16, 17)
        Me.Label28.TabIndex = 9
        Me.Label28.Text = "Y"
        '
        'TabPage_dwg_setup
        '
        Me.TabPage_dwg_setup.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage_dwg_setup.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage_dwg_setup.Controls.Add(Me.TextBox_matchline_length)
        Me.TabPage_dwg_setup.Controls.Add(Me.Panel8)
        Me.TabPage_dwg_setup.Controls.Add(Me.Panel10)
        Me.TabPage_dwg_setup.Controls.Add(Me.Label15)
        Me.TabPage_dwg_setup.Controls.Add(Me.Button_rename_layout)
        Me.TabPage_dwg_setup.Controls.Add(Me.Button_adjust_rectangle)
        Me.TabPage_dwg_setup.Controls.Add(Me.Button_rectangles_2Pts)
        Me.TabPage_dwg_setup.Controls.Add(Me.Button_PLACE_VIEWPORTS)
        Me.TabPage_dwg_setup.Controls.Add(Me.Button_generate_Platt)
        Me.TabPage_dwg_setup.Location = New System.Drawing.Point(4, 24)
        Me.TabPage_dwg_setup.Name = "TabPage_dwg_setup"
        Me.TabPage_dwg_setup.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_dwg_setup.Size = New System.Drawing.Size(436, 394)
        Me.TabPage_dwg_setup.TabIndex = 0
        Me.TabPage_dwg_setup.Text = "DWG"
        '
        'TextBox_matchline_length
        '
        Me.TextBox_matchline_length.Location = New System.Drawing.Point(175, 12)
        Me.TextBox_matchline_length.Name = "TextBox_matchline_length"
        Me.TextBox_matchline_length.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_matchline_length.TabIndex = 15
        Me.TextBox_matchline_length.Text = "4000"
        Me.TextBox_matchline_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel8
        '
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel8.Controls.Add(Me.Label4)
        Me.Panel8.Controls.Add(Me.TextBox_main_viewport_center_X)
        Me.Panel8.Controls.Add(Me.TextBox_main_viewport_center_Y)
        Me.Panel8.Controls.Add(Me.Label5)
        Me.Panel8.Controls.Add(Me.Label6)
        Me.Panel8.Location = New System.Drawing.Point(314, 103)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(112, 134)
        Me.Panel8.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Silver
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label4.Location = New System.Drawing.Point(6, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 47)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Main Viewport " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Center" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Coordinates"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_main_viewport_center_X
        '
        Me.TextBox_main_viewport_center_X.Location = New System.Drawing.Point(25, 72)
        Me.TextBox_main_viewport_center_X.Name = "TextBox_main_viewport_center_X"
        Me.TextBox_main_viewport_center_X.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_main_viewport_center_X.TabIndex = 8
        Me.TextBox_main_viewport_center_X.Text = "3747.0316"
        Me.TextBox_main_viewport_center_X.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_main_viewport_center_Y
        '
        Me.TextBox_main_viewport_center_Y.Location = New System.Drawing.Point(25, 99)
        Me.TextBox_main_viewport_center_Y.Name = "TextBox_main_viewport_center_Y"
        Me.TextBox_main_viewport_center_Y.Size = New System.Drawing.Size(70, 21)
        Me.TextBox_main_viewport_center_Y.TabIndex = 8
        Me.TextBox_main_viewport_center_Y.Text = "2766.6877"
        Me.TextBox_main_viewport_center_Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Silver
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label5.Location = New System.Drawing.Point(3, 75)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 17)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "X"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Silver
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Location = New System.Drawing.Point(4, 102)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(16, 17)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Y"
        '
        'Panel10
        '
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel10.Controls.Add(Me.Label11)
        Me.Panel10.Controls.Add(Me.TextBox_main_viewport_height)
        Me.Panel10.Controls.Add(Me.TextBox_main_viewport_width)
        Me.Panel10.Controls.Add(Me.Label36)
        Me.Panel10.Controls.Add(Me.Label37)
        Me.Panel10.Location = New System.Drawing.Point(6, 39)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(245, 79)
        Me.Panel10.TabIndex = 10
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Silver
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label11.Location = New System.Drawing.Point(5, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(130, 17)
        Me.Label11.TabIndex = 9
        Me.Label11.Text = "Viewport Dimensions"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_main_viewport_height
        '
        Me.TextBox_main_viewport_height.Location = New System.Drawing.Point(27, 45)
        Me.TextBox_main_viewport_height.Name = "TextBox_main_viewport_height"
        Me.TextBox_main_viewport_height.Size = New System.Drawing.Size(92, 21)
        Me.TextBox_main_viewport_height.TabIndex = 8
        Me.TextBox_main_viewport_height.Text = "1800"
        Me.TextBox_main_viewport_height.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_main_viewport_width
        '
        Me.TextBox_main_viewport_width.Location = New System.Drawing.Point(143, 45)
        Me.TextBox_main_viewport_width.Name = "TextBox_main_viewport_width"
        Me.TextBox_main_viewport_width.Size = New System.Drawing.Size(92, 21)
        Me.TextBox_main_viewport_width.TabIndex = 8
        Me.TextBox_main_viewport_width.Text = "6750.0000"
        Me.TextBox_main_viewport_width.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.BackColor = System.Drawing.Color.Silver
        Me.Label36.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label36.Location = New System.Drawing.Point(27, 25)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(45, 17)
        Me.Label36.TabIndex = 9
        Me.Label36.Text = "Height"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.BackColor = System.Drawing.Color.Silver
        Me.Label37.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label37.Location = New System.Drawing.Point(143, 25)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(42, 17)
        Me.Label37.TabIndex = 9
        Me.Label37.Text = "Width"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Silver
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label15.Location = New System.Drawing.Point(3, 14)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(166, 17)
        Me.Label15.TabIndex = 16
        Me.Label15.Text = "Length between matchlines"
        '
        'Button_rename_layout
        '
        Me.Button_rename_layout.Location = New System.Drawing.Point(279, 1)
        Me.Button_rename_layout.Name = "Button_rename_layout"
        Me.Button_rename_layout.Size = New System.Drawing.Size(147, 41)
        Me.Button_rename_layout.TabIndex = 14
        Me.Button_rename_layout.Text = "Rename current layout"
        Me.Button_rename_layout.UseVisualStyleBackColor = True
        '
        'Button_adjust_rectangle
        '
        Me.Button_adjust_rectangle.BackColor = System.Drawing.Color.Beige
        Me.Button_adjust_rectangle.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_adjust_rectangle.Location = New System.Drawing.Point(13, 227)
        Me.Button_adjust_rectangle.Name = "Button_adjust_rectangle"
        Me.Button_adjust_rectangle.Size = New System.Drawing.Size(245, 48)
        Me.Button_adjust_rectangle.TabIndex = 4
        Me.Button_adjust_rectangle.Text = "Adjust Viewport Rectangle"
        Me.Button_adjust_rectangle.UseVisualStyleBackColor = False
        '
        'Button_rectangles_2Pts
        '
        Me.Button_rectangles_2Pts.BackColor = System.Drawing.Color.Beige
        Me.Button_rectangles_2Pts.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_rectangles_2Pts.Location = New System.Drawing.Point(13, 296)
        Me.Button_rectangles_2Pts.Name = "Button_rectangles_2Pts"
        Me.Button_rectangles_2Pts.Size = New System.Drawing.Size(245, 49)
        Me.Button_rectangles_2Pts.TabIndex = 4
        Me.Button_rectangles_2Pts.Text = "Place Viewports Rectangles between 2 points"
        Me.Button_rectangles_2Pts.UseVisualStyleBackColor = False
        '
        'Button_PLACE_VIEWPORTS
        '
        Me.Button_PLACE_VIEWPORTS.BackColor = System.Drawing.Color.Beige
        Me.Button_PLACE_VIEWPORTS.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_PLACE_VIEWPORTS.Location = New System.Drawing.Point(13, 163)
        Me.Button_PLACE_VIEWPORTS.Name = "Button_PLACE_VIEWPORTS"
        Me.Button_PLACE_VIEWPORTS.Size = New System.Drawing.Size(245, 49)
        Me.Button_PLACE_VIEWPORTS.TabIndex = 4
        Me.Button_PLACE_VIEWPORTS.Text = "Place Viewports Rectangles"
        Me.Button_PLACE_VIEWPORTS.UseVisualStyleBackColor = False
        '
        'Button_generate_Platt
        '
        Me.Button_generate_Platt.BackColor = System.Drawing.Color.Beige
        Me.Button_generate_Platt.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_generate_Platt.Location = New System.Drawing.Point(279, 49)
        Me.Button_generate_Platt.Name = "Button_generate_Platt"
        Me.Button_generate_Platt.Size = New System.Drawing.Size(147, 48)
        Me.Button_generate_Platt.TabIndex = 4
        Me.Button_generate_Platt.Text = "GENERATE SHEET"
        Me.Button_generate_Platt.UseVisualStyleBackColor = False
        '
        'Alignment_Sheet_cutter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(453, 440)
        Me.Controls.Add(Me.TabPage_drawing_setup)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "Alignment_Sheet_cutter"
        Me.Text = "Property"
        Me.TabPage_drawing_setup.ResumeLayout(False)
        Me.TabPage_templates.ResumeLayout(False)
        Me.TabPage_templates.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.TabPage_rectangle_to_viewport.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.TabPageViewports.ResumeLayout(False)
        Me.TabPageViewports.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.TabPage_dwg_setup.ResumeLayout(False)
        Me.TabPage_dwg_setup.PerformLayout()
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.Panel10.ResumeLayout(False)
        Me.Panel10.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabPage_drawing_setup As System.Windows.Forms.TabControl
    Friend WithEvents TextBox_Output_Directory As System.Windows.Forms.TextBox
    Friend WithEvents Label_Output_Directory As System.Windows.Forms.Label
    Friend WithEvents Button_browse_Output_Directory As System.Windows.Forms.Button
    Friend WithEvents Button_dwt_template As System.Windows.Forms.Button
    Friend WithEvents TextBox_dwt_template As System.Windows.Forms.TextBox
    Friend WithEvents TabPageViewports As System.Windows.Forms.TabPage
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TextBox_north_arrow_Big_y As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_north_arrow_Big_X As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents TextBox_north_arrow As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TextBox_NEW_NAME_PREFIX As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_blockScale As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_start_number As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button_read_templates As System.Windows.Forms.Button
    Friend WithEvents TabPage_templates As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_template_viewport_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBox_TEMPLATES_main_viewport_center_X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_TEMPLATES_main_viewport_center_y As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Button_templates_drawing_template As System.Windows.Forms.Button
    Friend WithEvents Button_templates_Generate_sheet As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Button_templates_output_folder As System.Windows.Forms.Button
    Friend WithEvents TextBox_templates_dwt_template As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_templates_Output_Directory As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TextBox_templates_file_prefix As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox_OBJECT_DATA_FIELD_NAME As System.Windows.Forms.TextBox
    Friend WithEvents TabPage_rectangle_to_viewport As System.Windows.Forms.TabPage
    Friend WithEvents Button_adjust_viewport As System.Windows.Forms.Button
    Friend WithEvents Button_read_rectangle As System.Windows.Forms.Button
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TextBox_adjust_viewport_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TabPage_dwg_setup As System.Windows.Forms.TabPage
    Friend WithEvents TextBox_matchline_length As System.Windows.Forms.TextBox
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_main_viewport_center_X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_main_viewport_center_Y As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox_main_viewport_height As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_main_viewport_width As System.Windows.Forms.TextBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Button_rename_layout As System.Windows.Forms.Button
    Friend WithEvents Button_adjust_rectangle As System.Windows.Forms.Button
    Friend WithEvents Button_rectangles_2Pts As System.Windows.Forms.Button
    Friend WithEvents Button_PLACE_VIEWPORTS As System.Windows.Forms.Button
    Friend WithEvents Button_generate_Platt As System.Windows.Forms.Button
End Class
