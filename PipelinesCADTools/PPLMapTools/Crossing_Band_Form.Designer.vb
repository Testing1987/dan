<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Crossing_Band_Form
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Panel_LOAD_CL = New System.Windows.Forms.Panel()
        Me.Button_load_CL = New System.Windows.Forms.Button()
        Me.Button_draw = New System.Windows.Forms.Button()
        Me.Panel_REFRESH = New System.Windows.Forms.Panel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Panel_right_left = New System.Windows.Forms.Panel()
        Me.RadioButton_right_to_left = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Left_right = New System.Windows.Forms.RadioButton()
        Me.Panel_Load_from_excel = New System.Windows.Forms.Panel()
        Me.Button_load_From_excel = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TextBox_ROW_END = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.TextBox_ROW_START = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox_column_description = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_column_Station = New System.Windows.Forms.TextBox()
        Me.Panel_matchlines = New System.Windows.Forms.Panel()
        Me.TextBox_end = New System.Windows.Forms.TextBox()
        Me.TextBox_start = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Panel_design_param = New System.Windows.Forms.Panel()
        Me.TextBox_minimum_distance = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox_text_style = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layers_crossings = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layers_deflections = New System.Windows.Forms.ComboBox()
        Me.TextBox_textwidth = New System.Windows.Forms.TextBox()
        Me.TextBox_text_height = New System.Windows.Forms.TextBox()
        Me.TextBox_Y = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CheckBox_CALC_DEFLECTIONS = New System.Windows.Forms.CheckBox()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel_LOAD_CL.SuspendLayout()
        Me.Panel_REFRESH.SuspendLayout()
        Me.Panel_right_left.SuspendLayout()
        Me.Panel_Load_from_excel.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel_matchlines.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.Panel_design_param.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(352, 453)
        Me.TabControl1.TabIndex = 400
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage1.Controls.Add(Me.CheckBox_CALC_DEFLECTIONS)
        Me.TabPage1.Controls.Add(Me.Panel_LOAD_CL)
        Me.TabPage1.Controls.Add(Me.Button_draw)
        Me.TabPage1.Controls.Add(Me.Panel_REFRESH)
        Me.TabPage1.Controls.Add(Me.Panel_right_left)
        Me.TabPage1.Controls.Add(Me.Panel_Load_from_excel)
        Me.TabPage1.Controls.Add(Me.Panel_matchlines)
        Me.TabPage1.Location = New System.Drawing.Point(4, 24)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(344, 425)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Crossing List"
        '
        'Panel_LOAD_CL
        '
        Me.Panel_LOAD_CL.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_LOAD_CL.Controls.Add(Me.Button_load_CL)
        Me.Panel_LOAD_CL.Location = New System.Drawing.Point(6, 6)
        Me.Panel_LOAD_CL.Name = "Panel_LOAD_CL"
        Me.Panel_LOAD_CL.Size = New System.Drawing.Size(151, 50)
        Me.Panel_LOAD_CL.TabIndex = 100
        '
        'Button_load_CL
        '
        Me.Button_load_CL.Location = New System.Drawing.Point(3, 3)
        Me.Button_load_CL.Name = "Button_load_CL"
        Me.Button_load_CL.Size = New System.Drawing.Size(131, 39)
        Me.Button_load_CL.TabIndex = 200
        Me.Button_load_CL.Text = "Load Centerline"
        Me.Button_load_CL.UseVisualStyleBackColor = True
        '
        'Button_draw
        '
        Me.Button_draw.Location = New System.Drawing.Point(8, 379)
        Me.Button_draw.Name = "Button_draw"
        Me.Button_draw.Size = New System.Drawing.Size(136, 41)
        Me.Button_draw.TabIndex = 200
        Me.Button_draw.Text = "Insert band"
        Me.Button_draw.UseVisualStyleBackColor = True
        '
        'Panel_REFRESH
        '
        Me.Panel_REFRESH.BackColor = System.Drawing.Color.Wheat
        Me.Panel_REFRESH.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_REFRESH.Controls.Add(Me.Label12)
        Me.Panel_REFRESH.Location = New System.Drawing.Point(163, 4)
        Me.Panel_REFRESH.Name = "Panel_REFRESH"
        Me.Panel_REFRESH.Size = New System.Drawing.Size(168, 52)
        Me.Panel_REFRESH.TabIndex = 100
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(47, 17)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(59, 15)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "REFRESH"
        '
        'Panel_right_left
        '
        Me.Panel_right_left.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_right_left.Controls.Add(Me.RadioButton_right_to_left)
        Me.Panel_right_left.Controls.Add(Me.RadioButton_Left_right)
        Me.Panel_right_left.Location = New System.Drawing.Point(8, 319)
        Me.Panel_right_left.Name = "Panel_right_left"
        Me.Panel_right_left.Size = New System.Drawing.Size(325, 54)
        Me.Panel_right_left.TabIndex = 100
        '
        'RadioButton_right_to_left
        '
        Me.RadioButton_right_to_left.AutoSize = True
        Me.RadioButton_right_to_left.Location = New System.Drawing.Point(6, 28)
        Me.RadioButton_right_to_left.Name = "RadioButton_right_to_left"
        Me.RadioButton_right_to_left.Size = New System.Drawing.Size(93, 19)
        Me.RadioButton_right_to_left.TabIndex = 600
        Me.RadioButton_right_to_left.Text = "Right to Left"
        Me.RadioButton_right_to_left.UseVisualStyleBackColor = True
        '
        'RadioButton_Left_right
        '
        Me.RadioButton_Left_right.AutoSize = True
        Me.RadioButton_Left_right.Checked = True
        Me.RadioButton_Left_right.Location = New System.Drawing.Point(6, 3)
        Me.RadioButton_Left_right.Name = "RadioButton_Left_right"
        Me.RadioButton_Left_right.Size = New System.Drawing.Size(93, 19)
        Me.RadioButton_Left_right.TabIndex = 500
        Me.RadioButton_Left_right.TabStop = True
        Me.RadioButton_Left_right.Text = "Left to Right"
        Me.RadioButton_Left_right.UseVisualStyleBackColor = True
        '
        'Panel_Load_from_excel
        '
        Me.Panel_Load_from_excel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_Load_from_excel.Controls.Add(Me.Button_load_From_excel)
        Me.Panel_Load_from_excel.Controls.Add(Me.Panel3)
        Me.Panel_Load_from_excel.Controls.Add(Me.Panel2)
        Me.Panel_Load_from_excel.Location = New System.Drawing.Point(6, 83)
        Me.Panel_Load_from_excel.Name = "Panel_Load_from_excel"
        Me.Panel_Load_from_excel.Size = New System.Drawing.Size(325, 158)
        Me.Panel_Load_from_excel.TabIndex = 1
        '
        'Button_load_From_excel
        '
        Me.Button_load_From_excel.Location = New System.Drawing.Point(155, 65)
        Me.Button_load_From_excel.Name = "Button_load_From_excel"
        Me.Button_load_From_excel.Size = New System.Drawing.Size(163, 82)
        Me.Button_load_From_excel.TabIndex = 2000
        Me.Button_load_From_excel.Text = "Load From Excel"
        Me.Button_load_From_excel.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.TextBox_ROW_END)
        Me.Panel3.Controls.Add(Me.Label18)
        Me.Panel3.Controls.Add(Me.Label20)
        Me.Panel3.Controls.Add(Me.TextBox_ROW_START)
        Me.Panel3.Location = New System.Drawing.Point(3, 65)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(143, 80)
        Me.Panel3.TabIndex = 100
        '
        'TextBox_ROW_END
        '
        Me.TextBox_ROW_END.BackColor = System.Drawing.Color.White
        Me.TextBox_ROW_END.ForeColor = System.Drawing.Color.Black
        Me.TextBox_ROW_END.Location = New System.Drawing.Point(77, 36)
        Me.TextBox_ROW_END.Name = "TextBox_ROW_END"
        Me.TextBox_ROW_END.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_ROW_END.TabIndex = 3
        Me.TextBox_ROW_END.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(5, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(63, 15)
        Me.Label18.TabIndex = 300
        Me.Label18.Text = "Row Start"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(11, 39)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(56, 15)
        Me.Label20.TabIndex = 300
        Me.Label20.Text = "Row End"
        '
        'TextBox_ROW_START
        '
        Me.TextBox_ROW_START.BackColor = System.Drawing.Color.White
        Me.TextBox_ROW_START.ForeColor = System.Drawing.Color.Black
        Me.TextBox_ROW_START.Location = New System.Drawing.Point(77, 5)
        Me.TextBox_ROW_START.Name = "TextBox_ROW_START"
        Me.TextBox_ROW_START.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_ROW_START.TabIndex = 2
        Me.TextBox_ROW_START.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_column_description)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.TextBox_column_Station)
        Me.Panel2.Location = New System.Drawing.Point(6, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(312, 56)
        Me.Panel2.TabIndex = 100
        '
        'TextBox_column_description
        '
        Me.TextBox_column_description.BackColor = System.Drawing.Color.White
        Me.TextBox_column_description.ForeColor = System.Drawing.Color.Black
        Me.TextBox_column_description.Location = New System.Drawing.Point(174, 24)
        Me.TextBox_column_description.Name = "TextBox_column_description"
        Me.TextBox_column_description.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_column_description.TabIndex = 1
        Me.TextBox_column_description.Text = "B"
        Me.TextBox_column_description.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(23, 27)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 15)
        Me.Label7.TabIndex = 300
        Me.Label7.Text = "Columns"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(97, 6)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(47, 15)
        Me.Label9.TabIndex = 300
        Me.Label9.Text = "Station"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(161, 6)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 15)
        Me.Label11.TabIndex = 300
        Me.Label11.Text = "Description"
        '
        'TextBox_column_Station
        '
        Me.TextBox_column_Station.BackColor = System.Drawing.Color.White
        Me.TextBox_column_Station.ForeColor = System.Drawing.Color.Black
        Me.TextBox_column_Station.Location = New System.Drawing.Point(95, 24)
        Me.TextBox_column_Station.Name = "TextBox_column_Station"
        Me.TextBox_column_Station.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_column_Station.TabIndex = 0
        Me.TextBox_column_Station.Text = "A"
        Me.TextBox_column_Station.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel_matchlines
        '
        Me.Panel_matchlines.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_matchlines.Controls.Add(Me.TextBox_end)
        Me.Panel_matchlines.Controls.Add(Me.TextBox_start)
        Me.Panel_matchlines.Controls.Add(Me.Label3)
        Me.Panel_matchlines.Controls.Add(Me.Label2)
        Me.Panel_matchlines.Controls.Add(Me.Label1)
        Me.Panel_matchlines.Location = New System.Drawing.Point(6, 247)
        Me.Panel_matchlines.Name = "Panel_matchlines"
        Me.Panel_matchlines.Size = New System.Drawing.Size(325, 66)
        Me.Panel_matchlines.TabIndex = 100
        '
        'TextBox_end
        '
        Me.TextBox_end.BackColor = System.Drawing.Color.White
        Me.TextBox_end.ForeColor = System.Drawing.Color.Black
        Me.TextBox_end.Location = New System.Drawing.Point(175, 35)
        Me.TextBox_end.Name = "TextBox_end"
        Me.TextBox_end.Size = New System.Drawing.Size(100, 21)
        Me.TextBox_end.TabIndex = 5
        '
        'TextBox_start
        '
        Me.TextBox_start.BackColor = System.Drawing.Color.White
        Me.TextBox_start.ForeColor = System.Drawing.Color.Black
        Me.TextBox_start.Location = New System.Drawing.Point(34, 35)
        Me.TextBox_start.Name = "TextBox_start"
        Me.TextBox_start.Size = New System.Drawing.Size(100, 21)
        Me.TextBox_start.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(172, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 15)
        Me.Label3.TabIndex = 300
        Me.Label3.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 15)
        Me.Label2.TabIndex = 300
        Me.Label2.Text = "From"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 15)
        Me.Label1.TabIndex = 300
        Me.Label1.Text = "Matchlines:"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage2.Controls.Add(Me.Panel_design_param)
        Me.TabPage2.Location = New System.Drawing.Point(4, 24)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(344, 425)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Parameters"
        '
        'Panel_design_param
        '
        Me.Panel_design_param.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_design_param.Controls.Add(Me.TextBox_minimum_distance)
        Me.Panel_design_param.Controls.Add(Me.Label10)
        Me.Panel_design_param.Controls.Add(Me.Label15)
        Me.Panel_design_param.Controls.Add(Me.Label6)
        Me.Panel_design_param.Controls.Add(Me.Label8)
        Me.Panel_design_param.Controls.Add(Me.ComboBox_text_style)
        Me.Panel_design_param.Controls.Add(Me.ComboBox_layers_crossings)
        Me.Panel_design_param.Controls.Add(Me.ComboBox_layers_deflections)
        Me.Panel_design_param.Controls.Add(Me.TextBox_textwidth)
        Me.Panel_design_param.Controls.Add(Me.TextBox_text_height)
        Me.Panel_design_param.Controls.Add(Me.TextBox_Y)
        Me.Panel_design_param.Controls.Add(Me.Label4)
        Me.Panel_design_param.Controls.Add(Me.Label17)
        Me.Panel_design_param.Controls.Add(Me.Label16)
        Me.Panel_design_param.Controls.Add(Me.Label5)
        Me.Panel_design_param.Location = New System.Drawing.Point(6, 6)
        Me.Panel_design_param.Name = "Panel_design_param"
        Me.Panel_design_param.Size = New System.Drawing.Size(327, 403)
        Me.Panel_design_param.TabIndex = 100
        '
        'TextBox_minimum_distance
        '
        Me.TextBox_minimum_distance.BackColor = System.Drawing.Color.White
        Me.TextBox_minimum_distance.ForeColor = System.Drawing.Color.Black
        Me.TextBox_minimum_distance.Location = New System.Drawing.Point(121, 137)
        Me.TextBox_minimum_distance.Name = "TextBox_minimum_distance"
        Me.TextBox_minimum_distance.Size = New System.Drawing.Size(50, 21)
        Me.TextBox_minimum_distance.TabIndex = 14
        Me.TextBox_minimum_distance.Text = "45"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 140)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(112, 15)
        Me.Label10.TabIndex = 800
        Me.Label10.Text = "Minimum Distance"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(5, 81)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(62, 15)
        Me.Label15.TabIndex = 700
        Me.Label15.Text = "Text Style"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 220)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 15)
        Me.Label6.TabIndex = 700
        Me.Label6.Text = "Layer for crossing text"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 175)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(142, 15)
        Me.Label8.TabIndex = 700
        Me.Label8.Text = "Layer for deflection text"
        '
        'ComboBox_text_style
        '
        Me.ComboBox_text_style.BackColor = System.Drawing.Color.White
        Me.ComboBox_text_style.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_text_style.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_text_style.FormattingEnabled = True
        Me.ComboBox_text_style.Location = New System.Drawing.Point(73, 73)
        Me.ComboBox_text_style.Name = "ComboBox_text_style"
        Me.ComboBox_text_style.Size = New System.Drawing.Size(247, 23)
        Me.ComboBox_text_style.TabIndex = 12
        '
        'ComboBox_layers_crossings
        '
        Me.ComboBox_layers_crossings.BackColor = System.Drawing.Color.White
        Me.ComboBox_layers_crossings.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layers_crossings.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layers_crossings.FormattingEnabled = True
        Me.ComboBox_layers_crossings.Location = New System.Drawing.Point(73, 238)
        Me.ComboBox_layers_crossings.Name = "ComboBox_layers_crossings"
        Me.ComboBox_layers_crossings.Size = New System.Drawing.Size(247, 23)
        Me.ComboBox_layers_crossings.TabIndex = 16
        '
        'ComboBox_layers_deflections
        '
        Me.ComboBox_layers_deflections.BackColor = System.Drawing.Color.White
        Me.ComboBox_layers_deflections.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layers_deflections.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layers_deflections.FormattingEnabled = True
        Me.ComboBox_layers_deflections.Location = New System.Drawing.Point(73, 193)
        Me.ComboBox_layers_deflections.Name = "ComboBox_layers_deflections"
        Me.ComboBox_layers_deflections.Size = New System.Drawing.Size(247, 23)
        Me.ComboBox_layers_deflections.TabIndex = 15
        '
        'TextBox_textwidth
        '
        Me.TextBox_textwidth.BackColor = System.Drawing.Color.White
        Me.TextBox_textwidth.ForeColor = System.Drawing.Color.Black
        Me.TextBox_textwidth.Location = New System.Drawing.Point(122, 103)
        Me.TextBox_textwidth.Name = "TextBox_textwidth"
        Me.TextBox_textwidth.Size = New System.Drawing.Size(50, 21)
        Me.TextBox_textwidth.TabIndex = 13
        Me.TextBox_textwidth.Text = "0.8"
        '
        'TextBox_text_height
        '
        Me.TextBox_text_height.BackColor = System.Drawing.Color.White
        Me.TextBox_text_height.ForeColor = System.Drawing.Color.Black
        Me.TextBox_text_height.Location = New System.Drawing.Point(122, 46)
        Me.TextBox_text_height.Name = "TextBox_text_height"
        Me.TextBox_text_height.Size = New System.Drawing.Size(50, 21)
        Me.TextBox_text_height.TabIndex = 11
        Me.TextBox_text_height.Text = "16"
        '
        'TextBox_Y
        '
        Me.TextBox_Y.BackColor = System.Drawing.Color.White
        Me.TextBox_Y.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Y.Location = New System.Drawing.Point(122, 18)
        Me.TextBox_Y.Name = "TextBox_Y"
        Me.TextBox_Y.Size = New System.Drawing.Size(128, 21)
        Me.TextBox_Y.TabIndex = 10
        Me.TextBox_Y.Text = "3817"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 15)
        Me.Label4.TabIndex = 100
        Me.Label4.Text = "Design Parameters:"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(5, 106)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(106, 15)
        Me.Label17.TabIndex = 300
        Me.Label17.Text = "Text Width Factor"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(4, 49)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(70, 15)
        Me.Label16.TabIndex = 300
        Me.Label16.Text = "Text Height"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 15)
        Me.Label5.TabIndex = 300
        Me.Label5.Text = "Y Paperspace"
        '
        'CheckBox_CALC_DEFLECTIONS
        '
        Me.CheckBox_CALC_DEFLECTIONS.AutoSize = True
        Me.CheckBox_CALC_DEFLECTIONS.Location = New System.Drawing.Point(163, 62)
        Me.CheckBox_CALC_DEFLECTIONS.Name = "CheckBox_CALC_DEFLECTIONS"
        Me.CheckBox_CALC_DEFLECTIONS.Size = New System.Drawing.Size(151, 19)
        Me.CheckBox_CALC_DEFLECTIONS.TabIndex = 201
        Me.CheckBox_CALC_DEFLECTIONS.Text = "Calculate PI's from CL"
        Me.CheckBox_CALC_DEFLECTIONS.UseVisualStyleBackColor = True
        '
        'Crossing_Band_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(376, 474)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MinimizeBox = False
        Me.Name = "Crossing_Band_Form"
        Me.Text = "Crossing Band Form"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.Panel_LOAD_CL.ResumeLayout(False)
        Me.Panel_REFRESH.ResumeLayout(False)
        Me.Panel_REFRESH.PerformLayout()
        Me.Panel_right_left.ResumeLayout(False)
        Me.Panel_right_left.PerformLayout()
        Me.Panel_Load_from_excel.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel_matchlines.ResumeLayout(False)
        Me.Panel_matchlines.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.Panel_design_param.ResumeLayout(False)
        Me.Panel_design_param.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Panel_LOAD_CL As System.Windows.Forms.Panel
    Friend WithEvents Button_load_CL As System.Windows.Forms.Button
    Friend WithEvents Button_draw As System.Windows.Forms.Button
    Friend WithEvents Panel_REFRESH As System.Windows.Forms.Panel
    Friend WithEvents Panel_Load_from_excel As System.Windows.Forms.Panel
    Friend WithEvents Button_load_From_excel As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_column_description As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox_column_Station As System.Windows.Forms.TextBox
    Friend WithEvents Panel_matchlines As System.Windows.Forms.Panel
    Friend WithEvents TextBox_end As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_start As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Panel_design_param As System.Windows.Forms.Panel
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_text_style As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layers_deflections As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox_textwidth As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_text_height As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Y As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_ROW_END As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TextBox_ROW_START As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_layers_crossings As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_minimum_distance As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Panel_right_left As System.Windows.Forms.Panel
    Friend WithEvents RadioButton_right_to_left As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Left_right As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBox_CALC_DEFLECTIONS As System.Windows.Forms.CheckBox
End Class
