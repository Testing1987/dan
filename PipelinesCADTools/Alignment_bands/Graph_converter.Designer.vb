<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Graph_converter
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
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.TextBox_Hscale = New System.Windows.Forms.TextBox()
        Me.TextBox_Vscale = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.TextBox_Hincr = New System.Windows.Forms.TextBox()
        Me.Label_graph_lowest_el = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TextBox_H_Elevation = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.TextBox_L_elevation = New System.Windows.Forms.TextBox()
        Me.TextBox_Minimum_chainage = New System.Windows.Forms.TextBox()
        Me.TextBox_Maximum_chainage = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.TextBox_Vincr = New System.Windows.Forms.TextBox()
        Me.Panel_formating = New System.Windows.Forms.Panel()
        Me.Button_load_parameters = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox_layer_profile_polyline = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layer_text = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layer_grid_lines = New System.Windows.Forms.ComboBox()
        Me.ComboBox_text_styles = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TextBox_text_height = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Button_load_equations_from_excel = New System.Windows.Forms.Button()
        Me.Panel18 = New System.Windows.Forms.Panel()
        Me.TextBox_Row_End_eq = New System.Windows.Forms.TextBox()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.TextBox_Row_Start_eq = New System.Windows.Forms.TextBox()
        Me.Panel19 = New System.Windows.Forms.Panel()
        Me.TextBox_col_statation_ahead = New System.Windows.Forms.TextBox()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.TextBox_col_station_back = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button_draw_NEW_GRAPH = New System.Windows.Forms.Button()
        Me.Button_blocks_to_excel = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button_split_load_from_excel = New System.Windows.Forms.Button()
        Me.TextBox_split_label = New System.Windows.Forms.TextBox()
        Me.TextBox_split_sta2 = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox_split_row_end = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_split_row_start = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_split_sta1 = New System.Windows.Forms.TextBox()
        Me.Button_split_graph = New System.Windows.Forms.Button()
        Me.Panel6.SuspendLayout()
        Me.Panel_formating.SuspendLayout()
        Me.Panel18.SuspendLayout()
        Me.Panel19.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Controls.Add(Me.Label28)
        Me.Panel6.Controls.Add(Me.TextBox_Hscale)
        Me.Panel6.Controls.Add(Me.TextBox_Vscale)
        Me.Panel6.Controls.Add(Me.Label29)
        Me.Panel6.Controls.Add(Me.Label20)
        Me.Panel6.Controls.Add(Me.Label21)
        Me.Panel6.Controls.Add(Me.TextBox_Hincr)
        Me.Panel6.Controls.Add(Me.Label_graph_lowest_el)
        Me.Panel6.Controls.Add(Me.Label23)
        Me.Panel6.Controls.Add(Me.TextBox_H_Elevation)
        Me.Panel6.Controls.Add(Me.Label24)
        Me.Panel6.Controls.Add(Me.TextBox_L_elevation)
        Me.Panel6.Controls.Add(Me.TextBox_Minimum_chainage)
        Me.Panel6.Controls.Add(Me.TextBox_Maximum_chainage)
        Me.Panel6.Controls.Add(Me.Label26)
        Me.Panel6.Controls.Add(Me.TextBox_Vincr)
        Me.Panel6.Location = New System.Drawing.Point(12, 12)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(283, 258)
        Me.Panel6.TabIndex = 127
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label28.Location = New System.Drawing.Point(3, 11)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(94, 14)
        Me.Label28.TabIndex = 124
        Me.Label28.Text = "Horizontal Scale"
        '
        'TextBox_Hscale
        '
        Me.TextBox_Hscale.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TextBox_Hscale.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Hscale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Hscale.Location = New System.Drawing.Point(185, 6)
        Me.TextBox_Hscale.Name = "TextBox_Hscale"
        Me.TextBox_Hscale.ReadOnly = True
        Me.TextBox_Hscale.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_Hscale.TabIndex = 7
        Me.TextBox_Hscale.Text = "1"
        '
        'TextBox_Vscale
        '
        Me.TextBox_Vscale.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TextBox_Vscale.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Vscale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Vscale.Location = New System.Drawing.Point(185, 36)
        Me.TextBox_Vscale.Name = "TextBox_Vscale"
        Me.TextBox_Vscale.ReadOnly = True
        Me.TextBox_Vscale.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_Vscale.TabIndex = 8
        Me.TextBox_Vscale.Text = "1"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label29.Location = New System.Drawing.Point(3, 194)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(136, 14)
        Me.Label29.TabIndex = 120
        Me.Label29.Text = "Graph Minimum Station"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label20.Location = New System.Drawing.Point(3, 224)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(138, 14)
        Me.Label20.TabIndex = 120
        Me.Label20.Text = "Graph Maximum Station"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label21.Location = New System.Drawing.Point(3, 163)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(137, 14)
        Me.Label21.TabIndex = 117
        Me.Label21.Text = "Graph Highest Elevation"
        '
        'TextBox_Hincr
        '
        Me.TextBox_Hincr.BackColor = System.Drawing.Color.White
        Me.TextBox_Hincr.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Hincr.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Hincr.Location = New System.Drawing.Point(185, 67)
        Me.TextBox_Hincr.Name = "TextBox_Hincr"
        Me.TextBox_Hincr.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_Hincr.TabIndex = 9
        Me.TextBox_Hincr.Text = "100"
        '
        'Label_graph_lowest_el
        '
        Me.Label_graph_lowest_el.AutoSize = True
        Me.Label_graph_lowest_el.BackColor = System.Drawing.Color.Gainsboro
        Me.Label_graph_lowest_el.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label_graph_lowest_el.Location = New System.Drawing.Point(3, 133)
        Me.Label_graph_lowest_el.Name = "Label_graph_lowest_el"
        Me.Label_graph_lowest_el.Size = New System.Drawing.Size(137, 14)
        Me.Label_graph_lowest_el.TabIndex = 116
        Me.Label_graph_lowest_el.Text = "Graph Lowest Elevation"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label23.Location = New System.Drawing.Point(3, 102)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(108, 14)
        Me.Label23.TabIndex = 123
        Me.Label23.Text = "Vertical Increment"
        '
        'TextBox_H_Elevation
        '
        Me.TextBox_H_Elevation.BackColor = System.Drawing.Color.White
        Me.TextBox_H_Elevation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_H_Elevation.ForeColor = System.Drawing.Color.Black
        Me.TextBox_H_Elevation.Location = New System.Drawing.Point(185, 157)
        Me.TextBox_H_Elevation.Name = "TextBox_H_Elevation"
        Me.TextBox_H_Elevation.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_H_Elevation.TabIndex = 12
        Me.TextBox_H_Elevation.Text = "2300"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label24.Location = New System.Drawing.Point(3, 72)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(122, 14)
        Me.Label24.TabIndex = 118
        Me.Label24.Text = "Horizontal Increment"
        '
        'TextBox_L_elevation
        '
        Me.TextBox_L_elevation.BackColor = System.Drawing.Color.White
        Me.TextBox_L_elevation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_L_elevation.ForeColor = System.Drawing.Color.Black
        Me.TextBox_L_elevation.Location = New System.Drawing.Point(185, 127)
        Me.TextBox_L_elevation.Name = "TextBox_L_elevation"
        Me.TextBox_L_elevation.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_L_elevation.TabIndex = 11
        Me.TextBox_L_elevation.Text = "0"
        '
        'TextBox_Minimum_chainage
        '
        Me.TextBox_Minimum_chainage.BackColor = System.Drawing.Color.White
        Me.TextBox_Minimum_chainage.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Minimum_chainage.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Minimum_chainage.Location = New System.Drawing.Point(185, 189)
        Me.TextBox_Minimum_chainage.Name = "TextBox_Minimum_chainage"
        Me.TextBox_Minimum_chainage.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_Minimum_chainage.TabIndex = 13
        Me.TextBox_Minimum_chainage.Text = "0"
        '
        'TextBox_Maximum_chainage
        '
        Me.TextBox_Maximum_chainage.BackColor = System.Drawing.Color.White
        Me.TextBox_Maximum_chainage.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Maximum_chainage.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Maximum_chainage.Location = New System.Drawing.Point(185, 219)
        Me.TextBox_Maximum_chainage.Name = "TextBox_Maximum_chainage"
        Me.TextBox_Maximum_chainage.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_Maximum_chainage.TabIndex = 14
        Me.TextBox_Maximum_chainage.Text = "2300"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label26.Location = New System.Drawing.Point(3, 41)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(80, 14)
        Me.Label26.TabIndex = 121
        Me.Label26.Text = "Vertical Scale"
        '
        'TextBox_Vincr
        '
        Me.TextBox_Vincr.BackColor = System.Drawing.Color.White
        Me.TextBox_Vincr.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Vincr.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Vincr.Location = New System.Drawing.Point(185, 97)
        Me.TextBox_Vincr.Name = "TextBox_Vincr"
        Me.TextBox_Vincr.Size = New System.Drawing.Size(86, 21)
        Me.TextBox_Vincr.TabIndex = 10
        Me.TextBox_Vincr.Text = "100"
        '
        'Panel_formating
        '
        Me.Panel_formating.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_formating.Controls.Add(Me.Button_load_parameters)
        Me.Panel_formating.Controls.Add(Me.Label8)
        Me.Panel_formating.Controls.Add(Me.ComboBox_layer_profile_polyline)
        Me.Panel_formating.Controls.Add(Me.ComboBox_layer_text)
        Me.Panel_formating.Controls.Add(Me.ComboBox_layer_grid_lines)
        Me.Panel_formating.Controls.Add(Me.ComboBox_text_styles)
        Me.Panel_formating.Controls.Add(Me.Label16)
        Me.Panel_formating.Controls.Add(Me.TextBox_text_height)
        Me.Panel_formating.Controls.Add(Me.Label12)
        Me.Panel_formating.Controls.Add(Me.Label17)
        Me.Panel_formating.Controls.Add(Me.Label18)
        Me.Panel_formating.Controls.Add(Me.Label10)
        Me.Panel_formating.Location = New System.Drawing.Point(301, 12)
        Me.Panel_formating.Name = "Panel_formating"
        Me.Panel_formating.Size = New System.Drawing.Size(283, 258)
        Me.Panel_formating.TabIndex = 128
        '
        'Button_load_parameters
        '
        Me.Button_load_parameters.Location = New System.Drawing.Point(9, 205)
        Me.Button_load_parameters.Name = "Button_load_parameters"
        Me.Button_load_parameters.Size = New System.Drawing.Size(175, 46)
        Me.Button_load_parameters.TabIndex = 102
        Me.Button_load_parameters.Text = "Load Format parameters"
        Me.Button_load_parameters.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label8.Location = New System.Drawing.Point(6, 131)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(98, 14)
        Me.Label8.TabIndex = 100
        Me.Label8.Text = "Layer Grid Lines"
        '
        'ComboBox_layer_profile_polyline
        '
        Me.ComboBox_layer_profile_polyline.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_profile_polyline.FormattingEnabled = True
        Me.ComboBox_layer_profile_polyline.Location = New System.Drawing.Point(141, 158)
        Me.ComboBox_layer_profile_polyline.Name = "ComboBox_layer_profile_polyline"
        Me.ComboBox_layer_profile_polyline.Size = New System.Drawing.Size(135, 23)
        Me.ComboBox_layer_profile_polyline.TabIndex = 101
        Me.ComboBox_layer_profile_polyline.TabStop = False
        '
        'ComboBox_layer_text
        '
        Me.ComboBox_layer_text.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_text.FormattingEnabled = True
        Me.ComboBox_layer_text.Location = New System.Drawing.Point(141, 100)
        Me.ComboBox_layer_text.Name = "ComboBox_layer_text"
        Me.ComboBox_layer_text.Size = New System.Drawing.Size(135, 23)
        Me.ComboBox_layer_text.TabIndex = 101
        Me.ComboBox_layer_text.TabStop = False
        '
        'ComboBox_layer_grid_lines
        '
        Me.ComboBox_layer_grid_lines.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer_grid_lines.FormattingEnabled = True
        Me.ComboBox_layer_grid_lines.Location = New System.Drawing.Point(141, 129)
        Me.ComboBox_layer_grid_lines.Name = "ComboBox_layer_grid_lines"
        Me.ComboBox_layer_grid_lines.Size = New System.Drawing.Size(135, 23)
        Me.ComboBox_layer_grid_lines.TabIndex = 101
        Me.ComboBox_layer_grid_lines.TabStop = False
        '
        'ComboBox_text_styles
        '
        Me.ComboBox_text_styles.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_text_styles.FormattingEnabled = True
        Me.ComboBox_text_styles.Location = New System.Drawing.Point(141, 29)
        Me.ComboBox_text_styles.Name = "ComboBox_text_styles"
        Me.ComboBox_text_styles.Size = New System.Drawing.Size(135, 23)
        Me.ComboBox_text_styles.TabIndex = 101
        Me.ComboBox_text_styles.TabStop = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label16.Location = New System.Drawing.Point(6, 102)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 14)
        Me.Label16.TabIndex = 100
        Me.Label16.Text = "Layer Text"
        '
        'TextBox_text_height
        '
        Me.TextBox_text_height.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_text_height.Location = New System.Drawing.Point(141, 64)
        Me.TextBox_text_height.Name = "TextBox_text_height"
        Me.TextBox_text_height.Size = New System.Drawing.Size(42, 22)
        Me.TextBox_text_height.TabIndex = 20
        Me.TextBox_text_height.Text = "8"
        Me.TextBox_text_height.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label12.Location = New System.Drawing.Point(6, 71)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(68, 14)
        Me.Label12.TabIndex = 100
        Me.Label12.Text = "Text Height"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label17.Location = New System.Drawing.Point(6, 161)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(123, 14)
        Me.Label17.TabIndex = 100
        Me.Label17.Text = "Layer Profile Polyline"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label18.Location = New System.Drawing.Point(6, 5)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(150, 16)
        Me.Label18.TabIndex = 100
        Me.Label18.Text = "Graph format parameters"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label10.Location = New System.Drawing.Point(6, 31)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(60, 14)
        Me.Label10.TabIndex = 100
        Me.Label10.Text = "Text Style"
        '
        'Button_load_equations_from_excel
        '
        Me.Button_load_equations_from_excel.Location = New System.Drawing.Point(6, 173)
        Me.Button_load_equations_from_excel.Name = "Button_load_equations_from_excel"
        Me.Button_load_equations_from_excel.Size = New System.Drawing.Size(175, 46)
        Me.Button_load_equations_from_excel.TabIndex = 2006
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
        Me.Panel18.Location = New System.Drawing.Point(6, 69)
        Me.Panel18.Name = "Panel18"
        Me.Panel18.Size = New System.Drawing.Size(143, 70)
        Me.Panel18.TabIndex = 2004
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
        'Panel19
        '
        Me.Panel19.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel19.Controls.Add(Me.Button_load_equations_from_excel)
        Me.Panel19.Controls.Add(Me.TextBox_col_statation_ahead)
        Me.Panel19.Controls.Add(Me.Panel18)
        Me.Panel19.Controls.Add(Me.Label77)
        Me.Panel19.Controls.Add(Me.Label78)
        Me.Panel19.Controls.Add(Me.Label79)
        Me.Panel19.Controls.Add(Me.TextBox_col_station_back)
        Me.Panel19.Location = New System.Drawing.Point(590, 45)
        Me.Panel19.Name = "Panel19"
        Me.Panel19.Size = New System.Drawing.Size(262, 225)
        Me.Panel19.TabIndex = 2005
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
        Me.Label78.Location = New System.Drawing.Point(60, 10)
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
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(590, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(262, 31)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "STATION EQUATIONS"
        '
        'Button_draw_NEW_GRAPH
        '
        Me.Button_draw_NEW_GRAPH.Location = New System.Drawing.Point(716, 281)
        Me.Button_draw_NEW_GRAPH.Name = "Button_draw_NEW_GRAPH"
        Me.Button_draw_NEW_GRAPH.Size = New System.Drawing.Size(136, 41)
        Me.Button_draw_NEW_GRAPH.TabIndex = 2006
        Me.Button_draw_NEW_GRAPH.Text = "Create profile graph"
        Me.Button_draw_NEW_GRAPH.UseVisualStyleBackColor = True
        '
        'Button_blocks_to_excel
        '
        Me.Button_blocks_to_excel.Location = New System.Drawing.Point(12, 281)
        Me.Button_blocks_to_excel.Name = "Button_blocks_to_excel"
        Me.Button_blocks_to_excel.Size = New System.Drawing.Size(183, 41)
        Me.Button_blocks_to_excel.TabIndex = 2006
        Me.Button_blocks_to_excel.Text = "Station of blocks to Excel"
        Me.Button_blocks_to_excel.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(867, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(177, 31)
        Me.Label2.TabIndex = 117
        Me.Label2.Text = "SPLIT GRAPH"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Button_split_load_from_excel)
        Me.Panel1.Controls.Add(Me.TextBox_split_label)
        Me.Panel1.Controls.Add(Me.TextBox_split_sta2)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.TextBox_split_sta1)
        Me.Panel1.Location = New System.Drawing.Point(867, 45)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(262, 225)
        Me.Panel1.TabIndex = 2005
        '
        'Button_split_load_from_excel
        '
        Me.Button_split_load_from_excel.Location = New System.Drawing.Point(6, 173)
        Me.Button_split_load_from_excel.Name = "Button_split_load_from_excel"
        Me.Button_split_load_from_excel.Size = New System.Drawing.Size(175, 46)
        Me.Button_split_load_from_excel.TabIndex = 2006
        Me.Button_split_load_from_excel.Text = "Load From Excel"
        Me.Button_split_load_from_excel.UseVisualStyleBackColor = True
        '
        'TextBox_split_label
        '
        Me.TextBox_split_label.BackColor = System.Drawing.Color.White
        Me.TextBox_split_label.ForeColor = System.Drawing.Color.Black
        Me.TextBox_split_label.Location = New System.Drawing.Point(176, 39)
        Me.TextBox_split_label.Name = "TextBox_split_label"
        Me.TextBox_split_label.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_split_label.TabIndex = 1
        Me.TextBox_split_label.Text = "C"
        Me.TextBox_split_label.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_split_sta2
        '
        Me.TextBox_split_sta2.BackColor = System.Drawing.Color.White
        Me.TextBox_split_sta2.ForeColor = System.Drawing.Color.Black
        Me.TextBox_split_sta2.Location = New System.Drawing.Point(100, 39)
        Me.TextBox_split_sta2.Name = "TextBox_split_sta2"
        Me.TextBox_split_sta2.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_split_sta2.TabIndex = 1
        Me.TextBox_split_sta2.Text = "B"
        Me.TextBox_split_sta2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_split_row_end)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.TextBox_split_row_start)
        Me.Panel2.Location = New System.Drawing.Point(6, 69)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(143, 70)
        Me.Panel2.TabIndex = 2004
        '
        'TextBox_split_row_end
        '
        Me.TextBox_split_row_end.BackColor = System.Drawing.Color.White
        Me.TextBox_split_row_end.ForeColor = System.Drawing.Color.Black
        Me.TextBox_split_row_end.Location = New System.Drawing.Point(77, 36)
        Me.TextBox_split_row_end.Name = "TextBox_split_row_end"
        Me.TextBox_split_row_end.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_split_row_end.TabIndex = 3
        Me.TextBox_split_row_end.Text = "160"
        Me.TextBox_split_row_end.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 15)
        Me.Label3.TabIndex = 300
        Me.Label3.Text = "Row Start"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 39)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 15)
        Me.Label4.TabIndex = 300
        Me.Label4.Text = "Row End"
        '
        'TextBox_split_row_start
        '
        Me.TextBox_split_row_start.BackColor = System.Drawing.Color.White
        Me.TextBox_split_row_start.ForeColor = System.Drawing.Color.Black
        Me.TextBox_split_row_start.Location = New System.Drawing.Point(77, 5)
        Me.TextBox_split_row_start.Name = "TextBox_split_row_start"
        Me.TextBox_split_row_start.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_split_row_start.TabIndex = 2
        Me.TextBox_split_row_start.Text = "2"
        Me.TextBox_split_row_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 15)
        Me.Label5.TabIndex = 300
        Me.Label5.Text = "Columns"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(173, 20)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(38, 15)
        Me.Label9.TabIndex = 300
        Me.Label9.Text = "Label"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(-3, 20)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(82, 15)
        Me.Label6.TabIndex = 300
        Me.Label6.Text = "Station Begin"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(87, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 15)
        Me.Label7.TabIndex = 300
        Me.Label7.Text = "Station End"
        '
        'TextBox_split_sta1
        '
        Me.TextBox_split_sta1.BackColor = System.Drawing.Color.White
        Me.TextBox_split_sta1.ForeColor = System.Drawing.Color.Black
        Me.TextBox_split_sta1.Location = New System.Drawing.Point(6, 38)
        Me.TextBox_split_sta1.Name = "TextBox_split_sta1"
        Me.TextBox_split_sta1.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_split_sta1.TabIndex = 0
        Me.TextBox_split_sta1.Text = "A"
        Me.TextBox_split_sta1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_split_graph
        '
        Me.Button_split_graph.Location = New System.Drawing.Point(993, 281)
        Me.Button_split_graph.Name = "Button_split_graph"
        Me.Button_split_graph.Size = New System.Drawing.Size(136, 41)
        Me.Button_split_graph.TabIndex = 2006
        Me.Button_split_graph.Text = "Split profile graph"
        Me.Button_split_graph.UseVisualStyleBackColor = True
        '
        'Graph_converter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(1145, 334)
        Me.Controls.Add(Me.Button_blocks_to_excel)
        Me.Controls.Add(Me.Button_split_graph)
        Me.Controls.Add(Me.Button_draw_NEW_GRAPH)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel19)
        Me.Controls.Add(Me.Panel_formating)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "Graph_converter"
        Me.Text = "Graph converter"
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel_formating.ResumeLayout(False)
        Me.Panel_formating.PerformLayout()
        Me.Panel18.ResumeLayout(False)
        Me.Panel18.PerformLayout()
        Me.Panel19.ResumeLayout(False)
        Me.Panel19.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Hscale As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Vscale As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Hincr As System.Windows.Forms.TextBox
    Friend WithEvents Label_graph_lowest_el As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TextBox_H_Elevation As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TextBox_L_elevation As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Minimum_chainage As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Maximum_chainage As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Vincr As System.Windows.Forms.TextBox
    Friend WithEvents Panel_formating As System.Windows.Forms.Panel
    Friend WithEvents Button_load_parameters As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_layer_profile_polyline As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layer_text As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layer_grid_lines As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_text_styles As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox_text_height As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Button_load_equations_from_excel As System.Windows.Forms.Button
    Friend WithEvents Panel18 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_Row_End_eq As System.Windows.Forms.TextBox
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents Label76 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Row_Start_eq As System.Windows.Forms.TextBox
    Friend WithEvents Panel19 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_col_statation_ahead As System.Windows.Forms.TextBox
    Friend WithEvents Label77 As System.Windows.Forms.Label
    Friend WithEvents Label78 As System.Windows.Forms.Label
    Friend WithEvents Label79 As System.Windows.Forms.Label
    Friend WithEvents TextBox_col_station_back As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_draw_NEW_GRAPH As System.Windows.Forms.Button
    Friend WithEvents Button_blocks_to_excel As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button_split_load_from_excel As System.Windows.Forms.Button
    Friend WithEvents TextBox_split_label As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_split_sta2 As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_split_row_end As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_split_row_start As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_split_sta1 As System.Windows.Forms.TextBox
    Friend WithEvents Button_split_graph As System.Windows.Forms.Button
End Class
