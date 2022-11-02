<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Transcanada_layers_form
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
        Me.Button_GENERAL_LAYERING = New System.Windows.Forms.Button()
        Me.ListBox_LAYERING_name = New System.Windows.Forms.ListBox()
        Me.Button_Load_all_layers = New System.Windows.Forms.Button()
        Me.Button_LOAD_GRUP_LAYERS = New System.Windows.Forms.Button()
        Me.Button_LOAD_ONLY_LTYPES = New System.Windows.Forms.Button()
        Me.TextBox_description = New System.Windows.Forms.TextBox()
        Me.Button_CIVIL_LAYERYING = New System.Windows.Forms.Button()
        Me.Button_ELECTRICAL_LAYERING = New System.Windows.Forms.Button()
        Me.Button_MECHANICAL_LAYERING = New System.Windows.Forms.Button()
        Me.Button_PIPELINE_LAYERING = New System.Windows.Forms.Button()
        Me.Button_MAPPING_LAYERING = New System.Windows.Forms.Button()
        Me.Button_EXTRA_LAYERING = New System.Windows.Forms.Button()
        Me.Button_LOAD_1_LAYER = New System.Windows.Forms.Button()
        Me.Button_load_from_my_list = New System.Windows.Forms.Button()
        Me.CheckBox_my_list = New System.Windows.Forms.CheckBox()
        Me.Button_CREATE_UPDATE_MY_LIST = New System.Windows.Forms.Button()
        Me.ListBox_my_list = New System.Windows.Forms.ListBox()
        Me.Button_Read_my_list = New System.Windows.Forms.Button()
        Me.Button_fix_xrefs_ltypes = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button_GENERAL_LAYERING
        '
        Me.Button_GENERAL_LAYERING.BackColor = System.Drawing.Color.DodgerBlue
        Me.Button_GENERAL_LAYERING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_GENERAL_LAYERING.Location = New System.Drawing.Point(7, 74)
        Me.Button_GENERAL_LAYERING.Name = "Button_GENERAL_LAYERING"
        Me.Button_GENERAL_LAYERING.Size = New System.Drawing.Size(147, 35)
        Me.Button_GENERAL_LAYERING.TabIndex = 100
        Me.Button_GENERAL_LAYERING.Text = "GENERAL LAYERING"
        Me.Button_GENERAL_LAYERING.UseVisualStyleBackColor = False
        '
        'ListBox_LAYERING_name
        '
        Me.ListBox_LAYERING_name.BackColor = System.Drawing.Color.White
        Me.ListBox_LAYERING_name.ForeColor = System.Drawing.Color.DarkBlue
        Me.ListBox_LAYERING_name.FormattingEnabled = True
        Me.ListBox_LAYERING_name.HorizontalScrollbar = True
        Me.ListBox_LAYERING_name.ItemHeight = 14
        Me.ListBox_LAYERING_name.Location = New System.Drawing.Point(163, 71)
        Me.ListBox_LAYERING_name.Name = "ListBox_LAYERING_name"
        Me.ListBox_LAYERING_name.ScrollAlwaysVisible = True
        Me.ListBox_LAYERING_name.Size = New System.Drawing.Size(186, 284)
        Me.ListBox_LAYERING_name.TabIndex = 101
        '
        'Button_Load_all_layers
        '
        Me.Button_Load_all_layers.BackColor = System.Drawing.Color.Brown
        Me.Button_Load_all_layers.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Load_all_layers.ForeColor = System.Drawing.Color.White
        Me.Button_Load_all_layers.Location = New System.Drawing.Point(7, 11)
        Me.Button_Load_all_layers.Name = "Button_Load_all_layers"
        Me.Button_Load_all_layers.Size = New System.Drawing.Size(147, 33)
        Me.Button_Load_all_layers.TabIndex = 100
        Me.Button_Load_all_layers.Text = "LOAD ALL LAYERS"
        Me.Button_Load_all_layers.UseVisualStyleBackColor = False
        '
        'Button_LOAD_GRUP_LAYERS
        '
        Me.Button_LOAD_GRUP_LAYERS.BackColor = System.Drawing.Color.Yellow
        Me.Button_LOAD_GRUP_LAYERS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_LOAD_GRUP_LAYERS.Location = New System.Drawing.Point(163, 33)
        Me.Button_LOAD_GRUP_LAYERS.Name = "Button_LOAD_GRUP_LAYERS"
        Me.Button_LOAD_GRUP_LAYERS.Size = New System.Drawing.Size(186, 33)
        Me.Button_LOAD_GRUP_LAYERS.TabIndex = 100
        Me.Button_LOAD_GRUP_LAYERS.Text = "LOAD GROUP LAYERS"
        Me.Button_LOAD_GRUP_LAYERS.UseVisualStyleBackColor = False
        '
        'Button_LOAD_ONLY_LTYPES
        '
        Me.Button_LOAD_ONLY_LTYPES.BackColor = System.Drawing.Color.MediumTurquoise
        Me.Button_LOAD_ONLY_LTYPES.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_LOAD_ONLY_LTYPES.Location = New System.Drawing.Point(354, 12)
        Me.Button_LOAD_ONLY_LTYPES.Name = "Button_LOAD_ONLY_LTYPES"
        Me.Button_LOAD_ONLY_LTYPES.Size = New System.Drawing.Size(144, 53)
        Me.Button_LOAD_ONLY_LTYPES.TabIndex = 100
        Me.Button_LOAD_ONLY_LTYPES.Text = "LOAD ONLY " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "TCPL STANDARDS"
        Me.Button_LOAD_ONLY_LTYPES.UseVisualStyleBackColor = False
        '
        'TextBox_description
        '
        Me.TextBox_description.BackColor = System.Drawing.Color.White
        Me.TextBox_description.ForeColor = System.Drawing.Color.Black
        Me.TextBox_description.Location = New System.Drawing.Point(354, 71)
        Me.TextBox_description.Multiline = True
        Me.TextBox_description.Name = "TextBox_description"
        Me.TextBox_description.ReadOnly = True
        Me.TextBox_description.Size = New System.Drawing.Size(377, 88)
        Me.TextBox_description.TabIndex = 102
        '
        'Button_CIVIL_LAYERYING
        '
        Me.Button_CIVIL_LAYERYING.BackColor = System.Drawing.Color.Lime
        Me.Button_CIVIL_LAYERYING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_CIVIL_LAYERYING.Location = New System.Drawing.Point(7, 115)
        Me.Button_CIVIL_LAYERYING.Name = "Button_CIVIL_LAYERYING"
        Me.Button_CIVIL_LAYERYING.Size = New System.Drawing.Size(147, 35)
        Me.Button_CIVIL_LAYERYING.TabIndex = 100
        Me.Button_CIVIL_LAYERYING.Text = "CIVIL LAYERING"
        Me.Button_CIVIL_LAYERYING.UseVisualStyleBackColor = False
        '
        'Button_ELECTRICAL_LAYERING
        '
        Me.Button_ELECTRICAL_LAYERING.BackColor = System.Drawing.Color.Lime
        Me.Button_ELECTRICAL_LAYERING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_ELECTRICAL_LAYERING.Location = New System.Drawing.Point(7, 156)
        Me.Button_ELECTRICAL_LAYERING.Name = "Button_ELECTRICAL_LAYERING"
        Me.Button_ELECTRICAL_LAYERING.Size = New System.Drawing.Size(147, 35)
        Me.Button_ELECTRICAL_LAYERING.TabIndex = 100
        Me.Button_ELECTRICAL_LAYERING.Text = "ELECTRICAL LAYERING"
        Me.Button_ELECTRICAL_LAYERING.UseVisualStyleBackColor = False
        '
        'Button_MECHANICAL_LAYERING
        '
        Me.Button_MECHANICAL_LAYERING.BackColor = System.Drawing.Color.Lime
        Me.Button_MECHANICAL_LAYERING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_MECHANICAL_LAYERING.Location = New System.Drawing.Point(7, 197)
        Me.Button_MECHANICAL_LAYERING.Name = "Button_MECHANICAL_LAYERING"
        Me.Button_MECHANICAL_LAYERING.Size = New System.Drawing.Size(147, 35)
        Me.Button_MECHANICAL_LAYERING.TabIndex = 100
        Me.Button_MECHANICAL_LAYERING.Text = "MECHANICAL LAYERING"
        Me.Button_MECHANICAL_LAYERING.UseVisualStyleBackColor = False
        '
        'Button_PIPELINE_LAYERING
        '
        Me.Button_PIPELINE_LAYERING.BackColor = System.Drawing.Color.DodgerBlue
        Me.Button_PIPELINE_LAYERING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_PIPELINE_LAYERING.Location = New System.Drawing.Point(7, 238)
        Me.Button_PIPELINE_LAYERING.Name = "Button_PIPELINE_LAYERING"
        Me.Button_PIPELINE_LAYERING.Size = New System.Drawing.Size(147, 35)
        Me.Button_PIPELINE_LAYERING.TabIndex = 100
        Me.Button_PIPELINE_LAYERING.Text = "PIPELINE LAYERING"
        Me.Button_PIPELINE_LAYERING.UseVisualStyleBackColor = False
        '
        'Button_MAPPING_LAYERING
        '
        Me.Button_MAPPING_LAYERING.BackColor = System.Drawing.Color.Lime
        Me.Button_MAPPING_LAYERING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_MAPPING_LAYERING.Location = New System.Drawing.Point(7, 279)
        Me.Button_MAPPING_LAYERING.Name = "Button_MAPPING_LAYERING"
        Me.Button_MAPPING_LAYERING.Size = New System.Drawing.Size(147, 35)
        Me.Button_MAPPING_LAYERING.TabIndex = 100
        Me.Button_MAPPING_LAYERING.Text = "MAPPING LAYERING"
        Me.Button_MAPPING_LAYERING.UseVisualStyleBackColor = False
        '
        'Button_EXTRA_LAYERING
        '
        Me.Button_EXTRA_LAYERING.BackColor = System.Drawing.Color.Lime
        Me.Button_EXTRA_LAYERING.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_EXTRA_LAYERING.Location = New System.Drawing.Point(7, 320)
        Me.Button_EXTRA_LAYERING.Name = "Button_EXTRA_LAYERING"
        Me.Button_EXTRA_LAYERING.Size = New System.Drawing.Size(147, 35)
        Me.Button_EXTRA_LAYERING.TabIndex = 100
        Me.Button_EXTRA_LAYERING.Text = "EXTRA LAYERING"
        Me.Button_EXTRA_LAYERING.UseVisualStyleBackColor = False
        '
        'Button_LOAD_1_LAYER
        '
        Me.Button_LOAD_1_LAYER.BackColor = System.Drawing.Color.Gold
        Me.Button_LOAD_1_LAYER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_LOAD_1_LAYER.Location = New System.Drawing.Point(355, 164)
        Me.Button_LOAD_1_LAYER.Name = "Button_LOAD_1_LAYER"
        Me.Button_LOAD_1_LAYER.Size = New System.Drawing.Size(143, 39)
        Me.Button_LOAD_1_LAYER.TabIndex = 100
        Me.Button_LOAD_1_LAYER.Text = "LOAD Selected LAYER"
        Me.Button_LOAD_1_LAYER.UseVisualStyleBackColor = False
        '
        'Button_load_from_my_list
        '
        Me.Button_load_from_my_list.BackColor = System.Drawing.Color.Yellow
        Me.Button_load_from_my_list.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_load_from_my_list.Location = New System.Drawing.Point(355, 312)
        Me.Button_load_from_my_list.Name = "Button_load_from_my_list"
        Me.Button_load_from_my_list.Size = New System.Drawing.Size(143, 43)
        Me.Button_load_from_my_list.TabIndex = 100
        Me.Button_load_from_my_list.Text = "LOAD LAYERS from MY_LIST.txt"
        Me.Button_load_from_my_list.UseVisualStyleBackColor = False
        '
        'CheckBox_my_list
        '
        Me.CheckBox_my_list.AutoSize = True
        Me.CheckBox_my_list.Location = New System.Drawing.Point(354, 259)
        Me.CheckBox_my_list.Name = "CheckBox_my_list"
        Me.CheckBox_my_list.Size = New System.Drawing.Size(215, 32)
        Me.CheckBox_my_list.TabIndex = 103
        Me.CheckBox_my_list.Text = "USE My layers List" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(My Documents/Blocks/My_list.txt)"
        Me.CheckBox_my_list.UseVisualStyleBackColor = True
        '
        'Button_CREATE_UPDATE_MY_LIST
        '
        Me.Button_CREATE_UPDATE_MY_LIST.BackColor = System.Drawing.Color.LemonChiffon
        Me.Button_CREATE_UPDATE_MY_LIST.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_CREATE_UPDATE_MY_LIST.Location = New System.Drawing.Point(609, 238)
        Me.Button_CREATE_UPDATE_MY_LIST.Name = "Button_CREATE_UPDATE_MY_LIST"
        Me.Button_CREATE_UPDATE_MY_LIST.Size = New System.Drawing.Size(121, 35)
        Me.Button_CREATE_UPDATE_MY_LIST.TabIndex = 100
        Me.Button_CREATE_UPDATE_MY_LIST.Text = "Create My_list.txt"
        Me.Button_CREATE_UPDATE_MY_LIST.UseVisualStyleBackColor = False
        '
        'ListBox_my_list
        '
        Me.ListBox_my_list.BackColor = System.Drawing.Color.White
        Me.ListBox_my_list.ForeColor = System.Drawing.Color.DarkBlue
        Me.ListBox_my_list.FormattingEnabled = True
        Me.ListBox_my_list.HorizontalScrollbar = True
        Me.ListBox_my_list.ItemHeight = 14
        Me.ListBox_my_list.Location = New System.Drawing.Point(737, 71)
        Me.ListBox_my_list.Name = "ListBox_my_list"
        Me.ListBox_my_list.ScrollAlwaysVisible = True
        Me.ListBox_my_list.Size = New System.Drawing.Size(186, 284)
        Me.ListBox_my_list.TabIndex = 101
        '
        'Button_Read_my_list
        '
        Me.Button_Read_my_list.BackColor = System.Drawing.Color.LemonChiffon
        Me.Button_Read_my_list.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Read_my_list.Location = New System.Drawing.Point(608, 320)
        Me.Button_Read_my_list.Name = "Button_Read_my_list"
        Me.Button_Read_my_list.Size = New System.Drawing.Size(121, 35)
        Me.Button_Read_my_list.TabIndex = 100
        Me.Button_Read_my_list.Text = "Read My_list.txt"
        Me.Button_Read_my_list.UseVisualStyleBackColor = False
        '
        'Button_fix_xrefs_ltypes
        '
        Me.Button_fix_xrefs_ltypes.BackColor = System.Drawing.Color.MediumTurquoise
        Me.Button_fix_xrefs_ltypes.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_fix_xrefs_ltypes.Location = New System.Drawing.Point(609, 11)
        Me.Button_fix_xrefs_ltypes.Name = "Button_fix_xrefs_ltypes"
        Me.Button_fix_xrefs_ltypes.Size = New System.Drawing.Size(122, 54)
        Me.Button_fix_xrefs_ltypes.TabIndex = 100
        Me.Button_fix_xrefs_ltypes.Text = "Xref Ltype" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Replace"
        Me.Button_fix_xrefs_ltypes.UseVisualStyleBackColor = False
        '
        'Transcanada_layers_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(928, 368)
        Me.Controls.Add(Me.CheckBox_my_list)
        Me.Controls.Add(Me.TextBox_description)
        Me.Controls.Add(Me.ListBox_my_list)
        Me.Controls.Add(Me.ListBox_LAYERING_name)
        Me.Controls.Add(Me.Button_fix_xrefs_ltypes)
        Me.Controls.Add(Me.Button_LOAD_ONLY_LTYPES)
        Me.Controls.Add(Me.Button_Load_all_layers)
        Me.Controls.Add(Me.Button_LOAD_1_LAYER)
        Me.Controls.Add(Me.Button_Read_my_list)
        Me.Controls.Add(Me.Button_CREATE_UPDATE_MY_LIST)
        Me.Controls.Add(Me.Button_load_from_my_list)
        Me.Controls.Add(Me.Button_LOAD_GRUP_LAYERS)
        Me.Controls.Add(Me.Button_EXTRA_LAYERING)
        Me.Controls.Add(Me.Button_MAPPING_LAYERING)
        Me.Controls.Add(Me.Button_PIPELINE_LAYERING)
        Me.Controls.Add(Me.Button_MECHANICAL_LAYERING)
        Me.Controls.Add(Me.Button_ELECTRICAL_LAYERING)
        Me.Controls.Add(Me.Button_CIVIL_LAYERYING)
        Me.Controls.Add(Me.Button_GENERAL_LAYERING)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "Transcanada_layers_form"
        Me.Text = "TCPL LAYERS"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_GENERAL_LAYERING As System.Windows.Forms.Button
    Friend WithEvents ListBox_LAYERING_name As System.Windows.Forms.ListBox
    Friend WithEvents Button_Load_all_layers As System.Windows.Forms.Button
    Friend WithEvents Button_LOAD_GRUP_LAYERS As System.Windows.Forms.Button
    Friend WithEvents Button_LOAD_ONLY_LTYPES As System.Windows.Forms.Button
    Friend WithEvents TextBox_description As System.Windows.Forms.TextBox
    Friend WithEvents Button_CIVIL_LAYERYING As System.Windows.Forms.Button
    Friend WithEvents Button_ELECTRICAL_LAYERING As System.Windows.Forms.Button
    Friend WithEvents Button_MECHANICAL_LAYERING As System.Windows.Forms.Button
    Friend WithEvents Button_PIPELINE_LAYERING As System.Windows.Forms.Button
    Friend WithEvents Button_MAPPING_LAYERING As System.Windows.Forms.Button
    Friend WithEvents Button_EXTRA_LAYERING As System.Windows.Forms.Button
    Friend WithEvents Button_LOAD_1_LAYER As System.Windows.Forms.Button
    Friend WithEvents Button_load_from_my_list As System.Windows.Forms.Button
    Friend WithEvents CheckBox_my_list As System.Windows.Forms.CheckBox
    Friend WithEvents Button_CREATE_UPDATE_MY_LIST As System.Windows.Forms.Button
    Friend WithEvents ListBox_my_list As System.Windows.Forms.ListBox
    Friend WithEvents Button_Read_my_list As System.Windows.Forms.Button
    Friend WithEvents Button_fix_xrefs_ltypes As System.Windows.Forms.Button
End Class
