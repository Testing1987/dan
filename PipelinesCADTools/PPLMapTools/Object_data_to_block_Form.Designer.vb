<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Object_data_to_block_Form
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
        Me.ComboBox_blocks = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layers = New System.Windows.Forms.ComboBox()
        Me.Panel_od = New System.Windows.Forms.Panel()
        Me.Button_read_OD = New System.Windows.Forms.Button()
        Me.Button_insert_block = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button_load_attributes = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckBox_rotate_180 = New System.Windows.Forms.CheckBox()
        Me.TextBox_block_scale = New System.Windows.Forms.TextBox()
        Me.Label_BLOCK_SCALE = New System.Windows.Forms.Label()
        Me.CheckBox_specify_each_point = New System.Windows.Forms.CheckBox()
        Me.Panel_options = New System.Windows.Forms.Panel()
        Me.CheckBox_upper_case = New System.Windows.Forms.CheckBox()
        Me.Panel_options.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox_blocks
        '
        Me.ComboBox_blocks.BackColor = System.Drawing.Color.White
        Me.ComboBox_blocks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_blocks.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_blocks.FormattingEnabled = True
        Me.ComboBox_blocks.Location = New System.Drawing.Point(53, 21)
        Me.ComboBox_blocks.Name = "ComboBox_blocks"
        Me.ComboBox_blocks.Size = New System.Drawing.Size(128, 23)
        Me.ComboBox_blocks.TabIndex = 0
        '
        'ComboBox_layers
        '
        Me.ComboBox_layers.BackColor = System.Drawing.Color.White
        Me.ComboBox_layers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layers.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layers.FormattingEnabled = True
        Me.ComboBox_layers.Location = New System.Drawing.Point(53, 50)
        Me.ComboBox_layers.Name = "ComboBox_layers"
        Me.ComboBox_layers.Size = New System.Drawing.Size(128, 23)
        Me.ComboBox_layers.TabIndex = 0
        '
        'Panel_od
        '
        Me.Panel_od.AutoScroll = True
        Me.Panel_od.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_od.Location = New System.Drawing.Point(200, 21)
        Me.Panel_od.Name = "Panel_od"
        Me.Panel_od.Size = New System.Drawing.Size(296, 112)
        Me.Panel_od.TabIndex = 1
        '
        'Button_read_OD
        '
        Me.Button_read_OD.Location = New System.Drawing.Point(11, 139)
        Me.Button_read_OD.Name = "Button_read_OD"
        Me.Button_read_OD.Size = New System.Drawing.Size(170, 30)
        Me.Button_read_OD.TabIndex = 0
        Me.Button_read_OD.Text = "Read Object Data"
        Me.Button_read_OD.UseVisualStyleBackColor = True
        '
        'Button_insert_block
        '
        Me.Button_insert_block.Location = New System.Drawing.Point(499, 139)
        Me.Button_insert_block.Name = "Button_insert_block"
        Me.Button_insert_block.Size = New System.Drawing.Size(208, 30)
        Me.Button_insert_block.TabIndex = 0
        Me.Button_insert_block.Text = "Insert Block"
        Me.Button_insert_block.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Block"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Layer"
        '
        'Button_load_attributes
        '
        Me.Button_load_attributes.Location = New System.Drawing.Point(200, 139)
        Me.Button_load_attributes.Name = "Button_load_attributes"
        Me.Button_load_attributes.Size = New System.Drawing.Size(296, 30)
        Me.Button_load_attributes.TabIndex = 0
        Me.Button_load_attributes.Text = "Load Attributes"
        Me.Button_load_attributes.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(208, 3)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(97, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Block attributes"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(335, 3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(109, 15)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Object Data Fields"
        '
        'CheckBox_rotate_180
        '
        Me.CheckBox_rotate_180.AutoSize = True
        Me.CheckBox_rotate_180.Location = New System.Drawing.Point(3, 7)
        Me.CheckBox_rotate_180.Name = "CheckBox_rotate_180"
        Me.CheckBox_rotate_180.Size = New System.Drawing.Size(87, 19)
        Me.CheckBox_rotate_180.TabIndex = 3
        Me.CheckBox_rotate_180.Text = "Rotate 180"
        Me.CheckBox_rotate_180.UseVisualStyleBackColor = True
        '
        'TextBox_block_scale
        '
        Me.TextBox_block_scale.BackColor = System.Drawing.Color.White
        Me.TextBox_block_scale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_block_scale.Location = New System.Drawing.Point(3, 82)
        Me.TextBox_block_scale.Name = "TextBox_block_scale"
        Me.TextBox_block_scale.Size = New System.Drawing.Size(62, 21)
        Me.TextBox_block_scale.TabIndex = 0
        Me.TextBox_block_scale.Text = "1"
        Me.TextBox_block_scale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_BLOCK_SCALE
        '
        Me.Label_BLOCK_SCALE.AutoSize = True
        Me.Label_BLOCK_SCALE.Location = New System.Drawing.Point(71, 85)
        Me.Label_BLOCK_SCALE.Name = "Label_BLOCK_SCALE"
        Me.Label_BLOCK_SCALE.Size = New System.Drawing.Size(74, 15)
        Me.Label_BLOCK_SCALE.TabIndex = 2
        Me.Label_BLOCK_SCALE.Text = "Block Scale"
        '
        'CheckBox_specify_each_point
        '
        Me.CheckBox_specify_each_point.AutoSize = True
        Me.CheckBox_specify_each_point.Location = New System.Drawing.Point(3, 32)
        Me.CheckBox_specify_each_point.Name = "CheckBox_specify_each_point"
        Me.CheckBox_specify_each_point.Size = New System.Drawing.Size(183, 19)
        Me.CheckBox_specify_each_point.TabIndex = 3
        Me.CheckBox_specify_each_point.Text = "Specify each insertion point"
        Me.CheckBox_specify_each_point.UseVisualStyleBackColor = True
        '
        'Panel_options
        '
        Me.Panel_options.AutoScroll = True
        Me.Panel_options.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_options.Controls.Add(Me.Label_BLOCK_SCALE)
        Me.Panel_options.Controls.Add(Me.TextBox_block_scale)
        Me.Panel_options.Controls.Add(Me.CheckBox_rotate_180)
        Me.Panel_options.Controls.Add(Me.CheckBox_upper_case)
        Me.Panel_options.Controls.Add(Me.CheckBox_specify_each_point)
        Me.Panel_options.Location = New System.Drawing.Point(502, 21)
        Me.Panel_options.Name = "Panel_options"
        Me.Panel_options.Size = New System.Drawing.Size(205, 112)
        Me.Panel_options.TabIndex = 1
        '
        'CheckBox_upper_case
        '
        Me.CheckBox_upper_case.AutoSize = True
        Me.CheckBox_upper_case.Location = New System.Drawing.Point(3, 57)
        Me.CheckBox_upper_case.Name = "CheckBox_upper_case"
        Me.CheckBox_upper_case.Size = New System.Drawing.Size(152, 19)
        Me.CheckBox_upper_case.TabIndex = 3
        Me.CheckBox_upper_case.Text = "Change to Upper Case"
        Me.CheckBox_upper_case.UseVisualStyleBackColor = True
        '
        'Object_data_to_block_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(716, 172)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button_load_attributes)
        Me.Controls.Add(Me.Button_insert_block)
        Me.Controls.Add(Me.Button_read_OD)
        Me.Controls.Add(Me.Panel_options)
        Me.Controls.Add(Me.Panel_od)
        Me.Controls.Add(Me.ComboBox_layers)
        Me.Controls.Add(Me.ComboBox_blocks)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Object_data_to_block_Form"
        Me.Text = "OD2BLOCK"
        Me.Panel_options.ResumeLayout(False)
        Me.Panel_options.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox_blocks As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layers As System.Windows.Forms.ComboBox
    Friend WithEvents Panel_od As System.Windows.Forms.Panel
    Friend WithEvents Button_read_OD As System.Windows.Forms.Button
    Friend WithEvents Button_insert_block As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button_load_attributes As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_rotate_180 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox_block_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label_BLOCK_SCALE As System.Windows.Forms.Label
    Friend WithEvents CheckBox_specify_each_point As System.Windows.Forms.CheckBox
    Friend WithEvents Panel_options As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_upper_case As System.Windows.Forms.CheckBox
End Class
