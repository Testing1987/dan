<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Text_2_attributes_form
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
        Me.Panel_BLOCK = New System.Windows.Forms.Panel()
        Me.Label_DWG_BLOCK = New System.Windows.Forms.Label()
        Me.Button_LOAD_BLOCK = New System.Windows.Forms.Button()
        Me.Button_load_text = New System.Windows.Forms.Button()
        Me.Button_insert_block = New System.Windows.Forms.Button()
        Me.TextBox_X = New System.Windows.Forms.TextBox()
        Me.TextBox_Y = New System.Windows.Forms.TextBox()
        Me.ComboBox_layers = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel_layers = New System.Windows.Forms.Panel()
        Me.TextBox_scale = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel_text = New System.Windows.Forms.Panel()
        Me.Button_reddfine_block = New System.Windows.Forms.Button()
        Me.Panel_layers.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel_BLOCK
        '
        Me.Panel_BLOCK.AutoScroll = True
        Me.Panel_BLOCK.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_BLOCK.Location = New System.Drawing.Point(9, 105)
        Me.Panel_BLOCK.Name = "Panel_BLOCK"
        Me.Panel_BLOCK.Size = New System.Drawing.Size(202, 343)
        Me.Panel_BLOCK.TabIndex = 0
        '
        'Label_DWG_BLOCK
        '
        Me.Label_DWG_BLOCK.AutoSize = True
        Me.Label_DWG_BLOCK.Location = New System.Drawing.Point(6, 87)
        Me.Label_DWG_BLOCK.Name = "Label_DWG_BLOCK"
        Me.Label_DWG_BLOCK.Size = New System.Drawing.Size(45, 15)
        Me.Label_DWG_BLOCK.TabIndex = 0
        Me.Label_DWG_BLOCK.Text = "Label1"
        '
        'Button_LOAD_BLOCK
        '
        Me.Button_LOAD_BLOCK.Location = New System.Drawing.Point(9, 454)
        Me.Button_LOAD_BLOCK.Name = "Button_LOAD_BLOCK"
        Me.Button_LOAD_BLOCK.Size = New System.Drawing.Size(102, 39)
        Me.Button_LOAD_BLOCK.TabIndex = 1
        Me.Button_LOAD_BLOCK.Text = "Load Block"
        Me.Button_LOAD_BLOCK.UseVisualStyleBackColor = True
        '
        'Button_load_text
        '
        Me.Button_load_text.Location = New System.Drawing.Point(162, 454)
        Me.Button_load_text.Name = "Button_load_text"
        Me.Button_load_text.Size = New System.Drawing.Size(102, 39)
        Me.Button_load_text.TabIndex = 1
        Me.Button_load_text.Text = "Load Text"
        Me.Button_load_text.UseVisualStyleBackColor = True
        '
        'Button_insert_block
        '
        Me.Button_insert_block.Location = New System.Drawing.Point(307, 454)
        Me.Button_insert_block.Name = "Button_insert_block"
        Me.Button_insert_block.Size = New System.Drawing.Size(102, 39)
        Me.Button_insert_block.TabIndex = 1
        Me.Button_insert_block.Text = "Insert Block"
        Me.Button_insert_block.UseVisualStyleBackColor = True
        '
        'TextBox_X
        '
        Me.TextBox_X.BackColor = System.Drawing.Color.White
        Me.TextBox_X.ForeColor = System.Drawing.Color.Black
        Me.TextBox_X.Location = New System.Drawing.Point(27, 3)
        Me.TextBox_X.Name = "TextBox_X"
        Me.TextBox_X.Size = New System.Drawing.Size(65, 21)
        Me.TextBox_X.TabIndex = 2
        Me.TextBox_X.Text = "0"
        '
        'TextBox_Y
        '
        Me.TextBox_Y.BackColor = System.Drawing.Color.White
        Me.TextBox_Y.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Y.Location = New System.Drawing.Point(27, 30)
        Me.TextBox_Y.Name = "TextBox_Y"
        Me.TextBox_Y.Size = New System.Drawing.Size(65, 21)
        Me.TextBox_Y.TabIndex = 2
        Me.TextBox_Y.Text = "0"
        '
        'ComboBox_layers
        '
        Me.ComboBox_layers.BackColor = System.Drawing.Color.White
        Me.ComboBox_layers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layers.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layers.FormattingEnabled = True
        Me.ComboBox_layers.Location = New System.Drawing.Point(251, 25)
        Me.ComboBox_layers.Name = "ComboBox_layers"
        Me.ComboBox_layers.Size = New System.Drawing.Size(277, 23)
        Me.ComboBox_layers.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(14, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "x"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(13, 15)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "y"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(248, 3)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 15)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Layer"
        '
        'Panel_layers
        '
        Me.Panel_layers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_layers.Controls.Add(Me.TextBox_Y)
        Me.Panel_layers.Controls.Add(Me.ComboBox_layers)
        Me.Panel_layers.Controls.Add(Me.TextBox_scale)
        Me.Panel_layers.Controls.Add(Me.TextBox_X)
        Me.Panel_layers.Controls.Add(Me.Label2)
        Me.Panel_layers.Controls.Add(Me.Label4)
        Me.Panel_layers.Controls.Add(Me.Label3)
        Me.Panel_layers.Controls.Add(Me.Label1)
        Me.Panel_layers.Location = New System.Drawing.Point(9, 12)
        Me.Panel_layers.Name = "Panel_layers"
        Me.Panel_layers.Size = New System.Drawing.Size(545, 65)
        Me.Panel_layers.TabIndex = 0
        '
        'TextBox_scale
        '
        Me.TextBox_scale.BackColor = System.Drawing.Color.White
        Me.TextBox_scale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_scale.Location = New System.Drawing.Point(142, 22)
        Me.TextBox_scale.Name = "TextBox_scale"
        Me.TextBox_scale.Size = New System.Drawing.Size(49, 21)
        Me.TextBox_scale.TabIndex = 2
        Me.TextBox_scale.Text = "1"
        Me.TextBox_scale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(148, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 15)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Scale"
        '
        'Panel_text
        '
        Me.Panel_text.AutoScroll = True
        Me.Panel_text.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_text.Location = New System.Drawing.Point(217, 105)
        Me.Panel_text.Name = "Panel_text"
        Me.Panel_text.Size = New System.Drawing.Size(337, 343)
        Me.Panel_text.TabIndex = 0
        '
        'Button_reddfine_block
        '
        Me.Button_reddfine_block.Location = New System.Drawing.Point(452, 454)
        Me.Button_reddfine_block.Name = "Button_reddfine_block"
        Me.Button_reddfine_block.Size = New System.Drawing.Size(102, 39)
        Me.Button_reddfine_block.TabIndex = 1
        Me.Button_reddfine_block.Text = "Redefine Block"
        Me.Button_reddfine_block.UseVisualStyleBackColor = True
        '
        'Text_2_attributes_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(562, 498)
        Me.Controls.Add(Me.Panel_layers)
        Me.Controls.Add(Me.Label_DWG_BLOCK)
        Me.Controls.Add(Me.Button_reddfine_block)
        Me.Controls.Add(Me.Button_insert_block)
        Me.Controls.Add(Me.Button_load_text)
        Me.Controls.Add(Me.Button_LOAD_BLOCK)
        Me.Controls.Add(Me.Panel_text)
        Me.Controls.Add(Me.Panel_BLOCK)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Text_2_attributes_form"
        Me.Text = "Text 2 block"
        Me.Panel_layers.ResumeLayout(False)
        Me.Panel_layers.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel_BLOCK As System.Windows.Forms.Panel
    Friend WithEvents Button_LOAD_BLOCK As System.Windows.Forms.Button
    Friend WithEvents Label_DWG_BLOCK As System.Windows.Forms.Label
    Friend WithEvents Button_load_text As System.Windows.Forms.Button
    Friend WithEvents Button_insert_block As System.Windows.Forms.Button
    Friend WithEvents TextBox_X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Y As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox_layers As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel_layers As System.Windows.Forms.Panel
    Friend WithEvents TextBox_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel_text As System.Windows.Forms.Panel
    Friend WithEvents Button_reddfine_block As System.Windows.Forms.Button
End Class
