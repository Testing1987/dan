<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Block_layout_insert_Form
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
        Me.ComboBox_existing_blocks = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button_insert_block = New System.Windows.Forms.Button()
        Me.ComboBox_existing_layers = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_Layout_start = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_Layout_end = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'ComboBox_existing_blocks
        '
        Me.ComboBox_existing_blocks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_existing_blocks.FormattingEnabled = True
        Me.ComboBox_existing_blocks.Location = New System.Drawing.Point(17, 29)
        Me.ComboBox_existing_blocks.Name = "ComboBox_existing_blocks"
        Me.ComboBox_existing_blocks.Size = New System.Drawing.Size(213, 23)
        Me.ComboBox_existing_blocks.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Existing Blocks"
        '
        'Button_insert_block
        '
        Me.Button_insert_block.Location = New System.Drawing.Point(14, 198)
        Me.Button_insert_block.Name = "Button_insert_block"
        Me.Button_insert_block.Size = New System.Drawing.Size(155, 35)
        Me.Button_insert_block.TabIndex = 2
        Me.Button_insert_block.Text = "Insert Block"
        Me.Button_insert_block.UseVisualStyleBackColor = True
        '
        'ComboBox_existing_layers
        '
        Me.ComboBox_existing_layers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_existing_layers.FormattingEnabled = True
        Me.ComboBox_existing_layers.Location = New System.Drawing.Point(17, 77)
        Me.ComboBox_existing_layers.Name = "ComboBox_existing_layers"
        Me.ComboBox_existing_layers.Size = New System.Drawing.Size(213, 23)
        Me.ComboBox_existing_layers.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 15)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Existing Layers"
        '
        'TextBox_Layout_start
        '
        Me.TextBox_Layout_start.Location = New System.Drawing.Point(17, 150)
        Me.TextBox_Layout_start.Name = "TextBox_Layout_start"
        Me.TextBox_Layout_start.Size = New System.Drawing.Size(72, 21)
        Me.TextBox_Layout_start.TabIndex = 3
        Me.TextBox_Layout_start.Text = "1"
        Me.TextBox_Layout_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 132)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 15)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "From Layout"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(105, 132)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 15)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "To Layout"
        '
        'TextBox_Layout_end
        '
        Me.TextBox_Layout_end.Location = New System.Drawing.Point(97, 150)
        Me.TextBox_Layout_end.Name = "TextBox_Layout_end"
        Me.TextBox_Layout_end.Size = New System.Drawing.Size(72, 21)
        Me.TextBox_Layout_end.TabIndex = 3
        Me.TextBox_Layout_end.Text = "1"
        Me.TextBox_Layout_end.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Block_layout_insert_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(279, 242)
        Me.Controls.Add(Me.TextBox_Layout_end)
        Me.Controls.Add(Me.TextBox_Layout_start)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button_insert_block)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox_existing_layers)
        Me.Controls.Add(Me.ComboBox_existing_blocks)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MinimizeBox = False
        Me.Name = "Block_layout_insert_Form"
        Me.Text = "Insert Block"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox_existing_blocks As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_insert_block As System.Windows.Forms.Button
    Friend WithEvents ComboBox_existing_layers As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Layout_start As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Layout_end As System.Windows.Forms.TextBox
End Class
