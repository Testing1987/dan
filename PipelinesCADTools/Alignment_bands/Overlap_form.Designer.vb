<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Overlap_form
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
        Me.ComboBox_layer1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox_layer2 = New System.Windows.Forms.ComboBox()
        Me.ComboBox_layer3 = New System.Windows.Forms.ComboBox()
        Me.Button_analise_gaps = New System.Windows.Forms.Button()
        Me.Panel_layers = New System.Windows.Forms.Panel()
        Me.Button_add_new_combobox = New System.Windows.Forms.Button()
        Me.Button_refresh_layers = New System.Windows.Forms.Button()
        Me.Button_analise_overlapp = New System.Windows.Forms.Button()
        Me.Button_dimension_To_CL = New System.Windows.Forms.Button()
        Me.Button_done = New System.Windows.Forms.Button()
        Me.Panel_layers.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox_layer1
        '
        Me.ComboBox_layer1.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer1.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer1.FormattingEnabled = True
        Me.ComboBox_layer1.Location = New System.Drawing.Point(3, 25)
        Me.ComboBox_layer1.Name = "ComboBox_layer1"
        Me.ComboBox_layer1.Size = New System.Drawing.Size(211, 23)
        Me.ComboBox_layer1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(3, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Layers"
        '
        'ComboBox_layer2
        '
        Me.ComboBox_layer2.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer2.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer2.FormattingEnabled = True
        Me.ComboBox_layer2.Location = New System.Drawing.Point(3, 54)
        Me.ComboBox_layer2.Name = "ComboBox_layer2"
        Me.ComboBox_layer2.Size = New System.Drawing.Size(211, 23)
        Me.ComboBox_layer2.TabIndex = 0
        '
        'ComboBox_layer3
        '
        Me.ComboBox_layer3.BackColor = System.Drawing.Color.White
        Me.ComboBox_layer3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_layer3.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_layer3.FormattingEnabled = True
        Me.ComboBox_layer3.Location = New System.Drawing.Point(3, 83)
        Me.ComboBox_layer3.Name = "ComboBox_layer3"
        Me.ComboBox_layer3.Size = New System.Drawing.Size(211, 23)
        Me.ComboBox_layer3.TabIndex = 0
        '
        'Button_analise_gaps
        '
        Me.Button_analise_gaps.Location = New System.Drawing.Point(12, 186)
        Me.Button_analise_gaps.Name = "Button_analise_gaps"
        Me.Button_analise_gaps.Size = New System.Drawing.Size(251, 42)
        Me.Button_analise_gaps.TabIndex = 2
        Me.Button_analise_gaps.Text = "Detect gaps"
        Me.Button_analise_gaps.UseVisualStyleBackColor = True
        '
        'Panel_layers
        '
        Me.Panel_layers.AutoScroll = True
        Me.Panel_layers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_layers.Controls.Add(Me.ComboBox_layer1)
        Me.Panel_layers.Controls.Add(Me.ComboBox_layer2)
        Me.Panel_layers.Controls.Add(Me.ComboBox_layer3)
        Me.Panel_layers.Controls.Add(Me.Label1)
        Me.Panel_layers.Location = New System.Drawing.Point(12, 12)
        Me.Panel_layers.Name = "Panel_layers"
        Me.Panel_layers.Size = New System.Drawing.Size(251, 120)
        Me.Panel_layers.TabIndex = 3
        '
        'Button_add_new_combobox
        '
        Me.Button_add_new_combobox.Font = New System.Drawing.Font("Arial", 20.0!, System.Drawing.FontStyle.Bold)
        Me.Button_add_new_combobox.Location = New System.Drawing.Point(269, 12)
        Me.Button_add_new_combobox.Name = "Button_add_new_combobox"
        Me.Button_add_new_combobox.Size = New System.Drawing.Size(55, 45)
        Me.Button_add_new_combobox.TabIndex = 3
        Me.Button_add_new_combobox.Text = "+"
        Me.Button_add_new_combobox.UseVisualStyleBackColor = True
        '
        'Button_refresh_layers
        '
        Me.Button_refresh_layers.Location = New System.Drawing.Point(269, 86)
        Me.Button_refresh_layers.Name = "Button_refresh_layers"
        Me.Button_refresh_layers.Size = New System.Drawing.Size(96, 43)
        Me.Button_refresh_layers.TabIndex = 2
        Me.Button_refresh_layers.Text = "Refresh Layers List"
        Me.Button_refresh_layers.UseVisualStyleBackColor = True
        '
        'Button_analise_overlapp
        '
        Me.Button_analise_overlapp.Location = New System.Drawing.Point(12, 234)
        Me.Button_analise_overlapp.Name = "Button_analise_overlapp"
        Me.Button_analise_overlapp.Size = New System.Drawing.Size(251, 42)
        Me.Button_analise_overlapp.TabIndex = 2
        Me.Button_analise_overlapp.Text = "Detect overlapps"
        Me.Button_analise_overlapp.UseVisualStyleBackColor = True
        '
        'Button_dimension_To_CL
        '
        Me.Button_dimension_To_CL.Location = New System.Drawing.Point(12, 138)
        Me.Button_dimension_To_CL.Name = "Button_dimension_To_CL"
        Me.Button_dimension_To_CL.Size = New System.Drawing.Size(251, 42)
        Me.Button_dimension_To_CL.TabIndex = 2
        Me.Button_dimension_To_CL.Text = "Detect offset errors"
        Me.Button_dimension_To_CL.UseVisualStyleBackColor = True
        '
        'Button_done
        '
        Me.Button_done.Location = New System.Drawing.Point(290, 234)
        Me.Button_done.Name = "Button_done"
        Me.Button_done.Size = New System.Drawing.Size(75, 42)
        Me.Button_done.TabIndex = 4
        Me.Button_done.Text = "Done"
        Me.Button_done.UseVisualStyleBackColor = True
        '
        'Overlap_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(378, 282)
        Me.Controls.Add(Me.Button_done)
        Me.Controls.Add(Me.Button_add_new_combobox)
        Me.Controls.Add(Me.Button_refresh_layers)
        Me.Controls.Add(Me.Panel_layers)
        Me.Controls.Add(Me.Button_dimension_To_CL)
        Me.Controls.Add(Me.Button_analise_overlapp)
        Me.Controls.Add(Me.Button_analise_gaps)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Overlap_form"
        Me.Text = "TWS-ATWS"
        Me.Panel_layers.ResumeLayout(False)
        Me.Panel_layers.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ComboBox_layer1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_layer2 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_layer3 As System.Windows.Forms.ComboBox
    Friend WithEvents Button_analise_gaps As System.Windows.Forms.Button
    Friend WithEvents Panel_layers As System.Windows.Forms.Panel
    Friend WithEvents Button_analise_overlapp As System.Windows.Forms.Button
    Friend WithEvents Button_dimension_To_CL As System.Windows.Forms.Button
    Friend WithEvents Button_refresh_layers As System.Windows.Forms.Button
    Friend WithEvents Button_add_new_combobox As System.Windows.Forms.Button
    Friend WithEvents Button_done As System.Windows.Forms.Button
End Class
