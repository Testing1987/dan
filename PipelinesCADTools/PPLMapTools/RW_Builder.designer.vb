<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RW_Builder
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
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Panel16 = New System.Windows.Forms.Panel()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.TextBox_col_type = New System.Windows.Forms.TextBox()
        Me.TextBox_col_offset = New System.Windows.Forms.TextBox()
        Me.TextBox_col_end = New System.Windows.Forms.TextBox()
        Me.TextBox_col_start = New System.Windows.Forms.TextBox()
        Me.TextBox_col_width = New System.Windows.Forms.TextBox()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.TextBox_ROW_END = New System.Windows.Forms.TextBox()
        Me.TextBox_ROW_START = New System.Windows.Forms.TextBox()
        Me.Button_draw = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.Panel16.SuspendLayout()
        Me.Panel14.SuspendLayout()
        Me.Panel12.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 93)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(444, 453)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage1.Location = New System.Drawing.Point(4, 24)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(436, 425)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "R/W"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage2.Location = New System.Drawing.Point(4, 24)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(436, 425)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TabPage2"
        '
        'Panel16
        '
        Me.Panel16.BackColor = System.Drawing.Color.LightGray
        Me.Panel16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel16.Controls.Add(Me.Label60)
        Me.Panel16.Controls.Add(Me.Label59)
        Me.Panel16.Controls.Add(Me.Label52)
        Me.Panel16.Controls.Add(Me.Label61)
        Me.Panel16.Controls.Add(Me.Label51)
        Me.Panel16.Location = New System.Drawing.Point(12, 24)
        Me.Panel16.Name = "Panel16"
        Me.Panel16.Size = New System.Drawing.Size(319, 28)
        Me.Panel16.TabIndex = 4
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label60.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label60.Location = New System.Drawing.Point(129, 5)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(35, 16)
        Me.Label60.TabIndex = 109
        Me.Label60.Text = "Type"
        Me.Label60.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label59.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label59.Location = New System.Drawing.Point(15, 5)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(35, 16)
        Me.Label59.TabIndex = 109
        Me.Label59.Text = "Start"
        Me.Label59.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label52.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label52.Location = New System.Drawing.Point(72, 5)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(29, 16)
        Me.Label52.TabIndex = 109
        Me.Label52.Text = "End"
        Me.Label52.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label61
        '
        Me.Label61.AutoSize = True
        Me.Label61.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label61.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label61.Location = New System.Drawing.Point(253, 5)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(43, 16)
        Me.Label61.TabIndex = 108
        Me.Label61.Text = "Offset"
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label51.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label51.Location = New System.Drawing.Point(186, 5)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(40, 16)
        Me.Label51.TabIndex = 108
        Me.Label51.Text = "Width"
        '
        'Panel14
        '
        Me.Panel14.BackColor = System.Drawing.Color.LightGray
        Me.Panel14.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel14.Controls.Add(Me.TextBox_col_type)
        Me.Panel14.Controls.Add(Me.TextBox_col_offset)
        Me.Panel14.Controls.Add(Me.TextBox_col_end)
        Me.Panel14.Controls.Add(Me.TextBox_col_start)
        Me.Panel14.Controls.Add(Me.TextBox_col_width)
        Me.Panel14.Location = New System.Drawing.Point(12, 55)
        Me.Panel14.Name = "Panel14"
        Me.Panel14.Size = New System.Drawing.Size(319, 32)
        Me.Panel14.TabIndex = 5
        '
        'TextBox_col_type
        '
        Me.TextBox_col_type.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_col_type.Location = New System.Drawing.Point(129, 3)
        Me.TextBox_col_type.Name = "TextBox_col_type"
        Me.TextBox_col_type.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_col_type.TabIndex = 111
        Me.TextBox_col_type.Text = "C"
        Me.TextBox_col_type.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_col_offset
        '
        Me.TextBox_col_offset.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_col_offset.Location = New System.Drawing.Point(253, 3)
        Me.TextBox_col_offset.Name = "TextBox_col_offset"
        Me.TextBox_col_offset.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_col_offset.TabIndex = 111
        Me.TextBox_col_offset.Text = "E"
        Me.TextBox_col_offset.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_col_end
        '
        Me.TextBox_col_end.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_col_end.Location = New System.Drawing.Point(72, 3)
        Me.TextBox_col_end.Name = "TextBox_col_end"
        Me.TextBox_col_end.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_col_end.TabIndex = 111
        Me.TextBox_col_end.Text = "B"
        Me.TextBox_col_end.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_col_start
        '
        Me.TextBox_col_start.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_col_start.Location = New System.Drawing.Point(15, 3)
        Me.TextBox_col_start.Name = "TextBox_col_start"
        Me.TextBox_col_start.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_col_start.TabIndex = 111
        Me.TextBox_col_start.Text = "A"
        Me.TextBox_col_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_col_width
        '
        Me.TextBox_col_width.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_col_width.Location = New System.Drawing.Point(186, 3)
        Me.TextBox_col_width.Name = "TextBox_col_width"
        Me.TextBox_col_width.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_col_width.TabIndex = 110
        Me.TextBox_col_width.Text = "D"
        Me.TextBox_col_width.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel12
        '
        Me.Panel12.BackColor = System.Drawing.Color.LightGray
        Me.Panel12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel12.Controls.Add(Me.Label53)
        Me.Panel12.Controls.Add(Me.Label54)
        Me.Panel12.Controls.Add(Me.TextBox_ROW_END)
        Me.Panel12.Controls.Add(Me.TextBox_ROW_START)
        Me.Panel12.Location = New System.Drawing.Point(12, 552)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(132, 58)
        Me.Panel12.TabIndex = 114
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label53.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label53.Location = New System.Drawing.Point(43, 4)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(78, 15)
        Me.Label53.TabIndex = 6
        Me.Label53.Text = "FROM ROW"
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label54.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label54.Location = New System.Drawing.Point(43, 35)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(76, 15)
        Me.Label54.TabIndex = 6
        Me.Label54.Text = "TO ROW    "
        '
        'TextBox_ROW_END
        '
        Me.TextBox_ROW_END.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_ROW_END.Location = New System.Drawing.Point(3, 32)
        Me.TextBox_ROW_END.Name = "TextBox_ROW_END"
        Me.TextBox_ROW_END.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_ROW_END.TabIndex = 107
        Me.TextBox_ROW_END.Text = "4"
        '
        'TextBox_ROW_START
        '
        Me.TextBox_ROW_START.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_ROW_START.Location = New System.Drawing.Point(3, 2)
        Me.TextBox_ROW_START.Name = "TextBox_ROW_START"
        Me.TextBox_ROW_START.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_ROW_START.TabIndex = 104
        Me.TextBox_ROW_START.Text = "2"
        '
        'Button_draw
        '
        Me.Button_draw.Location = New System.Drawing.Point(315, 573)
        Me.Button_draw.Name = "Button_draw"
        Me.Button_draw.Size = New System.Drawing.Size(137, 37)
        Me.Button_draw.TabIndex = 115
        Me.Button_draw.Text = "Draw"
        Me.Button_draw.UseVisualStyleBackColor = True
        '
        'RW_Builder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(464, 620)
        Me.Controls.Add(Me.Button_draw)
        Me.Controls.Add(Me.Panel12)
        Me.Controls.Add(Me.Panel16)
        Me.Controls.Add(Me.Panel14)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "RW_Builder"
        Me.Text = "RW Builder"
        Me.TabControl1.ResumeLayout(False)
        Me.Panel16.ResumeLayout(False)
        Me.Panel16.PerformLayout()
        Me.Panel14.ResumeLayout(False)
        Me.Panel14.PerformLayout()
        Me.Panel12.ResumeLayout(False)
        Me.Panel12.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Panel16 As System.Windows.Forms.Panel
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents Panel14 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_col_type As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_col_offset As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_col_end As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_col_start As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_col_width As System.Windows.Forms.TextBox
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents TextBox_ROW_END As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_ROW_START As System.Windows.Forms.TextBox
    Friend WithEvents Button_draw As System.Windows.Forms.Button
End Class
