<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class text_2_excel_form
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
        Me.Button_table_to_excel = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TextBox_column_width = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox_COLUMNS_FROM = New System.Windows.Forms.TextBox()
        Me.TextBox_COLUMN_TO = New System.Windows.Forms.TextBox()
        Me.TextBox_row_height = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button_Mtext_to_excel = New System.Windows.Forms.Button()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_table_to_excel
        '
        Me.Button_table_to_excel.Location = New System.Drawing.Point(12, 174)
        Me.Button_table_to_excel.Name = "Button_table_to_excel"
        Me.Button_table_to_excel.Size = New System.Drawing.Size(178, 43)
        Me.Button_table_to_excel.TabIndex = 4
        Me.Button_table_to_excel.Text = "AutoCAD to Excel"
        Me.Button_table_to_excel.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.TextBox_column_width)
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.TextBox_row_height)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Location = New System.Drawing.Point(12, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(178, 156)
        Me.Panel3.TabIndex = 5
        '
        'TextBox_column_width
        '
        Me.TextBox_column_width.BackColor = System.Drawing.Color.White
        Me.TextBox_column_width.ForeColor = System.Drawing.Color.Black
        Me.TextBox_column_width.Location = New System.Drawing.Point(107, 118)
        Me.TextBox_column_width.Name = "TextBox_column_width"
        Me.TextBox_column_width.Size = New System.Drawing.Size(52, 22)
        Me.TextBox_column_width.TabIndex = 0
        Me.TextBox_column_width.Text = "21.51"
        Me.TextBox_column_width.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TextBox_COLUMNS_FROM)
        Me.Panel2.Controls.Add(Me.TextBox_COLUMN_TO)
        Me.Panel2.Location = New System.Drawing.Point(6, 42)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(129, 35)
        Me.Panel2.TabIndex = 3
        '
        'TextBox_COLUMNS_FROM
        '
        Me.TextBox_COLUMNS_FROM.BackColor = System.Drawing.Color.White
        Me.TextBox_COLUMNS_FROM.ForeColor = System.Drawing.Color.Black
        Me.TextBox_COLUMNS_FROM.Location = New System.Drawing.Point(3, 3)
        Me.TextBox_COLUMNS_FROM.Name = "TextBox_COLUMNS_FROM"
        Me.TextBox_COLUMNS_FROM.Size = New System.Drawing.Size(37, 22)
        Me.TextBox_COLUMNS_FROM.TabIndex = 0
        Me.TextBox_COLUMNS_FROM.Text = "A"
        Me.TextBox_COLUMNS_FROM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_COLUMN_TO
        '
        Me.TextBox_COLUMN_TO.BackColor = System.Drawing.Color.White
        Me.TextBox_COLUMN_TO.ForeColor = System.Drawing.Color.Black
        Me.TextBox_COLUMN_TO.Location = New System.Drawing.Point(79, 3)
        Me.TextBox_COLUMN_TO.Name = "TextBox_COLUMN_TO"
        Me.TextBox_COLUMN_TO.Size = New System.Drawing.Size(37, 22)
        Me.TextBox_COLUMN_TO.TabIndex = 0
        Me.TextBox_COLUMN_TO.Text = "G"
        Me.TextBox_COLUMN_TO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_row_height
        '
        Me.TextBox_row_height.BackColor = System.Drawing.Color.White
        Me.TextBox_row_height.ForeColor = System.Drawing.Color.Black
        Me.TextBox_row_height.Location = New System.Drawing.Point(107, 87)
        Me.TextBox_row_height.Name = "TextBox_row_height"
        Me.TextBox_row_height.Size = New System.Drawing.Size(52, 22)
        Me.TextBox_row_height.TabIndex = 0
        Me.TextBox_row_height.Text = "6.43"
        Me.TextBox_row_height.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 32)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "From" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(81, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 32)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "To" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Location = New System.Drawing.Point(3, 121)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 18)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Column width"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label5.Location = New System.Drawing.Point(3, 90)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 18)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Row height"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button_Mtext_to_excel
        '
        Me.Button_Mtext_to_excel.Location = New System.Drawing.Point(12, 223)
        Me.Button_Mtext_to_excel.Name = "Button_Mtext_to_excel"
        Me.Button_Mtext_to_excel.Size = New System.Drawing.Size(178, 71)
        Me.Button_Mtext_to_excel.TabIndex = 4
        Me.Button_Mtext_to_excel.Text = "Mtext to Excel" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(splits rows on different columns)"
        Me.Button_Mtext_to_excel.UseVisualStyleBackColor = True
        '
        'text_2_excel_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(207, 299)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Button_Mtext_to_excel)
        Me.Controls.Add(Me.Button_table_to_excel)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "text_2_excel_form"
        Me.Text = "text as table to excel"
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button_table_to_excel As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_COLUMNS_FROM As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_COLUMN_TO As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_row_height As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox_column_width As System.Windows.Forms.TextBox
    Friend WithEvents Button_Mtext_to_excel As System.Windows.Forms.Button
End Class
