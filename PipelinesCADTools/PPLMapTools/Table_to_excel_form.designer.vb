<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Table_to_excel_form
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
        Me.TextBox_COLUMNS_FROM = New System.Windows.Forms.TextBox()
        Me.TextBox_COLUMN_TO = New System.Windows.Forms.TextBox()
        Me.TextBox_Row_start = New System.Windows.Forms.TextBox()
        Me.TextBox_Row_End = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button_Excel_TO_Existing_Table = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_table_to_excel = New System.Windows.Forms.Button()
        Me.Button_create_new_table = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_COLUMNS_FROM
        '
        Me.TextBox_COLUMNS_FROM.BackColor = System.Drawing.Color.White
        Me.TextBox_COLUMNS_FROM.ForeColor = System.Drawing.Color.Black
        Me.TextBox_COLUMNS_FROM.Location = New System.Drawing.Point(3, 3)
        Me.TextBox_COLUMNS_FROM.Name = "TextBox_COLUMNS_FROM"
        Me.TextBox_COLUMNS_FROM.Size = New System.Drawing.Size(37, 21)
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
        Me.TextBox_COLUMN_TO.Size = New System.Drawing.Size(37, 21)
        Me.TextBox_COLUMN_TO.TabIndex = 0
        Me.TextBox_COLUMN_TO.Text = "G"
        Me.TextBox_COLUMN_TO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_Row_start
        '
        Me.TextBox_Row_start.BackColor = System.Drawing.Color.White
        Me.TextBox_Row_start.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Row_start.Location = New System.Drawing.Point(82, 3)
        Me.TextBox_Row_start.Name = "TextBox_Row_start"
        Me.TextBox_Row_start.Size = New System.Drawing.Size(37, 21)
        Me.TextBox_Row_start.TabIndex = 0
        Me.TextBox_Row_start.Text = "1"
        Me.TextBox_Row_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_Row_End
        '
        Me.TextBox_Row_End.BackColor = System.Drawing.Color.White
        Me.TextBox_Row_End.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Row_End.Location = New System.Drawing.Point(82, 30)
        Me.TextBox_Row_End.Name = "TextBox_Row_End"
        Me.TextBox_Row_End.Size = New System.Drawing.Size(37, 21)
        Me.TextBox_Row_End.TabIndex = 0
        Me.TextBox_Row_End.Text = "60"
        Me.TextBox_Row_End.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 30)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "From" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(81, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 30)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "To" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 15)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Row Start"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 33)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 15)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Row End"
        '
        'Button_Excel_TO_Existing_Table
        '
        Me.Button_Excel_TO_Existing_Table.Location = New System.Drawing.Point(165, 61)
        Me.Button_Excel_TO_Existing_Table.Name = "Button_Excel_TO_Existing_Table"
        Me.Button_Excel_TO_Existing_Table.Size = New System.Drawing.Size(154, 43)
        Me.Button_Excel_TO_Existing_Table.TabIndex = 2
        Me.Button_Excel_TO_Existing_Table.Text = "Excel to Existing Table"
        Me.Button_Excel_TO_Existing_Table.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.TextBox_Row_start)
        Me.Panel1.Controls.Add(Me.TextBox_Row_End)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(6, 83)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(129, 62)
        Me.Panel1.TabIndex = 3
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
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Location = New System.Drawing.Point(12, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(147, 159)
        Me.Panel3.TabIndex = 3
        '
        'Button_table_to_excel
        '
        Me.Button_table_to_excel.Location = New System.Drawing.Point(165, 12)
        Me.Button_table_to_excel.Name = "Button_table_to_excel"
        Me.Button_table_to_excel.Size = New System.Drawing.Size(154, 43)
        Me.Button_table_to_excel.TabIndex = 2
        Me.Button_table_to_excel.Text = "AutoCAD to Excel"
        Me.Button_table_to_excel.UseVisualStyleBackColor = True
        '
        'Button_create_new_table
        '
        Me.Button_create_new_table.Location = New System.Drawing.Point(165, 110)
        Me.Button_create_new_table.Name = "Button_create_new_table"
        Me.Button_create_new_table.Size = New System.Drawing.Size(154, 43)
        Me.Button_create_new_table.TabIndex = 2
        Me.Button_create_new_table.Text = "Excel to a New Table"
        Me.Button_create_new_table.UseVisualStyleBackColor = True
        '
        'Table_to_excel_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(325, 181)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Button_table_to_excel)
        Me.Controls.Add(Me.Button_create_new_table)
        Me.Controls.Add(Me.Button_Excel_TO_Existing_Table)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Table_to_excel_form"
        Me.Text = "Table to excel"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TextBox_COLUMNS_FROM As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_COLUMN_TO As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Row_start As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Row_End As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button_Excel_TO_Existing_Table As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button_table_to_excel As System.Windows.Forms.Button
    Friend WithEvents Button_create_new_table As System.Windows.Forms.Button
End Class
