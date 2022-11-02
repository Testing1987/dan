<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FIND_REPLACE_Form
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
        Me.TextBox_source_folder = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_destination_folder = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button_current_dwg = New System.Windows.Forms.Button()
        Me.Button_Load_Info_from_excel = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_REPLACE_COL_XL = New System.Windows.Forms.TextBox()
        Me.ListBox_REPLACE = New System.Windows.Forms.ListBox()
        Me.TextBox_FIND_col_xl = New System.Windows.Forms.TextBox()
        Me.ListBox_FIND = New System.Windows.Forms.ListBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox_ROW_END = New System.Windows.Forms.TextBox()
        Me.TextBox_ROW_START = New System.Windows.Forms.TextBox()
        Me.CheckBox_REPLACE_ONLY_WORD = New System.Windows.Forms.CheckBox()
        Me.CheckBox_search_modelspace = New System.Windows.Forms.CheckBox()
        Me.Button_EXECUTE = New System.Windows.Forms.Button()
        Me.Button_rename_files = New System.Windows.Forms.Button()
        Me.ComboBox_color = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Button_FIND_REPLACE_SELECTION = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_source_folder
        '
        Me.TextBox_source_folder.BackColor = System.Drawing.Color.White
        Me.TextBox_source_folder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_source_folder.ForeColor = System.Drawing.Color.Black
        Me.TextBox_source_folder.Location = New System.Drawing.Point(845, 22)
        Me.TextBox_source_folder.Multiline = True
        Me.TextBox_source_folder.Name = "TextBox_source_folder"
        Me.TextBox_source_folder.Size = New System.Drawing.Size(473, 45)
        Me.TextBox_source_folder.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(842, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 14)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Source Folder"
        '
        'TextBox_destination_folder
        '
        Me.TextBox_destination_folder.BackColor = System.Drawing.Color.White
        Me.TextBox_destination_folder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_destination_folder.ForeColor = System.Drawing.Color.Black
        Me.TextBox_destination_folder.Location = New System.Drawing.Point(845, 93)
        Me.TextBox_destination_folder.Multiline = True
        Me.TextBox_destination_folder.Name = "TextBox_destination_folder"
        Me.TextBox_destination_folder.Size = New System.Drawing.Size(476, 45)
        Me.TextBox_destination_folder.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(842, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 14)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Destination Folder"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.ComboBox_color)
        Me.Panel1.Controls.Add(Me.Button_FIND_REPLACE_SELECTION)
        Me.Panel1.Controls.Add(Me.Button_current_dwg)
        Me.Panel1.Controls.Add(Me.Button_Load_Info_from_excel)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.TextBox_REPLACE_COL_XL)
        Me.Panel1.Controls.Add(Me.ListBox_REPLACE)
        Me.Panel1.Controls.Add(Me.TextBox_FIND_col_xl)
        Me.Panel1.Controls.Add(Me.ListBox_FIND)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(561, 441)
        Me.Panel1.TabIndex = 102
        '
        'Button_current_dwg
        '
        Me.Button_current_dwg.Location = New System.Drawing.Point(3, 214)
        Me.Button_current_dwg.Name = "Button_current_dwg"
        Me.Button_current_dwg.Size = New System.Drawing.Size(129, 42)
        Me.Button_current_dwg.TabIndex = 103
        Me.Button_current_dwg.Text = "Find - Replace" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ALL DWG"
        Me.Button_current_dwg.UseVisualStyleBackColor = True
        '
        'Button_Load_Info_from_excel
        '
        Me.Button_Load_Info_from_excel.Location = New System.Drawing.Point(8, 140)
        Me.Button_Load_Info_from_excel.Name = "Button_Load_Info_from_excel"
        Me.Button_Load_Info_from_excel.Size = New System.Drawing.Size(100, 49)
        Me.Button_Load_Info_from_excel.TabIndex = 103
        Me.Button_Load_Info_from_excel.Text = "Load from Excel"
        Me.Button_Load_Info_from_excel.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(141, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(114, 14)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "FIND FROM COLUMN"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(351, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 14)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "REPLACE WITH"
        '
        'TextBox_REPLACE_COL_XL
        '
        Me.TextBox_REPLACE_COL_XL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_REPLACE_COL_XL.Location = New System.Drawing.Point(515, 4)
        Me.TextBox_REPLACE_COL_XL.Name = "TextBox_REPLACE_COL_XL"
        Me.TextBox_REPLACE_COL_XL.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_REPLACE_COL_XL.TabIndex = 107
        Me.TextBox_REPLACE_COL_XL.Text = "B"
        Me.TextBox_REPLACE_COL_XL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ListBox_REPLACE
        '
        Me.ListBox_REPLACE.BackColor = System.Drawing.Color.White
        Me.ListBox_REPLACE.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_REPLACE.ForeColor = System.Drawing.Color.Black
        Me.ListBox_REPLACE.FormattingEnabled = True
        Me.ListBox_REPLACE.HorizontalScrollbar = True
        Me.ListBox_REPLACE.ItemHeight = 14
        Me.ListBox_REPLACE.Location = New System.Drawing.Point(354, 29)
        Me.ListBox_REPLACE.Name = "ListBox_REPLACE"
        Me.ListBox_REPLACE.Size = New System.Drawing.Size(194, 396)
        Me.ListBox_REPLACE.TabIndex = 4
        '
        'TextBox_FIND_col_xl
        '
        Me.TextBox_FIND_col_xl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_FIND_col_xl.Location = New System.Drawing.Point(304, 4)
        Me.TextBox_FIND_col_xl.Name = "TextBox_FIND_col_xl"
        Me.TextBox_FIND_col_xl.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_FIND_col_xl.TabIndex = 104
        Me.TextBox_FIND_col_xl.Text = "A"
        Me.TextBox_FIND_col_xl.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ListBox_FIND
        '
        Me.ListBox_FIND.BackColor = System.Drawing.Color.White
        Me.ListBox_FIND.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_FIND.ForeColor = System.Drawing.Color.Black
        Me.ListBox_FIND.FormattingEnabled = True
        Me.ListBox_FIND.HorizontalScrollbar = True
        Me.ListBox_FIND.ItemHeight = 14
        Me.ListBox_FIND.Location = New System.Drawing.Point(144, 29)
        Me.ListBox_FIND.Name = "ListBox_FIND"
        Me.ListBox_FIND.Size = New System.Drawing.Size(194, 396)
        Me.ListBox_FIND.TabIndex = 5
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.TextBox_ROW_END)
        Me.Panel2.Controls.Add(Me.TextBox_ROW_START)
        Me.Panel2.Location = New System.Drawing.Point(10, 13)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(98, 121)
        Me.Panel2.TabIndex = 102
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 14)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "FROM ROW"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 54)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 14)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "TO ROW"
        '
        'TextBox_ROW_END
        '
        Me.TextBox_ROW_END.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_ROW_END.Location = New System.Drawing.Point(15, 71)
        Me.TextBox_ROW_END.Name = "TextBox_ROW_END"
        Me.TextBox_ROW_END.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_ROW_END.TabIndex = 107
        '
        'TextBox_ROW_START
        '
        Me.TextBox_ROW_START.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_ROW_START.Location = New System.Drawing.Point(15, 28)
        Me.TextBox_ROW_START.Name = "TextBox_ROW_START"
        Me.TextBox_ROW_START.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_ROW_START.TabIndex = 104
        '
        'CheckBox_REPLACE_ONLY_WORD
        '
        Me.CheckBox_REPLACE_ONLY_WORD.AutoSize = True
        Me.CheckBox_REPLACE_ONLY_WORD.Location = New System.Drawing.Point(845, 151)
        Me.CheckBox_REPLACE_ONLY_WORD.Name = "CheckBox_REPLACE_ONLY_WORD"
        Me.CheckBox_REPLACE_ONLY_WORD.Size = New System.Drawing.Size(114, 32)
        Me.CheckBox_REPLACE_ONLY_WORD.TabIndex = 104
        Me.CheckBox_REPLACE_ONLY_WORD.Text = "Replace" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Only if is a word"
        Me.CheckBox_REPLACE_ONLY_WORD.UseVisualStyleBackColor = True
        '
        'CheckBox_search_modelspace
        '
        Me.CheckBox_search_modelspace.AutoSize = True
        Me.CheckBox_search_modelspace.Location = New System.Drawing.Point(845, 189)
        Me.CheckBox_search_modelspace.Name = "CheckBox_search_modelspace"
        Me.CheckBox_search_modelspace.Size = New System.Drawing.Size(100, 32)
        Me.CheckBox_search_modelspace.TabIndex = 104
        Me.CheckBox_search_modelspace.Text = "Search" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "MODEL Space"
        Me.CheckBox_search_modelspace.UseVisualStyleBackColor = True
        '
        'Button_EXECUTE
        '
        Me.Button_EXECUTE.Location = New System.Drawing.Point(997, 145)
        Me.Button_EXECUTE.Name = "Button_EXECUTE"
        Me.Button_EXECUTE.Size = New System.Drawing.Size(129, 42)
        Me.Button_EXECUTE.TabIndex = 103
        Me.Button_EXECUTE.Text = "Find - Replace" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Background)"
        Me.Button_EXECUTE.UseVisualStyleBackColor = True
        '
        'Button_rename_files
        '
        Me.Button_rename_files.Location = New System.Drawing.Point(1153, 150)
        Me.Button_rename_files.Name = "Button_rename_files"
        Me.Button_rename_files.Size = New System.Drawing.Size(129, 53)
        Me.Button_rename_files.TabIndex = 103
        Me.Button_rename_files.Text = "Find - Replace" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "RENAME FILES" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Destination Folder)"
        Me.Button_rename_files.UseVisualStyleBackColor = True
        '
        'ComboBox_color
        '
        Me.ComboBox_color.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_color.FormattingEnabled = True
        Me.ComboBox_color.Location = New System.Drawing.Point(8, 403)
        Me.ComboBox_color.Name = "ComboBox_color"
        Me.ComboBox_color.Size = New System.Drawing.Size(121, 22)
        Me.ComboBox_color.TabIndex = 108
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 386)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(37, 14)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Color"
        '
        'Button_FIND_REPLACE_SELECTION
        '
        Me.Button_FIND_REPLACE_SELECTION.Location = New System.Drawing.Point(3, 286)
        Me.Button_FIND_REPLACE_SELECTION.Name = "Button_FIND_REPLACE_SELECTION"
        Me.Button_FIND_REPLACE_SELECTION.Size = New System.Drawing.Size(129, 42)
        Me.Button_FIND_REPLACE_SELECTION.TabIndex = 103
        Me.Button_FIND_REPLACE_SELECTION.Text = "Find - Replace" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SELECTION"
        Me.Button_FIND_REPLACE_SELECTION.UseVisualStyleBackColor = True
        '
        'FIND_REPLACE_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(591, 461)
        Me.Controls.Add(Me.CheckBox_search_modelspace)
        Me.Controls.Add(Me.CheckBox_REPLACE_ONLY_WORD)
        Me.Controls.Add(Me.Button_rename_files)
        Me.Controls.Add(Me.Button_EXECUTE)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_destination_folder)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_source_folder)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "FIND_REPLACE_Form"
        Me.Text = "FIND - REPLACE WITH DATA FROM EXCEL"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_source_folder As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_destination_folder As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button_EXECUTE As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ListBox_REPLACE As System.Windows.Forms.ListBox
    Friend WithEvents ListBox_FIND As System.Windows.Forms.ListBox
    Friend WithEvents Button_Load_Info_from_excel As System.Windows.Forms.Button
    Friend WithEvents TextBox_REPLACE_COL_XL As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_FIND_col_xl As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox_ROW_END As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_ROW_START As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox_search_modelspace As System.Windows.Forms.CheckBox
    Friend WithEvents Button_current_dwg As System.Windows.Forms.Button
    Friend WithEvents CheckBox_REPLACE_ONLY_WORD As System.Windows.Forms.CheckBox
    Friend WithEvents Button_rename_files As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_color As System.Windows.Forms.ComboBox
    Friend WithEvents Button_FIND_REPLACE_SELECTION As System.Windows.Forms.Button

End Class
