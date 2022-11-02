<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Read_from_acad_w_excel_form
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button_create_mtext = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button_mtext_defl_chainage = New System.Windows.Forms.Button()
        Me.CheckBox_REROUTE = New System.Windows.Forms.CheckBox()
        Me.Button_deflections = New System.Windows.Forms.Button()
        Me.CheckBox_3d_polyline = New System.Windows.Forms.CheckBox()
        Me.CheckBox_populate_deflection = New System.Windows.Forms.CheckBox()
        Me.Button_place_chainages_on_polyline = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_LOAD_TEXT_FROM_EXCEL = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_end_load = New System.Windows.Forms.TextBox()
        Me.TextBox_col_load = New System.Windows.Forms.TextBox()
        Me.TextBox_start_load = New System.Windows.Forms.TextBox()
        Me.ListBox_deflection = New System.Windows.Forms.ListBox()
        Me.ListBox_Text_from_Autocad = New System.Windows.Forms.ListBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button_W2XL = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_column = New System.Windows.Forms.TextBox()
        Me.TextBox_ROW_START = New System.Windows.Forms.TextBox()
        Me.Button_load_text_to_combo = New System.Windows.Forms.Button()
        Me.Button_load_chainages_from_text = New System.Windows.Forms.Button()
        Me.Button_clear_defl = New System.Windows.Forms.Button()
        Me.Button_clear_text = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Button_clear_text)
        Me.Panel1.Controls.Add(Me.Button_clear_defl)
        Me.Panel1.Controls.Add(Me.Button_create_mtext)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.Button_place_chainages_on_polyline)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.ListBox_deflection)
        Me.Panel1.Controls.Add(Me.ListBox_Text_from_Autocad)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Button_load_text_to_combo)
        Me.Panel1.Controls.Add(Me.Button_load_chainages_from_text)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(534, 611)
        Me.Panel1.TabIndex = 102
        '
        'Button_create_mtext
        '
        Me.Button_create_mtext.BackColor = System.Drawing.Color.Blue
        Me.Button_create_mtext.ForeColor = System.Drawing.Color.White
        Me.Button_create_mtext.Location = New System.Drawing.Point(123, 421)
        Me.Button_create_mtext.Name = "Button_create_mtext"
        Me.Button_create_mtext.Size = New System.Drawing.Size(194, 57)
        Me.Button_create_mtext.TabIndex = 108
        Me.Button_create_mtext.Text = "Create Mtext in Autocad"
        Me.Button_create_mtext.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(323, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 14)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "LIST 2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(120, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 14)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "LIST 1"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Button_mtext_defl_chainage)
        Me.Panel4.Controls.Add(Me.CheckBox_REROUTE)
        Me.Panel4.Controls.Add(Me.Button_deflections)
        Me.Panel4.Controls.Add(Me.CheckBox_3d_polyline)
        Me.Panel4.Controls.Add(Me.CheckBox_populate_deflection)
        Me.Panel4.Location = New System.Drawing.Point(3, 484)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(514, 114)
        Me.Panel4.TabIndex = 102
        '
        'Button_mtext_defl_chainage
        '
        Me.Button_mtext_defl_chainage.BackColor = System.Drawing.Color.Lime
        Me.Button_mtext_defl_chainage.Location = New System.Drawing.Point(370, 55)
        Me.Button_mtext_defl_chainage.Name = "Button_mtext_defl_chainage"
        Me.Button_mtext_defl_chainage.Size = New System.Drawing.Size(137, 52)
        Me.Button_mtext_defl_chainage.TabIndex = 109
        Me.Button_mtext_defl_chainage.Text = "Deflection - Chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Create Alignment text"
        Me.Button_mtext_defl_chainage.UseVisualStyleBackColor = False
        '
        'CheckBox_REROUTE
        '
        Me.CheckBox_REROUTE.AutoSize = True
        Me.CheckBox_REROUTE.Location = New System.Drawing.Point(3, 88)
        Me.CheckBox_REROUTE.Name = "CheckBox_REROUTE"
        Me.CheckBox_REROUTE.Size = New System.Drawing.Size(161, 18)
        Me.CheckBox_REROUTE.TabIndex = 109
        Me.CheckBox_REROUTE.Text = "Use CHAINAGE EQUATION"
        Me.CheckBox_REROUTE.UseVisualStyleBackColor = True
        '
        'Button_deflections
        '
        Me.Button_deflections.BackColor = System.Drawing.Color.Green
        Me.Button_deflections.ForeColor = System.Drawing.Color.Yellow
        Me.Button_deflections.Location = New System.Drawing.Point(3, 3)
        Me.Button_deflections.Name = "Button_deflections"
        Me.Button_deflections.Size = New System.Drawing.Size(187, 33)
        Me.Button_deflections.TabIndex = 108
        Me.Button_deflections.Text = "Place Deflections on Polyline"
        Me.Button_deflections.UseVisualStyleBackColor = False
        '
        'CheckBox_3d_polyline
        '
        Me.CheckBox_3d_polyline.AutoSize = True
        Me.CheckBox_3d_polyline.Location = New System.Drawing.Point(3, 64)
        Me.CheckBox_3d_polyline.Name = "CheckBox_3d_polyline"
        Me.CheckBox_3d_polyline.Size = New System.Drawing.Size(300, 18)
        Me.CheckBox_3d_polyline.TabIndex = 109
        Me.CheckBox_3d_polyline.Text = "Use 3D Polyline for chainages (SLACK CHAINAGES)"
        Me.CheckBox_3d_polyline.UseVisualStyleBackColor = True
        '
        'CheckBox_populate_deflection
        '
        Me.CheckBox_populate_deflection.AutoSize = True
        Me.CheckBox_populate_deflection.Location = New System.Drawing.Point(3, 40)
        Me.CheckBox_populate_deflection.Name = "CheckBox_populate_deflection"
        Me.CheckBox_populate_deflection.Size = New System.Drawing.Size(106, 18)
        Me.CheckBox_populate_deflection.TabIndex = 109
        Me.CheckBox_populate_deflection.Text = "Add to ListBox"
        Me.CheckBox_populate_deflection.UseVisualStyleBackColor = True
        '
        'Button_place_chainages_on_polyline
        '
        Me.Button_place_chainages_on_polyline.BackColor = System.Drawing.Color.Yellow
        Me.Button_place_chainages_on_polyline.Location = New System.Drawing.Point(3, 149)
        Me.Button_place_chainages_on_polyline.Name = "Button_place_chainages_on_polyline"
        Me.Button_place_chainages_on_polyline.Size = New System.Drawing.Size(98, 62)
        Me.Button_place_chainages_on_polyline.TabIndex = 108
        Me.Button_place_chainages_on_polyline.Text = "Place Chainages on" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Polyline"
        Me.Button_place_chainages_on_polyline.UseVisualStyleBackColor = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Button_LOAD_TEXT_FROM_EXCEL)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.TextBox_end_load)
        Me.Panel3.Controls.Add(Me.TextBox_col_load)
        Me.Panel3.Controls.Add(Me.TextBox_start_load)
        Me.Panel3.Location = New System.Drawing.Point(8, 235)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(98, 130)
        Me.Panel3.TabIndex = 102
        '
        'Button_LOAD_TEXT_FROM_EXCEL
        '
        Me.Button_LOAD_TEXT_FROM_EXCEL.Location = New System.Drawing.Point(3, 76)
        Me.Button_LOAD_TEXT_FROM_EXCEL.Name = "Button_LOAD_TEXT_FROM_EXCEL"
        Me.Button_LOAD_TEXT_FROM_EXCEL.Size = New System.Drawing.Size(85, 44)
        Me.Button_LOAD_TEXT_FROM_EXCEL.TabIndex = 109
        Me.Button_LOAD_TEXT_FROM_EXCEL.Text = "Load from Excel"
        Me.Button_LOAD_TEXT_FROM_EXCEL.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 14)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "COLUMN, ROW"
        '
        'TextBox_end_load
        '
        Me.TextBox_end_load.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_end_load.Location = New System.Drawing.Point(54, 50)
        Me.TextBox_end_load.Name = "TextBox_end_load"
        Me.TextBox_end_load.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_end_load.TabIndex = 104
        '
        'TextBox_col_load
        '
        Me.TextBox_col_load.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_col_load.Location = New System.Drawing.Point(3, 24)
        Me.TextBox_col_load.Name = "TextBox_col_load"
        Me.TextBox_col_load.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_col_load.TabIndex = 104
        '
        'TextBox_start_load
        '
        Me.TextBox_start_load.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_start_load.Location = New System.Drawing.Point(54, 24)
        Me.TextBox_start_load.Name = "TextBox_start_load"
        Me.TextBox_start_load.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_start_load.TabIndex = 104
        '
        'ListBox_deflection
        '
        Me.ListBox_deflection.BackColor = System.Drawing.Color.White
        Me.ListBox_deflection.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_deflection.ForeColor = System.Drawing.Color.Black
        Me.ListBox_deflection.FormattingEnabled = True
        Me.ListBox_deflection.HorizontalScrollbar = True
        Me.ListBox_deflection.ItemHeight = 14
        Me.ListBox_deflection.Location = New System.Drawing.Point(323, 21)
        Me.ListBox_deflection.Name = "ListBox_deflection"
        Me.ListBox_deflection.Size = New System.Drawing.Size(194, 396)
        Me.ListBox_deflection.TabIndex = 5
        '
        'ListBox_Text_from_Autocad
        '
        Me.ListBox_Text_from_Autocad.BackColor = System.Drawing.Color.White
        Me.ListBox_Text_from_Autocad.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_Text_from_Autocad.ForeColor = System.Drawing.Color.Black
        Me.ListBox_Text_from_Autocad.FormattingEnabled = True
        Me.ListBox_Text_from_Autocad.HorizontalScrollbar = True
        Me.ListBox_Text_from_Autocad.ItemHeight = 14
        Me.ListBox_Text_from_Autocad.Location = New System.Drawing.Point(123, 21)
        Me.ListBox_Text_from_Autocad.Name = "ListBox_Text_from_Autocad"
        Me.ListBox_Text_from_Autocad.Size = New System.Drawing.Size(194, 396)
        Me.ListBox_Text_from_Autocad.TabIndex = 5
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Button_W2XL)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.TextBox_column)
        Me.Panel2.Controls.Add(Me.TextBox_ROW_START)
        Me.Panel2.Location = New System.Drawing.Point(8, 371)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(98, 107)
        Me.Panel2.TabIndex = 102
        '
        'Button_W2XL
        '
        Me.Button_W2XL.Location = New System.Drawing.Point(3, 50)
        Me.Button_W2XL.Name = "Button_W2XL"
        Me.Button_W2XL.Size = New System.Drawing.Size(85, 44)
        Me.Button_W2XL.TabIndex = 109
        Me.Button_W2XL.Text = "Wr. to Excel (LIST 1)"
        Me.Button_W2XL.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 14)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "COLUMN, ROW"
        '
        'TextBox_column
        '
        Me.TextBox_column.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_column.Location = New System.Drawing.Point(3, 24)
        Me.TextBox_column.Name = "TextBox_column"
        Me.TextBox_column.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_column.TabIndex = 104
        '
        'TextBox_ROW_START
        '
        Me.TextBox_ROW_START.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_ROW_START.Location = New System.Drawing.Point(54, 24)
        Me.TextBox_ROW_START.Name = "TextBox_ROW_START"
        Me.TextBox_ROW_START.Size = New System.Drawing.Size(34, 20)
        Me.TextBox_ROW_START.TabIndex = 104
        '
        'Button_load_text_to_combo
        '
        Me.Button_load_text_to_combo.Location = New System.Drawing.Point(3, 3)
        Me.Button_load_text_to_combo.Name = "Button_load_text_to_combo"
        Me.Button_load_text_to_combo.Size = New System.Drawing.Size(98, 62)
        Me.Button_load_text_to_combo.TabIndex = 108
        Me.Button_load_text_to_combo.Text = "Load ALL Text"
        Me.Button_load_text_to_combo.UseVisualStyleBackColor = True
        '
        'Button_load_chainages_from_text
        '
        Me.Button_load_chainages_from_text.BackColor = System.Drawing.Color.Yellow
        Me.Button_load_chainages_from_text.Location = New System.Drawing.Point(3, 81)
        Me.Button_load_chainages_from_text.Name = "Button_load_chainages_from_text"
        Me.Button_load_chainages_from_text.Size = New System.Drawing.Size(98, 62)
        Me.Button_load_chainages_from_text.TabIndex = 108
        Me.Button_load_chainages_from_text.Text = "Load Chainages from TEXT"
        Me.Button_load_chainages_from_text.UseVisualStyleBackColor = False
        '
        'Button_clear_defl
        '
        Me.Button_clear_defl.BackColor = System.Drawing.Color.OrangeRed
        Me.Button_clear_defl.ForeColor = System.Drawing.Color.White
        Me.Button_clear_defl.Location = New System.Drawing.Point(442, 0)
        Me.Button_clear_defl.Name = "Button_clear_defl"
        Me.Button_clear_defl.Size = New System.Drawing.Size(75, 23)
        Me.Button_clear_defl.TabIndex = 109
        Me.Button_clear_defl.Text = "Clear"
        Me.Button_clear_defl.UseVisualStyleBackColor = False
        '
        'Button_clear_text
        '
        Me.Button_clear_text.BackColor = System.Drawing.Color.OrangeRed
        Me.Button_clear_text.ForeColor = System.Drawing.Color.White
        Me.Button_clear_text.Location = New System.Drawing.Point(242, 0)
        Me.Button_clear_text.Name = "Button_clear_text"
        Me.Button_clear_text.Size = New System.Drawing.Size(75, 23)
        Me.Button_clear_text.TabIndex = 109
        Me.Button_clear_text.Text = "Clear"
        Me.Button_clear_text.UseVisualStyleBackColor = False
        '
        'Read_from_acad_w_excel_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(553, 631)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Read_from_acad_w_excel_form"
        Me.Text = "Read text from Autocad "
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ListBox_Text_from_Autocad As System.Windows.Forms.ListBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBox_ROW_START As System.Windows.Forms.TextBox
    Friend WithEvents Button_place_chainages_on_polyline As System.Windows.Forms.Button
    Friend WithEvents Button_load_chainages_from_text As System.Windows.Forms.Button
    Friend WithEvents Button_W2XL As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_column As System.Windows.Forms.TextBox
    Friend WithEvents Button_LOAD_TEXT_FROM_EXCEL As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_end_load As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_col_load As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_start_load As System.Windows.Forms.TextBox
    Friend WithEvents Button_create_mtext As System.Windows.Forms.Button
    Friend WithEvents Button_load_text_to_combo As System.Windows.Forms.Button
    Friend WithEvents Button_deflections As System.Windows.Forms.Button
    Friend WithEvents CheckBox_3d_polyline As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_populate_deflection As System.Windows.Forms.CheckBox
    Friend WithEvents ListBox_deflection As System.Windows.Forms.ListBox
    Friend WithEvents CheckBox_REROUTE As System.Windows.Forms.CheckBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button_mtext_defl_chainage As System.Windows.Forms.Button
    Friend WithEvents Button_clear_text As System.Windows.Forms.Button
    Friend WithEvents Button_clear_defl As System.Windows.Forms.Button

End Class
