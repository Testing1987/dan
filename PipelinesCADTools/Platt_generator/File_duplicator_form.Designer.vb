<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class File_duplicator_form
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
        Me.Button_sheet_set_template = New System.Windows.Forms.Button()
        Me.TextBox_sheet_set_template = New System.Windows.Forms.TextBox()
        Me.TextBox_NAME = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_INCREMENT_START = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_INCREMENT_END = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button_duplicate_ = New System.Windows.Forms.Button()
        Me.TextBox_Storage_folder = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_publish_folder = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_page_setup_overrides = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button_sheet_set_template
        '
        Me.Button_sheet_set_template.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_sheet_set_template.Location = New System.Drawing.Point(13, 13)
        Me.Button_sheet_set_template.Margin = New System.Windows.Forms.Padding(4)
        Me.Button_sheet_set_template.Name = "Button_sheet_set_template"
        Me.Button_sheet_set_template.Size = New System.Drawing.Size(127, 40)
        Me.Button_sheet_set_template.TabIndex = 3
        Me.Button_sheet_set_template.Text = "SELECT FILE"
        Me.Button_sheet_set_template.UseVisualStyleBackColor = True
        '
        'TextBox_sheet_set_template
        '
        Me.TextBox_sheet_set_template.BackColor = System.Drawing.Color.White
        Me.TextBox_sheet_set_template.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_sheet_set_template.ForeColor = System.Drawing.Color.Black
        Me.TextBox_sheet_set_template.Location = New System.Drawing.Point(155, 13)
        Me.TextBox_sheet_set_template.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox_sheet_set_template.Multiline = True
        Me.TextBox_sheet_set_template.Name = "TextBox_sheet_set_template"
        Me.TextBox_sheet_set_template.Size = New System.Drawing.Size(544, 40)
        Me.TextBox_sheet_set_template.TabIndex = 4
        '
        'TextBox_NAME
        '
        Me.TextBox_NAME.BackColor = System.Drawing.Color.White
        Me.TextBox_NAME.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_NAME.ForeColor = System.Drawing.Color.Black
        Me.TextBox_NAME.Location = New System.Drawing.Point(166, 214)
        Me.TextBox_NAME.Name = "TextBox_NAME"
        Me.TextBox_NAME.Size = New System.Drawing.Size(206, 26)
        Me.TextBox_NAME.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(64, 219)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 18)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "FILE NAME"
        '
        'TextBox_INCREMENT_START
        '
        Me.TextBox_INCREMENT_START.BackColor = System.Drawing.Color.White
        Me.TextBox_INCREMENT_START.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_INCREMENT_START.ForeColor = System.Drawing.Color.Black
        Me.TextBox_INCREMENT_START.Location = New System.Drawing.Point(301, 246)
        Me.TextBox_INCREMENT_START.Name = "TextBox_INCREMENT_START"
        Me.TextBox_INCREMENT_START.Size = New System.Drawing.Size(71, 26)
        Me.TextBox_INCREMENT_START.TabIndex = 5
        Me.TextBox_INCREMENT_START.Text = "01"
        Me.TextBox_INCREMENT_START.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(168, 251)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 18)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "START NUMBER"
        '
        'TextBox_INCREMENT_END
        '
        Me.TextBox_INCREMENT_END.BackColor = System.Drawing.Color.White
        Me.TextBox_INCREMENT_END.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_INCREMENT_END.ForeColor = System.Drawing.Color.Black
        Me.TextBox_INCREMENT_END.Location = New System.Drawing.Point(301, 278)
        Me.TextBox_INCREMENT_END.Name = "TextBox_INCREMENT_END"
        Me.TextBox_INCREMENT_END.Size = New System.Drawing.Size(71, 26)
        Me.TextBox_INCREMENT_END.TabIndex = 5
        Me.TextBox_INCREMENT_END.Text = "10"
        Me.TextBox_INCREMENT_END.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label4.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(182, 283)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "END NUMBER"
        '
        'Button_duplicate_
        '
        Me.Button_duplicate_.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_duplicate_.Location = New System.Drawing.Point(572, 281)
        Me.Button_duplicate_.Margin = New System.Windows.Forms.Padding(4)
        Me.Button_duplicate_.Name = "Button_duplicate_"
        Me.Button_duplicate_.Size = New System.Drawing.Size(127, 40)
        Me.Button_duplicate_.TabIndex = 3
        Me.Button_duplicate_.Text = "DUPLICATE FILE"
        Me.Button_duplicate_.UseVisualStyleBackColor = True
        '
        'TextBox_Storage_folder
        '
        Me.TextBox_Storage_folder.BackColor = System.Drawing.Color.White
        Me.TextBox_Storage_folder.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Storage_folder.ForeColor = System.Drawing.Color.Black
        Me.TextBox_Storage_folder.Location = New System.Drawing.Point(155, 61)
        Me.TextBox_Storage_folder.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox_Storage_folder.Multiline = True
        Me.TextBox_Storage_folder.Name = "TextBox_Storage_folder"
        Me.TextBox_Storage_folder.Size = New System.Drawing.Size(544, 40)
        Me.TextBox_Storage_folder.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(16, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 18)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Storage FOLDER"
        '
        'TextBox_publish_folder
        '
        Me.TextBox_publish_folder.BackColor = System.Drawing.Color.White
        Me.TextBox_publish_folder.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_publish_folder.ForeColor = System.Drawing.Color.Black
        Me.TextBox_publish_folder.Location = New System.Drawing.Point(155, 109)
        Me.TextBox_publish_folder.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox_publish_folder.Multiline = True
        Me.TextBox_publish_folder.Name = "TextBox_publish_folder"
        Me.TextBox_publish_folder.Size = New System.Drawing.Size(544, 40)
        Me.TextBox_publish_folder.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label5.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(16, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(123, 18)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Publish FOLDER"
        '
        'TextBox_page_setup_overrides
        '
        Me.TextBox_page_setup_overrides.BackColor = System.Drawing.Color.White
        Me.TextBox_page_setup_overrides.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_page_setup_overrides.ForeColor = System.Drawing.Color.Black
        Me.TextBox_page_setup_overrides.Location = New System.Drawing.Point(155, 157)
        Me.TextBox_page_setup_overrides.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox_page_setup_overrides.Multiline = True
        Me.TextBox_page_setup_overrides.Name = "TextBox_page_setup_overrides"
        Me.TextBox_page_setup_overrides.Size = New System.Drawing.Size(544, 40)
        Me.TextBox_page_setup_overrides.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(16, 162)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(111, 34)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Page Setup " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Overrides File"
        '
        'File_duplicator_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(712, 334)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_INCREMENT_END)
        Me.Controls.Add(Me.TextBox_INCREMENT_START)
        Me.Controls.Add(Me.TextBox_NAME)
        Me.Controls.Add(Me.TextBox_page_setup_overrides)
        Me.Controls.Add(Me.TextBox_publish_folder)
        Me.Controls.Add(Me.TextBox_Storage_folder)
        Me.Controls.Add(Me.TextBox_sheet_set_template)
        Me.Controls.Add(Me.Button_duplicate_)
        Me.Controls.Add(Me.Button_sheet_set_template)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.Name = "File_duplicator_form"
        Me.Text = "DST duplicator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_sheet_set_template As System.Windows.Forms.Button
    Friend WithEvents TextBox_sheet_set_template As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_NAME As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_INCREMENT_START As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_INCREMENT_END As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button_duplicate_ As System.Windows.Forms.Button
    Friend WithEvents TextBox_Storage_folder As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_publish_folder As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_page_setup_overrides As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
