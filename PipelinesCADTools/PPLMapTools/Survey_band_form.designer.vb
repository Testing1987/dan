<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Survey_band_form
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
        Me.Button_TRANSFER_TO_AUTOCAD = New System.Windows.Forms.Button()
        Me.Button_read_from_Excel = New System.Windows.Forms.Button()
        Me.TextBox_row_start = New System.Windows.Forms.MaskedTextBox()
        Me.TextBox_row_end = New System.Windows.Forms.MaskedTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.RadioButton_right_left = New System.Windows.Forms.RadioButton()
        Me.RadioButton_left_right = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_TRANSFER_TO_AUTOCAD
        '
        Me.Button_TRANSFER_TO_AUTOCAD.BackColor = System.Drawing.Color.LimeGreen
        Me.Button_TRANSFER_TO_AUTOCAD.ForeColor = System.Drawing.Color.Black
        Me.Button_TRANSFER_TO_AUTOCAD.Location = New System.Drawing.Point(162, 147)
        Me.Button_TRANSFER_TO_AUTOCAD.Name = "Button_TRANSFER_TO_AUTOCAD"
        Me.Button_TRANSFER_TO_AUTOCAD.Size = New System.Drawing.Size(138, 47)
        Me.Button_TRANSFER_TO_AUTOCAD.TabIndex = 109
        Me.Button_TRANSFER_TO_AUTOCAD.Text = "Transfer info to Autocad"
        Me.Button_TRANSFER_TO_AUTOCAD.UseVisualStyleBackColor = False
        '
        'Button_read_from_Excel
        '
        Me.Button_read_from_Excel.BackColor = System.Drawing.Color.LimeGreen
        Me.Button_read_from_Excel.ForeColor = System.Drawing.Color.Black
        Me.Button_read_from_Excel.Location = New System.Drawing.Point(6, 77)
        Me.Button_read_from_Excel.Name = "Button_read_from_Excel"
        Me.Button_read_from_Excel.Size = New System.Drawing.Size(125, 32)
        Me.Button_read_from_Excel.TabIndex = 114
        Me.Button_read_from_Excel.Text = "Load from Excel"
        Me.Button_read_from_Excel.UseVisualStyleBackColor = False
        '
        'TextBox_row_start
        '
        Me.TextBox_row_start.BackColor = System.Drawing.Color.White
        Me.TextBox_row_start.ForeColor = System.Drawing.Color.Black
        Me.TextBox_row_start.Location = New System.Drawing.Point(6, 23)
        Me.TextBox_row_start.Name = "TextBox_row_start"
        Me.TextBox_row_start.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_row_start.TabIndex = 0
        '
        'TextBox_row_end
        '
        Me.TextBox_row_end.BackColor = System.Drawing.Color.White
        Me.TextBox_row_end.ForeColor = System.Drawing.Color.Black
        Me.TextBox_row_end.Location = New System.Drawing.Point(6, 50)
        Me.TextBox_row_end.Name = "TextBox_row_end"
        Me.TextBox_row_end.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_row_end.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(57, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 15)
        Me.Label4.TabIndex = 115
        Me.Label4.Text = "Row Start"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(57, 53)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 15)
        Me.Label5.TabIndex = 115
        Me.Label5.Text = "Row End"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 2)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 15)
        Me.Label6.TabIndex = 115
        Me.Label6.Text = "Column A B C D E"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.TextBox_row_start)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.TextBox_row_end)
        Me.Panel2.Controls.Add(Me.Button_read_from_Excel)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Location = New System.Drawing.Point(12, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(138, 123)
        Me.Panel2.TabIndex = 116
        '
        'RadioButton_right_left
        '
        Me.RadioButton_right_left.AutoSize = True
        Me.RadioButton_right_left.Checked = True
        Me.RadioButton_right_left.Location = New System.Drawing.Point(3, 3)
        Me.RadioButton_right_left.Name = "RadioButton_right_left"
        Me.RadioButton_right_left.Size = New System.Drawing.Size(93, 19)
        Me.RadioButton_right_left.TabIndex = 117
        Me.RadioButton_right_left.TabStop = True
        Me.RadioButton_right_left.Text = "Right to Left"
        Me.RadioButton_right_left.UseVisualStyleBackColor = True
        '
        'RadioButton_left_right
        '
        Me.RadioButton_left_right.AutoSize = True
        Me.RadioButton_left_right.Location = New System.Drawing.Point(3, 28)
        Me.RadioButton_left_right.Name = "RadioButton_left_right"
        Me.RadioButton_left_right.Size = New System.Drawing.Size(93, 19)
        Me.RadioButton_left_right.TabIndex = 117
        Me.RadioButton_left_right.Text = "Left to Right"
        Me.RadioButton_left_right.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.RadioButton_left_right)
        Me.Panel3.Controls.Add(Me.RadioButton_right_left)
        Me.Panel3.Location = New System.Drawing.Point(12, 141)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(138, 53)
        Me.Panel3.TabIndex = 116
        '
        'Survey_band_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(305, 203)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Button_TRANSFER_TO_AUTOCAD)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Survey_band_form"
        Me.Text = "Survey Band"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button_TRANSFER_TO_AUTOCAD As System.Windows.Forms.Button
    Friend WithEvents Button_read_from_Excel As System.Windows.Forms.Button
    Friend WithEvents TextBox_row_start As System.Windows.Forms.MaskedTextBox
    Friend WithEvents TextBox_row_end As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents RadioButton_right_left As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_left_right As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
End Class
