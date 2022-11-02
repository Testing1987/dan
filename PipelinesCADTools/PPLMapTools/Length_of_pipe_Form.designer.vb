<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Length_of_pipe_Form
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
        Me.TextBox_REF_chainage = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_horizontal_Exxag = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_vertical_exag = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button_output = New System.Windows.Forms.Button()
        Me.TextBox_dwg_no = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_ref_ch_col = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_len_of_pipe = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox_start_row = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TextBox_REF_chainage
        '
        Me.TextBox_REF_chainage.Location = New System.Drawing.Point(19, 27)
        Me.TextBox_REF_chainage.Name = "TextBox_REF_chainage"
        Me.TextBox_REF_chainage.Size = New System.Drawing.Size(100, 21)
        Me.TextBox_REF_chainage.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(122, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Reference Chainage"
        '
        'TextBox_horizontal_Exxag
        '
        Me.TextBox_horizontal_Exxag.Location = New System.Drawing.Point(33, 73)
        Me.TextBox_horizontal_Exxag.Name = "TextBox_horizontal_Exxag"
        Me.TextBox_horizontal_Exxag.Size = New System.Drawing.Size(69, 21)
        Me.TextBox_horizontal_Exxag.TabIndex = 0
        Me.TextBox_horizontal_Exxag.Text = "1"
        Me.TextBox_horizontal_Exxag.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(142, 15)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Horizontal Exaggeration"
        '
        'TextBox_vertical_exag
        '
        Me.TextBox_vertical_exag.Location = New System.Drawing.Point(207, 73)
        Me.TextBox_vertical_exag.Name = "TextBox_vertical_exag"
        Me.TextBox_vertical_exag.Size = New System.Drawing.Size(69, 21)
        Me.TextBox_vertical_exag.TabIndex = 0
        Me.TextBox_vertical_exag.Text = "1"
        Me.TextBox_vertical_exag.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(186, 55)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 15)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Vertical Exaggeration"
        '
        'Button_output
        '
        Me.Button_output.Location = New System.Drawing.Point(255, 190)
        Me.Button_output.Name = "Button_output"
        Me.Button_output.Size = New System.Drawing.Size(118, 31)
        Me.Button_output.TabIndex = 2
        Me.Button_output.Text = "Output!"
        Me.Button_output.UseVisualStyleBackColor = True
        '
        'TextBox_dwg_no
        '
        Me.TextBox_dwg_no.Location = New System.Drawing.Point(18, 147)
        Me.TextBox_dwg_no.Name = "TextBox_dwg_no"
        Me.TextBox_dwg_no.Size = New System.Drawing.Size(69, 21)
        Me.TextBox_dwg_no.TabIndex = 0
        Me.TextBox_dwg_no.Text = "A"
        Me.TextBox_dwg_no.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 114)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(71, 30)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Drawing no" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_ref_ch_col
        '
        Me.TextBox_ref_ch_col.Location = New System.Drawing.Point(143, 147)
        Me.TextBox_ref_ch_col.Name = "TextBox_ref_ch_col"
        Me.TextBox_ref_ch_col.Size = New System.Drawing.Size(69, 21)
        Me.TextBox_ref_ch_col.TabIndex = 0
        Me.TextBox_ref_ch_col.Text = "B"
        Me.TextBox_ref_ch_col.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(114, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(122, 30)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Reference Chainage" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_len_of_pipe
        '
        Me.TextBox_len_of_pipe.Location = New System.Drawing.Point(255, 147)
        Me.TextBox_len_of_pipe.Name = "TextBox_len_of_pipe"
        Me.TextBox_len_of_pipe.Size = New System.Drawing.Size(69, 21)
        Me.TextBox_len_of_pipe.TabIndex = 0
        Me.TextBox_len_of_pipe.Text = "C"
        Me.TextBox_len_of_pipe.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(252, 114)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 30)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Length of Pipe" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Column"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox_start_row
        '
        Me.TextBox_start_row.Location = New System.Drawing.Point(37, 195)
        Me.TextBox_start_row.Name = "TextBox_start_row"
        Me.TextBox_start_row.Size = New System.Drawing.Size(69, 21)
        Me.TextBox_start_row.TabIndex = 0
        Me.TextBox_start_row.Text = "1"
        Me.TextBox_start_row.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(41, 177)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(63, 15)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Excel row"
        '
        'Length_of_pipe_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(405, 227)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Button_output)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox_len_of_pipe)
        Me.Controls.Add(Me.TextBox_ref_ch_col)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox_dwg_no)
        Me.Controls.Add(Me.TextBox_start_row)
        Me.Controls.Add(Me.TextBox_vertical_exag)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_horizontal_Exxag)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_REF_chainage)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Length_of_pipe_Form"
        Me.Text = "Length of pipe"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_REF_chainage As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_horizontal_Exxag As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_vertical_exag As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button_output As System.Windows.Forms.Button
    Friend WithEvents TextBox_dwg_no As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox_ref_ch_col As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_len_of_pipe As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox_start_row As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
End Class
