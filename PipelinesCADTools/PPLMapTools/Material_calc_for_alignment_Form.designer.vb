<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Material_calc_for_alignment_Form
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
        Me.ListBox_mat_type = New System.Windows.Forms.ListBox()
        Me.ListBox_lenghts = New System.Windows.Forms.ListBox()
        Me.Button_pick = New System.Windows.Forms.Button()
        Me.ListBox_Totals = New System.Windows.Forms.ListBox()
        Me.Button_calc_totals = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button_clear = New System.Windows.Forms.Button()
        Me.Button_TRANSFER_TOTALS = New System.Windows.Forms.Button()
        Me.Button_pick_all = New System.Windows.Forms.Button()
        Me.Button_calc_difference = New System.Windows.Forms.Button()
        Me.TextBox_diference = New System.Windows.Forms.TextBox()
        Me.TextBox_END_chainage = New System.Windows.Forms.TextBox()
        Me.TextBox_BEG_chainage = New System.Windows.Forms.TextBox()
        Me.Button_transfer_to_excel = New System.Windows.Forms.Button()
        Me.TextBox_ROW_START_XL = New System.Windows.Forms.TextBox()
        Me.ListBox_chainage1 = New System.Windows.Forms.ListBox()
        Me.ListBox_chainage2 = New System.Windows.Forms.ListBox()
        Me.ListBox_true_length = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CheckBox_US_style = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'ListBox_mat_type
        '
        Me.ListBox_mat_type.BackColor = System.Drawing.Color.White
        Me.ListBox_mat_type.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_mat_type.ForeColor = System.Drawing.Color.Black
        Me.ListBox_mat_type.FormattingEnabled = True
        Me.ListBox_mat_type.HorizontalScrollbar = True
        Me.ListBox_mat_type.ItemHeight = 14
        Me.ListBox_mat_type.Location = New System.Drawing.Point(172, 39)
        Me.ListBox_mat_type.Name = "ListBox_mat_type"
        Me.ListBox_mat_type.Size = New System.Drawing.Size(100, 340)
        Me.ListBox_mat_type.TabIndex = 112
        '
        'ListBox_lenghts
        '
        Me.ListBox_lenghts.BackColor = System.Drawing.Color.White
        Me.ListBox_lenghts.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_lenghts.ForeColor = System.Drawing.Color.Black
        Me.ListBox_lenghts.FormattingEnabled = True
        Me.ListBox_lenghts.HorizontalScrollbar = True
        Me.ListBox_lenghts.ItemHeight = 14
        Me.ListBox_lenghts.Location = New System.Drawing.Point(278, 39)
        Me.ListBox_lenghts.Name = "ListBox_lenghts"
        Me.ListBox_lenghts.Size = New System.Drawing.Size(104, 340)
        Me.ListBox_lenghts.TabIndex = 113
        '
        'Button_pick
        '
        Me.Button_pick.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_pick.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick.ForeColor = System.Drawing.Color.Black
        Me.Button_pick.Location = New System.Drawing.Point(8, 6)
        Me.Button_pick.Name = "Button_pick"
        Me.Button_pick.Size = New System.Drawing.Size(158, 39)
        Me.Button_pick.TabIndex = 114
        Me.Button_pick.Text = "Pick Values"
        Me.Button_pick.UseVisualStyleBackColor = False
        '
        'ListBox_Totals
        '
        Me.ListBox_Totals.BackColor = System.Drawing.Color.White
        Me.ListBox_Totals.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_Totals.ForeColor = System.Drawing.Color.Black
        Me.ListBox_Totals.FormattingEnabled = True
        Me.ListBox_Totals.HorizontalScrollbar = True
        Me.ListBox_Totals.ItemHeight = 14
        Me.ListBox_Totals.Location = New System.Drawing.Point(562, 39)
        Me.ListBox_Totals.Name = "ListBox_Totals"
        Me.ListBox_Totals.Size = New System.Drawing.Size(98, 340)
        Me.ListBox_Totals.TabIndex = 115
        '
        'Button_calc_totals
        '
        Me.Button_calc_totals.BackColor = System.Drawing.Color.LawnGreen
        Me.Button_calc_totals.Location = New System.Drawing.Point(775, 21)
        Me.Button_calc_totals.Name = "Button_calc_totals"
        Me.Button_calc_totals.Size = New System.Drawing.Size(110, 53)
        Me.Button_calc_totals.TabIndex = 116
        Me.Button_calc_totals.Text = "Calculate Totals"
        Me.Button_calc_totals.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(169, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 15)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "Material"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(275, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 15)
        Me.Label2.TabIndex = 117
        Me.Label2.Text = "Length"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(798, 105)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 15)
        Me.Label3.TabIndex = 117
        Me.Label3.Text = "Row Excel"
        '
        'Button_clear
        '
        Me.Button_clear.BackColor = System.Drawing.Color.Red
        Me.Button_clear.ForeColor = System.Drawing.Color.White
        Me.Button_clear.Location = New System.Drawing.Point(784, 343)
        Me.Button_clear.Name = "Button_clear"
        Me.Button_clear.Size = New System.Drawing.Size(92, 36)
        Me.Button_clear.TabIndex = 118
        Me.Button_clear.Text = "Clear"
        Me.Button_clear.UseVisualStyleBackColor = False
        '
        'Button_TRANSFER_TOTALS
        '
        Me.Button_TRANSFER_TOTALS.Location = New System.Drawing.Point(784, 217)
        Me.Button_TRANSFER_TOTALS.Name = "Button_TRANSFER_TOTALS"
        Me.Button_TRANSFER_TOTALS.Size = New System.Drawing.Size(92, 48)
        Me.Button_TRANSFER_TOTALS.TabIndex = 116
        Me.Button_TRANSFER_TOTALS.Text = "Change totals"
        Me.Button_TRANSFER_TOTALS.UseVisualStyleBackColor = True
        '
        'Button_pick_all
        '
        Me.Button_pick_all.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_pick_all.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_pick_all.ForeColor = System.Drawing.Color.Black
        Me.Button_pick_all.Location = New System.Drawing.Point(8, 197)
        Me.Button_pick_all.Name = "Button_pick_all"
        Me.Button_pick_all.Size = New System.Drawing.Size(142, 39)
        Me.Button_pick_all.TabIndex = 114
        Me.Button_pick_all.Text = "Pick From Blocks"
        Me.Button_pick_all.UseVisualStyleBackColor = False
        '
        'Button_calc_difference
        '
        Me.Button_calc_difference.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_calc_difference.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_calc_difference.ForeColor = System.Drawing.Color.Black
        Me.Button_calc_difference.Location = New System.Drawing.Point(12, 337)
        Me.Button_calc_difference.Name = "Button_calc_difference"
        Me.Button_calc_difference.Size = New System.Drawing.Size(122, 41)
        Me.Button_calc_difference.TabIndex = 121
        Me.Button_calc_difference.Text = "Pick Stations"
        Me.Button_calc_difference.UseVisualStyleBackColor = False
        '
        'TextBox_diference
        '
        Me.TextBox_diference.BackColor = System.Drawing.Color.PeachPuff
        Me.TextBox_diference.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_diference.ForeColor = System.Drawing.Color.Black
        Me.TextBox_diference.Location = New System.Drawing.Point(12, 309)
        Me.TextBox_diference.Name = "TextBox_diference"
        Me.TextBox_diference.ReadOnly = True
        Me.TextBox_diference.Size = New System.Drawing.Size(122, 22)
        Me.TextBox_diference.TabIndex = 122
        '
        'TextBox_END_chainage
        '
        Me.TextBox_END_chainage.BackColor = System.Drawing.Color.White
        Me.TextBox_END_chainage.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_END_chainage.ForeColor = System.Drawing.Color.Black
        Me.TextBox_END_chainage.Location = New System.Drawing.Point(12, 281)
        Me.TextBox_END_chainage.Name = "TextBox_END_chainage"
        Me.TextBox_END_chainage.Size = New System.Drawing.Size(122, 22)
        Me.TextBox_END_chainage.TabIndex = 120
        '
        'TextBox_BEG_chainage
        '
        Me.TextBox_BEG_chainage.BackColor = System.Drawing.Color.White
        Me.TextBox_BEG_chainage.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_BEG_chainage.ForeColor = System.Drawing.Color.Black
        Me.TextBox_BEG_chainage.Location = New System.Drawing.Point(12, 253)
        Me.TextBox_BEG_chainage.Name = "TextBox_BEG_chainage"
        Me.TextBox_BEG_chainage.Size = New System.Drawing.Size(122, 22)
        Me.TextBox_BEG_chainage.TabIndex = 119
        '
        'Button_transfer_to_excel
        '
        Me.Button_transfer_to_excel.Location = New System.Drawing.Point(784, 153)
        Me.Button_transfer_to_excel.Name = "Button_transfer_to_excel"
        Me.Button_transfer_to_excel.Size = New System.Drawing.Size(92, 42)
        Me.Button_transfer_to_excel.TabIndex = 116
        Me.Button_transfer_to_excel.Text = "Transfer to Excel"
        Me.Button_transfer_to_excel.UseVisualStyleBackColor = True
        '
        'TextBox_ROW_START_XL
        '
        Me.TextBox_ROW_START_XL.BackColor = System.Drawing.Color.White
        Me.TextBox_ROW_START_XL.ForeColor = System.Drawing.Color.Black
        Me.TextBox_ROW_START_XL.Location = New System.Drawing.Point(801, 123)
        Me.TextBox_ROW_START_XL.Name = "TextBox_ROW_START_XL"
        Me.TextBox_ROW_START_XL.Size = New System.Drawing.Size(53, 21)
        Me.TextBox_ROW_START_XL.TabIndex = 123
        Me.TextBox_ROW_START_XL.Text = "1"
        Me.TextBox_ROW_START_XL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ListBox_chainage1
        '
        Me.ListBox_chainage1.BackColor = System.Drawing.Color.White
        Me.ListBox_chainage1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_chainage1.ForeColor = System.Drawing.Color.Black
        Me.ListBox_chainage1.FormattingEnabled = True
        Me.ListBox_chainage1.HorizontalScrollbar = True
        Me.ListBox_chainage1.ItemHeight = 14
        Me.ListBox_chainage1.Location = New System.Drawing.Point(388, 38)
        Me.ListBox_chainage1.Name = "ListBox_chainage1"
        Me.ListBox_chainage1.Size = New System.Drawing.Size(81, 340)
        Me.ListBox_chainage1.TabIndex = 113
        '
        'ListBox_chainage2
        '
        Me.ListBox_chainage2.BackColor = System.Drawing.Color.White
        Me.ListBox_chainage2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_chainage2.ForeColor = System.Drawing.Color.Black
        Me.ListBox_chainage2.FormattingEnabled = True
        Me.ListBox_chainage2.HorizontalScrollbar = True
        Me.ListBox_chainage2.ItemHeight = 14
        Me.ListBox_chainage2.Location = New System.Drawing.Point(475, 39)
        Me.ListBox_chainage2.Name = "ListBox_chainage2"
        Me.ListBox_chainage2.Size = New System.Drawing.Size(81, 340)
        Me.ListBox_chainage2.TabIndex = 113
        '
        'ListBox_true_length
        '
        Me.ListBox_true_length.BackColor = System.Drawing.Color.White
        Me.ListBox_true_length.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.ListBox_true_length.ForeColor = System.Drawing.Color.Black
        Me.ListBox_true_length.FormattingEnabled = True
        Me.ListBox_true_length.HorizontalScrollbar = True
        Me.ListBox_true_length.ItemHeight = 14
        Me.ListBox_true_length.Location = New System.Drawing.Point(666, 39)
        Me.ListBox_true_length.Name = "ListBox_true_length"
        Me.ListBox_true_length.Size = New System.Drawing.Size(93, 340)
        Me.ListBox_true_length.TabIndex = 115
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(663, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 30)
        Me.Label4.TabIndex = 117
        Me.Label4.Text = "Totals with " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "True Length"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(559, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 15)
        Me.Label5.TabIndex = 117
        Me.Label5.Text = "Totals"
        '
        'CheckBox_US_style
        '
        Me.CheckBox_US_style.AutoSize = True
        Me.CheckBox_US_style.Location = New System.Drawing.Point(12, 105)
        Me.CheckBox_US_style.Name = "CheckBox_US_style"
        Me.CheckBox_US_style.Size = New System.Drawing.Size(97, 19)
        Me.CheckBox_US_style.TabIndex = 124
        Me.CheckBox_US_style.Text = "US style mat"
        Me.CheckBox_US_style.UseVisualStyleBackColor = True
        '
        'Material_calc_for_alignment_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(898, 389)
        Me.Controls.Add(Me.CheckBox_US_style)
        Me.Controls.Add(Me.TextBox_ROW_START_XL)
        Me.Controls.Add(Me.Button_calc_difference)
        Me.Controls.Add(Me.TextBox_diference)
        Me.Controls.Add(Me.TextBox_END_chainage)
        Me.Controls.Add(Me.TextBox_BEG_chainage)
        Me.Controls.Add(Me.Button_clear)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button_transfer_to_excel)
        Me.Controls.Add(Me.Button_TRANSFER_TOTALS)
        Me.Controls.Add(Me.Button_calc_totals)
        Me.Controls.Add(Me.ListBox_true_length)
        Me.Controls.Add(Me.ListBox_Totals)
        Me.Controls.Add(Me.Button_pick_all)
        Me.Controls.Add(Me.Button_pick)
        Me.Controls.Add(Me.ListBox_mat_type)
        Me.Controls.Add(Me.ListBox_chainage2)
        Me.Controls.Add(Me.ListBox_chainage1)
        Me.Controls.Add(Me.ListBox_lenghts)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Material_calc_for_alignment_Form"
        Me.Text = "Material Types"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBox_mat_type As System.Windows.Forms.ListBox
    Friend WithEvents ListBox_lenghts As System.Windows.Forms.ListBox
    Friend WithEvents Button_pick As System.Windows.Forms.Button
    Friend WithEvents ListBox_Totals As System.Windows.Forms.ListBox
    Friend WithEvents Button_calc_totals As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button_clear As System.Windows.Forms.Button
    Friend WithEvents Button_TRANSFER_TOTALS As System.Windows.Forms.Button
    Friend WithEvents Button_pick_all As System.Windows.Forms.Button
    Friend WithEvents Button_calc_difference As System.Windows.Forms.Button
    Friend WithEvents TextBox_diference As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_END_chainage As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_BEG_chainage As System.Windows.Forms.TextBox
    Friend WithEvents Button_transfer_to_excel As System.Windows.Forms.Button
    Friend WithEvents TextBox_ROW_START_XL As System.Windows.Forms.TextBox
    Friend WithEvents ListBox_chainage1 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox_chainage2 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox_true_length As System.Windows.Forms.ListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_US_style As System.Windows.Forms.CheckBox
End Class
