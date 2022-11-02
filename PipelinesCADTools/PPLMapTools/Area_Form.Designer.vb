<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Area_Form
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
        Me.RadioButton_sqft_to_AC = New System.Windows.Forms.RadioButton()
        Me.RadioButton_no_conversion = New System.Windows.Forms.RadioButton()
        Me.Button_calculate = New System.Windows.Forms.Button()
        Me.TextBox_decimals = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_result = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.RadioButton_sqft_to_AC)
        Me.Panel1.Controls.Add(Me.RadioButton_no_conversion)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 60)
        Me.Panel1.TabIndex = 0
        '
        'RadioButton_sqft_to_AC
        '
        Me.RadioButton_sqft_to_AC.AutoSize = True
        Me.RadioButton_sqft_to_AC.Location = New System.Drawing.Point(3, 29)
        Me.RadioButton_sqft_to_AC.Name = "RadioButton_sqft_to_AC"
        Me.RadioButton_sqft_to_AC.Size = New System.Drawing.Size(116, 20)
        Me.RadioButton_sqft_to_AC.TabIndex = 1
        Me.RadioButton_sqft_to_AC.Text = "SQFT to Acres"
        Me.RadioButton_sqft_to_AC.UseVisualStyleBackColor = True
        '
        'RadioButton_no_conversion
        '
        Me.RadioButton_no_conversion.AutoSize = True
        Me.RadioButton_no_conversion.Checked = True
        Me.RadioButton_no_conversion.Location = New System.Drawing.Point(3, 3)
        Me.RadioButton_no_conversion.Name = "RadioButton_no_conversion"
        Me.RadioButton_no_conversion.Size = New System.Drawing.Size(116, 20)
        Me.RadioButton_no_conversion.TabIndex = 1
        Me.RadioButton_no_conversion.TabStop = True
        Me.RadioButton_no_conversion.Text = "No conversion"
        Me.RadioButton_no_conversion.UseVisualStyleBackColor = True
        '
        'Button_calculate
        '
        Me.Button_calculate.Location = New System.Drawing.Point(12, 143)
        Me.Button_calculate.Name = "Button_calculate"
        Me.Button_calculate.Size = New System.Drawing.Size(200, 32)
        Me.Button_calculate.TabIndex = 1
        Me.Button_calculate.Text = "Calculate"
        Me.Button_calculate.UseVisualStyleBackColor = True
        '
        'TextBox_decimals
        '
        Me.TextBox_decimals.BackColor = System.Drawing.Color.White
        Me.TextBox_decimals.ForeColor = System.Drawing.Color.Black
        Me.TextBox_decimals.Location = New System.Drawing.Point(163, 78)
        Me.TextBox_decimals.Name = "TextBox_decimals"
        Me.TextBox_decimals.Size = New System.Drawing.Size(49, 22)
        Me.TextBox_decimals.TabIndex = 2
        Me.TextBox_decimals.Text = "5"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(54, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "No of Decimals"
        '
        'TextBox_result
        '
        Me.TextBox_result.Location = New System.Drawing.Point(102, 112)
        Me.TextBox_result.Name = "TextBox_result"
        Me.TextBox_result.Size = New System.Drawing.Size(110, 22)
        Me.TextBox_result.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(40, 115)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Result"
        '
        'Area_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(219, 185)
        Me.Controls.Add(Me.TextBox_result)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_decimals)
        Me.Controls.Add(Me.Button_calculate)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "Area_Form"
        Me.Text = "Area Form"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadioButton_sqft_to_AC As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_no_conversion As System.Windows.Forms.RadioButton
    Friend WithEvents Button_calculate As System.Windows.Forms.Button
    Friend WithEvents TextBox_decimals As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_result As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
End Class
