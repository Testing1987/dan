<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class text2mtext_Form
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
        Me.TextBox_text_height = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button_change = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox_text_height
        '
        Me.TextBox_text_height.Location = New System.Drawing.Point(6, 25)
        Me.TextBox_text_height.Name = "TextBox_text_height"
        Me.TextBox_text_height.Size = New System.Drawing.Size(68, 20)
        Me.TextBox_text_height.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Text Height"
        '
        'Button_change
        '
        Me.Button_change.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_change.Location = New System.Drawing.Point(6, 60)
        Me.Button_change.Name = "Button_change"
        Me.Button_change.Size = New System.Drawing.Size(68, 33)
        Me.Button_change.TabIndex = 2
        Me.Button_change.Text = "Do It"
        Me.Button_change.UseVisualStyleBackColor = True
        '
        'text2mtext_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(227, 100)
        Me.Controls.Add(Me.Button_change)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_text_height)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "text2mtext_Form"
        Me.Text = "Text 2 Mtext"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox_text_height As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_change As System.Windows.Forms.Button
End Class
