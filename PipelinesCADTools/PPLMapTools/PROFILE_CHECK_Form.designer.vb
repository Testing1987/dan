<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PROFILE_CHECK_Form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PROFILE_CHECK_Form))
        Me.Button_scale_PICK = New System.Windows.Forms.Button()
        Me.TextBox_horiz_scale = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_vert_scale = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_printing_scale = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button_LABEL_POSITION = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_arrrow_size = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox_text_size = New System.Windows.Forms.TextBox()
        Me.TextBox_X = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox_Y = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox_dog_length = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox_gap = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBox_chainage_prefix = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_elevation_prefix = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TextBox_spacing = New System.Windows.Forms.TextBox()
        Me.Button_transfer_XL = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button_scale_PICK
        '
        Me.Button_scale_PICK.BackColor = System.Drawing.Color.White
        Me.Button_scale_PICK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button_scale_PICK.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Button_scale_PICK.ForeColor = System.Drawing.Color.Black
        Me.Button_scale_PICK.Location = New System.Drawing.Point(12, 110)
        Me.Button_scale_PICK.Name = "Button_scale_PICK"
        Me.Button_scale_PICK.Size = New System.Drawing.Size(168, 26)
        Me.Button_scale_PICK.TabIndex = 0
        Me.Button_scale_PICK.Text = "Scales Pick - Check"
        Me.Button_scale_PICK.UseVisualStyleBackColor = False
        '
        'TextBox_horiz_scale
        '
        Me.TextBox_horiz_scale.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TextBox_horiz_scale.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_horiz_scale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_horiz_scale.Location = New System.Drawing.Point(116, 11)
        Me.TextBox_horiz_scale.Name = "TextBox_horiz_scale"
        Me.TextBox_horiz_scale.Size = New System.Drawing.Size(64, 22)
        Me.TextBox_horiz_scale.TabIndex = 1
        Me.TextBox_horiz_scale.Text = "1000"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Gainsboro
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(12, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Horizontal Scale"
        '
        'TextBox_vert_scale
        '
        Me.TextBox_vert_scale.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TextBox_vert_scale.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_vert_scale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_vert_scale.Location = New System.Drawing.Point(116, 47)
        Me.TextBox_vert_scale.Name = "TextBox_vert_scale"
        Me.TextBox_vert_scale.Size = New System.Drawing.Size(64, 22)
        Me.TextBox_vert_scale.TabIndex = 1
        Me.TextBox_vert_scale.Text = "1000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Gainsboro
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(12, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Vertical Scale"
        '
        'TextBox_printing_scale
        '
        Me.TextBox_printing_scale.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_printing_scale.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_printing_scale.ForeColor = System.Drawing.Color.Black
        Me.TextBox_printing_scale.Location = New System.Drawing.Point(116, 82)
        Me.TextBox_printing_scale.Name = "TextBox_printing_scale"
        Me.TextBox_printing_scale.Size = New System.Drawing.Size(64, 22)
        Me.TextBox_printing_scale.TabIndex = 1
        Me.TextBox_printing_scale.Text = "1000"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Gainsboro
        Me.Label4.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(12, 81)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 16)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Printing Scale"
        '
        'Button_LABEL_POSITION
        '
        Me.Button_LABEL_POSITION.BackColor = System.Drawing.Color.White
        Me.Button_LABEL_POSITION.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button_LABEL_POSITION.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Button_LABEL_POSITION.ForeColor = System.Drawing.Color.Black
        Me.Button_LABEL_POSITION.Location = New System.Drawing.Point(370, 187)
        Me.Button_LABEL_POSITION.Name = "Button_LABEL_POSITION"
        Me.Button_LABEL_POSITION.Size = New System.Drawing.Size(93, 49)
        Me.Button_LABEL_POSITION.TabIndex = 0
        Me.Button_LABEL_POSITION.Text = "Label Position"
        Me.Button_LABEL_POSITION.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 231)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 16)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Arrow Size"
        '
        'TextBox_arrrow_size
        '
        Me.TextBox_arrrow_size.Location = New System.Drawing.Point(41, 250)
        Me.TextBox_arrrow_size.Name = "TextBox_arrrow_size"
        Me.TextBox_arrrow_size.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_arrrow_size.TabIndex = 4
        Me.TextBox_arrrow_size.Text = "1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(347, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 16)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Text Size"
        '
        'TextBox_text_size
        '
        Me.TextBox_text_size.Location = New System.Drawing.Point(404, 12)
        Me.TextBox_text_size.Name = "TextBox_text_size"
        Me.TextBox_text_size.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_text_size.TabIndex = 4
        Me.TextBox_text_size.Text = "1"
        '
        'TextBox_X
        '
        Me.TextBox_X.Location = New System.Drawing.Point(238, 133)
        Me.TextBox_X.Name = "TextBox_X"
        Me.TextBox_X.Size = New System.Drawing.Size(32, 21)
        Me.TextBox_X.TabIndex = 5
        Me.TextBox_X.Text = "5"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(214, 137)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(22, 16)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "∆X"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(214, 164)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(22, 16)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "∆Y"
        '
        'TextBox_Y
        '
        Me.TextBox_Y.Location = New System.Drawing.Point(238, 160)
        Me.TextBox_Y.Name = "TextBox_Y"
        Me.TextBox_Y.Size = New System.Drawing.Size(32, 21)
        Me.TextBox_Y.TabIndex = 5
        Me.TextBox_Y.Text = "5"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(235, 4)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 16)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "Length"
        '
        'TextBox_dog_length
        '
        Me.TextBox_dog_length.Location = New System.Drawing.Point(238, 23)
        Me.TextBox_dog_length.Name = "TextBox_dog_length"
        Me.TextBox_dog_length.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_dog_length.TabIndex = 4
        Me.TextBox_dog_length.Text = "1"
        Me.TextBox_dog_length.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(286, 44)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(26, 16)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Gap"
        '
        'TextBox_gap
        '
        Me.TextBox_gap.Location = New System.Drawing.Point(270, 63)
        Me.TextBox_gap.Name = "TextBox_gap"
        Me.TextBox_gap.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_gap.TabIndex = 4
        Me.TextBox_gap.Text = "1"
        Me.TextBox_gap.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(193, 190)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(71, 16)
        Me.Label10.TabIndex = 3
        Me.Label10.Text = "Station Prefix"
        '
        'TextBox_chainage_prefix
        '
        Me.TextBox_chainage_prefix.Location = New System.Drawing.Point(270, 187)
        Me.TextBox_chainage_prefix.Name = "TextBox_chainage_prefix"
        Me.TextBox_chainage_prefix.Size = New System.Drawing.Size(84, 21)
        Me.TextBox_chainage_prefix.TabIndex = 5
        Me.TextBox_chainage_prefix.Text = "STATION ="
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(183, 215)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(81, 16)
        Me.Label11.TabIndex = 3
        Me.Label11.Text = "Elevation Prefix"
        '
        'TextBox_elevation_prefix
        '
        Me.TextBox_elevation_prefix.Location = New System.Drawing.Point(270, 215)
        Me.TextBox_elevation_prefix.Name = "TextBox_elevation_prefix"
        Me.TextBox_elevation_prefix.Size = New System.Drawing.Size(84, 21)
        Me.TextBox_elevation_prefix.TabIndex = 5
        Me.TextBox_elevation_prefix.Text = "ELEV."
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(218, 274)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 16)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Spacing"
        '
        'TextBox_spacing
        '
        Me.TextBox_spacing.Location = New System.Drawing.Point(270, 271)
        Me.TextBox_spacing.Name = "TextBox_spacing"
        Me.TextBox_spacing.Size = New System.Drawing.Size(42, 21)
        Me.TextBox_spacing.TabIndex = 4
        Me.TextBox_spacing.Text = "50"
        Me.TextBox_spacing.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_transfer_XL
        '
        Me.Button_transfer_XL.Location = New System.Drawing.Point(324, 263)
        Me.Button_transfer_XL.Name = "Button_transfer_XL"
        Me.Button_transfer_XL.Size = New System.Drawing.Size(139, 38)
        Me.Button_transfer_XL.TabIndex = 6
        Me.Button_transfer_XL.Text = "Read HDD to Excel"
        Me.Button_transfer_XL.UseVisualStyleBackColor = True
        '
        'PROFILE_CHECK_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(475, 328)
        Me.Controls.Add(Me.Button_transfer_XL)
        Me.Controls.Add(Me.TextBox_Y)
        Me.Controls.Add(Me.TextBox_elevation_prefix)
        Me.Controls.Add(Me.TextBox_chainage_prefix)
        Me.Controls.Add(Me.TextBox_X)
        Me.Controls.Add(Me.TextBox_gap)
        Me.Controls.Add(Me.TextBox_spacing)
        Me.Controls.Add(Me.TextBox_dog_length)
        Me.Controls.Add(Me.TextBox_text_size)
        Me.Controls.Add(Me.TextBox_arrrow_size)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_vert_scale)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox_printing_scale)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox_horiz_scale)
        Me.Controls.Add(Me.Button_LABEL_POSITION)
        Me.Controls.Add(Me.Button_scale_PICK)
        Me.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "PROFILE_CHECK_Form"
        Me.Text = "PROFILE CHECK"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_scale_PICK As System.Windows.Forms.Button
    Friend WithEvents TextBox_horiz_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_vert_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox_printing_scale As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button_LABEL_POSITION As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox_arrrow_size As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_text_size As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_X As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Y As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox_dog_length As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBox_gap As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBox_chainage_prefix As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox_elevation_prefix As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextBox_spacing As System.Windows.Forms.TextBox
    Friend WithEvents Button_transfer_XL As System.Windows.Forms.Button
End Class
