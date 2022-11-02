<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Multiple_drawings_change_form
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
        Me.components = New System.ComponentModel.Container()
        Me.TabPageVW2DWG = New System.Windows.Forms.TabPage()
        Me.Button_remove_items_list = New System.Windows.Forms.Button()
        Me.Button_block_modify = New System.Windows.Forms.Button()
        Me.ListBox_DWG = New System.Windows.Forms.ListBox()
        Me.Button_load_DWG = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage_xref = New System.Windows.Forms.TabPage()
        Me.Panel_xref_path = New System.Windows.Forms.Panel()
        Me.Panel_xref_name = New System.Windows.Forms.Panel()
        Me.Button_read_Xref = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TabPageVW2DWG.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_xref.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPageVW2DWG
        '
        Me.TabPageVW2DWG.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPageVW2DWG.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPageVW2DWG.Controls.Add(Me.Button_remove_items_list)
        Me.TabPageVW2DWG.Controls.Add(Me.Button_block_modify)
        Me.TabPageVW2DWG.Controls.Add(Me.ListBox_DWG)
        Me.TabPageVW2DWG.Controls.Add(Me.Button_load_DWG)
        Me.TabPageVW2DWG.Location = New System.Drawing.Point(4, 25)
        Me.TabPageVW2DWG.Name = "TabPageVW2DWG"
        Me.TabPageVW2DWG.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageVW2DWG.Size = New System.Drawing.Size(799, 537)
        Me.TabPageVW2DWG.TabIndex = 11
        Me.TabPageVW2DWG.Text = "DWG List"
        '
        'Button_remove_items_list
        '
        Me.Button_remove_items_list.BackColor = System.Drawing.Color.Orange
        Me.Button_remove_items_list.Location = New System.Drawing.Point(1, 486)
        Me.Button_remove_items_list.Name = "Button_remove_items_list"
        Me.Button_remove_items_list.Size = New System.Drawing.Size(119, 41)
        Me.Button_remove_items_list.TabIndex = 136
        Me.Button_remove_items_list.Text = "Remove selected"
        Me.Button_remove_items_list.UseVisualStyleBackColor = False
        '
        'Button_block_modify
        '
        Me.Button_block_modify.Location = New System.Drawing.Point(6, 48)
        Me.Button_block_modify.Name = "Button_block_modify"
        Me.Button_block_modify.Size = New System.Drawing.Size(116, 44)
        Me.Button_block_modify.TabIndex = 2
        Me.Button_block_modify.Text = "Block Modify"
        Me.Button_block_modify.UseVisualStyleBackColor = True
        '
        'ListBox_DWG
        '
        Me.ListBox_DWG.BackColor = System.Drawing.Color.White
        Me.ListBox_DWG.ForeColor = System.Drawing.Color.Black
        Me.ListBox_DWG.FormattingEnabled = True
        Me.ListBox_DWG.HorizontalScrollbar = True
        Me.ListBox_DWG.ItemHeight = 16
        Me.ListBox_DWG.Location = New System.Drawing.Point(128, 6)
        Me.ListBox_DWG.Name = "ListBox_DWG"
        Me.ListBox_DWG.ScrollAlwaysVisible = True
        Me.ListBox_DWG.Size = New System.Drawing.Size(308, 532)
        Me.ListBox_DWG.Sorted = True
        Me.ListBox_DWG.TabIndex = 1
        '
        'Button_load_DWG
        '
        Me.Button_load_DWG.Location = New System.Drawing.Point(6, 6)
        Me.Button_load_DWG.Name = "Button_load_DWG"
        Me.Button_load_DWG.Size = New System.Drawing.Size(116, 36)
        Me.Button_load_DWG.TabIndex = 0
        Me.Button_load_DWG.Text = "Load Drawings"
        Me.Button_load_DWG.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_xref)
        Me.TabControl1.Controls.Add(Me.TabPageVW2DWG)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(807, 566)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage_xref
        '
        Me.TabPage_xref.BackColor = System.Drawing.Color.Gainsboro
        Me.TabPage_xref.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage_xref.Controls.Add(Me.Panel_xref_path)
        Me.TabPage_xref.Controls.Add(Me.Panel_xref_name)
        Me.TabPage_xref.Controls.Add(Me.Button_read_Xref)
        Me.TabPage_xref.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_xref.Name = "TabPage_xref"
        Me.TabPage_xref.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_xref.Size = New System.Drawing.Size(799, 537)
        Me.TabPage_xref.TabIndex = 13
        Me.TabPage_xref.Text = "XREF"
        '
        'Panel_xref_path
        '
        Me.Panel_xref_path.AutoScroll = True
        Me.Panel_xref_path.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_xref_path.Location = New System.Drawing.Point(205, 6)
        Me.Panel_xref_path.Name = "Panel_xref_path"
        Me.Panel_xref_path.Size = New System.Drawing.Size(584, 471)
        Me.Panel_xref_path.TabIndex = 1
        '
        'Panel_xref_name
        '
        Me.Panel_xref_name.AutoScroll = True
        Me.Panel_xref_name.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_xref_name.Location = New System.Drawing.Point(1, 6)
        Me.Panel_xref_name.Name = "Panel_xref_name"
        Me.Panel_xref_name.Size = New System.Drawing.Size(198, 471)
        Me.Panel_xref_name.TabIndex = 1
        '
        'Button_read_Xref
        '
        Me.Button_read_Xref.Location = New System.Drawing.Point(1, 488)
        Me.Button_read_Xref.Name = "Button_read_Xref"
        Me.Button_read_Xref.Size = New System.Drawing.Size(165, 44)
        Me.Button_read_Xref.TabIndex = 0
        Me.Button_read_Xref.Text = "Read existing XREFs"
        Me.Button_read_Xref.UseVisualStyleBackColor = True
        '
        'Multiple_drawings_change_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(831, 585)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "Multiple_drawings_change_form"
        Me.Text = "Multiple drawings change form"
        Me.TabPageVW2DWG.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_xref.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabPageVW2DWG As System.Windows.Forms.TabPage
    Friend WithEvents ListBox_DWG As System.Windows.Forms.ListBox
    Friend WithEvents Button_load_DWG As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents Button_block_modify As System.Windows.Forms.Button
    Friend WithEvents Button_remove_items_list As System.Windows.Forms.Button
    Friend WithEvents TabPage_xref As System.Windows.Forms.TabPage
    Friend WithEvents Button_read_Xref As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Panel_xref_name As System.Windows.Forms.Panel
    Friend WithEvents Panel_xref_path As System.Windows.Forms.Panel
End Class
