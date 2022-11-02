namespace Alignment_mdi
{
    partial class OD_layer_change_form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Button_read_OD = new System.Windows.Forms.Button();
            this.comboBox_OD1 = new System.Windows.Forms.ComboBox();
            this.button_read_excel_column = new System.Windows.Forms.Button();
            this.textBox_excel_column = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_start = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_end = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox_Layers = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button_change_layer = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.panel_Excel = new System.Windows.Forms.Panel();
            this.checkBox_use_null_values = new System.Windows.Forms.CheckBox();
            this.comboBox_OD2 = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.button_blocks = new System.Windows.Forms.Button();
            this.comboBox_blocks = new System.Windows.Forms.ComboBox();
            this.comboBox_block_atr1 = new System.Windows.Forms.ComboBox();
            this.button_change_attribute_value = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.button_load_OD = new System.Windows.Forms.Button();
            this.panel_Excel.SuspendLayout();
            this.SuspendLayout();
            // 
            // Button_read_OD
            // 
            this.Button_read_OD.Location = new System.Drawing.Point(12, 12);
            this.Button_read_OD.Name = "Button_read_OD";
            this.Button_read_OD.Size = new System.Drawing.Size(197, 44);
            this.Button_read_OD.TabIndex = 1;
            this.Button_read_OD.Text = "Read Object Data\r\nLoad Definitions to ComboBox";
            this.Button_read_OD.UseVisualStyleBackColor = true;
            this.Button_read_OD.Click += new System.EventHandler(this.Button_read_OD_Click);
            // 
            // comboBox_OD1
            // 
            this.comboBox_OD1.BackColor = System.Drawing.Color.White;
            this.comboBox_OD1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_OD1.ForeColor = System.Drawing.Color.Black;
            this.comboBox_OD1.FormattingEnabled = true;
            this.comboBox_OD1.Location = new System.Drawing.Point(29, 85);
            this.comboBox_OD1.Name = "comboBox_OD1";
            this.comboBox_OD1.Size = new System.Drawing.Size(180, 23);
            this.comboBox_OD1.TabIndex = 2;
            // 
            // button_read_excel_column
            // 
            this.button_read_excel_column.Location = new System.Drawing.Point(3, 86);
            this.button_read_excel_column.Name = "button_read_excel_column";
            this.button_read_excel_column.Size = new System.Drawing.Size(184, 44);
            this.button_read_excel_column.TabIndex = 1;
            this.button_read_excel_column.Text = "Read from Excel";
            this.button_read_excel_column.UseVisualStyleBackColor = true;
            this.button_read_excel_column.Click += new System.EventHandler(this.Button_read_Excel_Click);
            // 
            // textBox_excel_column
            // 
            this.textBox_excel_column.BackColor = System.Drawing.Color.White;
            this.textBox_excel_column.ForeColor = System.Drawing.Color.Black;
            this.textBox_excel_column.Location = new System.Drawing.Point(97, 4);
            this.textBox_excel_column.Name = "textBox_excel_column";
            this.textBox_excel_column.Size = new System.Drawing.Size(46, 21);
            this.textBox_excel_column.TabIndex = 3;
            this.textBox_excel_column.Text = "C";
            this.textBox_excel_column.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(3, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Excel Colum";
            // 
            // textBox_start
            // 
            this.textBox_start.BackColor = System.Drawing.Color.White;
            this.textBox_start.ForeColor = System.Drawing.Color.Black;
            this.textBox_start.Location = new System.Drawing.Point(97, 31);
            this.textBox_start.Name = "textBox_start";
            this.textBox_start.Size = new System.Drawing.Size(46, 21);
            this.textBox_start.TabIndex = 3;
            this.textBox_start.Text = "2";
            this.textBox_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Location = new System.Drawing.Point(3, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Start row";
            // 
            // textBox_end
            // 
            this.textBox_end.BackColor = System.Drawing.Color.White;
            this.textBox_end.ForeColor = System.Drawing.Color.Black;
            this.textBox_end.Location = new System.Drawing.Point(97, 59);
            this.textBox_end.Name = "textBox_end";
            this.textBox_end.Size = new System.Drawing.Size(46, 21);
            this.textBox_end.TabIndex = 3;
            this.textBox_end.Text = "38";
            this.textBox_end.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label3.Location = new System.Drawing.Point(3, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "End row";
            // 
            // comboBox_Layers
            // 
            this.comboBox_Layers.BackColor = System.Drawing.Color.White;
            this.comboBox_Layers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Layers.ForeColor = System.Drawing.Color.Black;
            this.comboBox_Layers.FormattingEnabled = true;
            this.comboBox_Layers.Location = new System.Drawing.Point(119, 444);
            this.comboBox_Layers.Name = "comboBox_Layers";
            this.comboBox_Layers.Size = new System.Drawing.Size(197, 23);
            this.comboBox_Layers.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(119, 426);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(184, 15);
            this.label4.TabIndex = 4;
            this.label4.Text = "Autocad Layer for matching OD";
            // 
            // button_change_layer
            // 
            this.button_change_layer.Location = new System.Drawing.Point(119, 484);
            this.button_change_layer.Name = "button_change_layer";
            this.button_change_layer.Size = new System.Drawing.Size(197, 44);
            this.button_change_layer.TabIndex = 1;
            this.button_change_layer.Text = "Change Layer";
            this.button_change_layer.UseVisualStyleBackColor = true;
            this.button_change_layer.Click += new System.EventHandler(this.Change_layer_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 61);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(109, 15);
            this.label5.TabIndex = 4;
            this.label5.Text = "Object Data Fields";
            // 
            // panel_Excel
            // 
            this.panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel_Excel.Controls.Add(this.button_read_excel_column);
            this.panel_Excel.Controls.Add(this.textBox_excel_column);
            this.panel_Excel.Controls.Add(this.label1);
            this.panel_Excel.Controls.Add(this.label3);
            this.panel_Excel.Controls.Add(this.textBox_start);
            this.panel_Excel.Controls.Add(this.textBox_end);
            this.panel_Excel.Controls.Add(this.label2);
            this.panel_Excel.Location = new System.Drawing.Point(342, 385);
            this.panel_Excel.Name = "panel_Excel";
            this.panel_Excel.Size = new System.Drawing.Size(194, 143);
            this.panel_Excel.TabIndex = 5;
            // 
            // checkBox_use_null_values
            // 
            this.checkBox_use_null_values.AutoSize = true;
            this.checkBox_use_null_values.Location = new System.Drawing.Point(119, 401);
            this.checkBox_use_null_values.Name = "checkBox_use_null_values";
            this.checkBox_use_null_values.Size = new System.Drawing.Size(144, 19);
            this.checkBox_use_null_values.TabIndex = 6;
            this.checkBox_use_null_values.Text = "Check for null values";
            this.checkBox_use_null_values.UseVisualStyleBackColor = true;
            // 
            // comboBox_OD2
            // 
            this.comboBox_OD2.BackColor = System.Drawing.Color.White;
            this.comboBox_OD2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_OD2.ForeColor = System.Drawing.Color.Black;
            this.comboBox_OD2.FormattingEnabled = true;
            this.comboBox_OD2.Location = new System.Drawing.Point(29, 114);
            this.comboBox_OD2.Name = "comboBox_OD2";
            this.comboBox_OD2.Size = new System.Drawing.Size(180, 23);
            this.comboBox_OD2.TabIndex = 2;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 88);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(14, 15);
            this.label6.TabIndex = 4;
            this.label6.Text = "1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 117);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(14, 15);
            this.label7.TabIndex = 4;
            this.label7.Text = "2";
            // 
            // button_blocks
            // 
            this.button_blocks.Location = new System.Drawing.Point(287, 12);
            this.button_blocks.Name = "button_blocks";
            this.button_blocks.Size = new System.Drawing.Size(180, 44);
            this.button_blocks.TabIndex = 7;
            this.button_blocks.Text = "Load Blocks in the combobox";
            this.button_blocks.UseVisualStyleBackColor = true;
            this.button_blocks.Click += new System.EventHandler(this.button_blocks_Click);
            // 
            // comboBox_blocks
            // 
            this.comboBox_blocks.BackColor = System.Drawing.Color.White;
            this.comboBox_blocks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_blocks.ForeColor = System.Drawing.Color.Black;
            this.comboBox_blocks.FormattingEnabled = true;
            this.comboBox_blocks.Location = new System.Drawing.Point(287, 85);
            this.comboBox_blocks.Name = "comboBox_blocks";
            this.comboBox_blocks.Size = new System.Drawing.Size(180, 23);
            this.comboBox_blocks.TabIndex = 2;
            this.comboBox_blocks.SelectedIndexChanged += new System.EventHandler(this.comboBox_blocks_SelectedIndexChanged);
            // 
            // comboBox_block_atr1
            // 
            this.comboBox_block_atr1.BackColor = System.Drawing.Color.White;
            this.comboBox_block_atr1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_block_atr1.ForeColor = System.Drawing.Color.Black;
            this.comboBox_block_atr1.FormattingEnabled = true;
            this.comboBox_block_atr1.Location = new System.Drawing.Point(287, 114);
            this.comboBox_block_atr1.Name = "comboBox_block_atr1";
            this.comboBox_block_atr1.Size = new System.Drawing.Size(180, 23);
            this.comboBox_block_atr1.TabIndex = 2;
            // 
            // button_change_attribute_value
            // 
            this.button_change_attribute_value.Location = new System.Drawing.Point(119, 193);
            this.button_change_attribute_value.Name = "button_change_attribute_value";
            this.button_change_attribute_value.Size = new System.Drawing.Size(265, 44);
            this.button_change_attribute_value.TabIndex = 7;
            this.button_change_attribute_value.Text = "Replace 1 with 2";
            this.button_change_attribute_value.UseVisualStyleBackColor = true;
            this.button_change_attribute_value.Click += new System.EventHandler(this.button_change_attribute_value_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(238, 88);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(39, 15);
            this.label8.TabIndex = 4;
            this.label8.Text = "Block";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(221, 117);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(56, 15);
            this.label9.TabIndex = 4;
            this.label9.Text = "Attribute";
            // 
            // button_load_OD
            // 
            this.button_load_OD.Location = new System.Drawing.Point(12, 143);
            this.button_load_OD.Name = "button_load_OD";
            this.button_load_OD.Size = new System.Drawing.Size(455, 44);
            this.button_load_OD.TabIndex = 7;
            this.button_load_OD.Text = "Load OD values";
            this.button_load_OD.UseVisualStyleBackColor = true;
            this.button_load_OD.Click += new System.EventHandler(this.button_load_OD_Click);
            // 
            // OD_layer_change_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(483, 251);
            this.Controls.Add(this.button_load_OD);
            this.Controls.Add(this.button_change_attribute_value);
            this.Controls.Add(this.button_blocks);
            this.Controls.Add(this.checkBox_use_null_values);
            this.Controls.Add(this.panel_Excel);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBox_Layers);
            this.Controls.Add(this.comboBox_OD2);
            this.Controls.Add(this.comboBox_block_atr1);
            this.Controls.Add(this.comboBox_blocks);
            this.Controls.Add(this.comboBox_OD1);
            this.Controls.Add(this.button_change_layer);
            this.Controls.Add(this.Button_read_OD);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "OD_layer_change_form";
            this.Text = "Object Data Form";
            this.Click += new System.EventHandler(this.Form_Click);
            this.panel_Excel.ResumeLayout(false);
            this.panel_Excel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Button Button_read_OD;
        private System.Windows.Forms.ComboBox comboBox_OD1;
        internal System.Windows.Forms.Button button_read_excel_column;
        private System.Windows.Forms.TextBox textBox_excel_column;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_start;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_end;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox_Layers;
        private System.Windows.Forms.Label label4;
        internal System.Windows.Forms.Button button_change_layer;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel_Excel;
        private System.Windows.Forms.CheckBox checkBox_use_null_values;
        private System.Windows.Forms.ComboBox comboBox_OD2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button_blocks;
        private System.Windows.Forms.ComboBox comboBox_blocks;
        private System.Windows.Forms.ComboBox comboBox_block_atr1;
        private System.Windows.Forms.Button button_change_attribute_value;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button button_load_OD;
    }
}