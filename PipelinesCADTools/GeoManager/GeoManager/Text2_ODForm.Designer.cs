namespace MMGeoTools
{
    partial class Text2_ODForm
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
            this.comboBox_OD_table = new System.Windows.Forms.ComboBox();
            this.comboBox_OD_field = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button_load_OD = new System.Windows.Forms.Button();
            this.button_tranfer_text = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBox_OD_table
            // 
            this.comboBox_OD_table.BackColor = System.Drawing.Color.White;
            this.comboBox_OD_table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_OD_table.ForeColor = System.Drawing.Color.Black;
            this.comboBox_OD_table.FormattingEnabled = true;
            this.comboBox_OD_table.Location = new System.Drawing.Point(76, 56);
            this.comboBox_OD_table.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.comboBox_OD_table.Name = "comboBox_OD_table";
            this.comboBox_OD_table.Size = new System.Drawing.Size(172, 24);
            this.comboBox_OD_table.TabIndex = 0;
            this.comboBox_OD_table.SelectedIndexChanged += new System.EventHandler(this.comboBox_OD_table_SelectedIndexChanged);
            // 
            // comboBox_OD_field
            // 
            this.comboBox_OD_field.BackColor = System.Drawing.Color.White;
            this.comboBox_OD_field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_OD_field.ForeColor = System.Drawing.Color.Black;
            this.comboBox_OD_field.FormattingEnabled = true;
            this.comboBox_OD_field.Location = new System.Drawing.Point(76, 88);
            this.comboBox_OD_field.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.comboBox_OD_field.Name = "comboBox_OD_field";
            this.comboBox_OD_field.Size = new System.Drawing.Size(172, 24);
            this.comboBox_OD_field.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "OD table";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(2, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "OD field";
            // 
            // button_load_OD
            // 
            this.button_load_OD.Location = new System.Drawing.Point(76, 12);
            this.button_load_OD.Name = "button_load_OD";
            this.button_load_OD.Size = new System.Drawing.Size(75, 37);
            this.button_load_OD.TabIndex = 2;
            this.button_load_OD.Text = "Load OD";
            this.button_load_OD.UseVisualStyleBackColor = true;
            this.button_load_OD.Click += new System.EventHandler(this.button_load_OD_Click);
            // 
            // button_tranfer_text
            // 
            this.button_tranfer_text.Location = new System.Drawing.Point(173, 119);
            this.button_tranfer_text.Name = "button_tranfer_text";
            this.button_tranfer_text.Size = new System.Drawing.Size(75, 37);
            this.button_tranfer_text.TabIndex = 2;
            this.button_tranfer_text.Text = "Text 2 OD";
            this.button_tranfer_text.UseVisualStyleBackColor = true;
            this.button_tranfer_text.Click += new System.EventHandler(this.button_tranfer_text_Click);
            // 
            // Text2_ODForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(258, 163);
            this.Controls.Add(this.button_tranfer_text);
            this.Controls.Add(this.button_load_OD);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox_OD_field);
            this.Controls.Add(this.comboBox_OD_table);
            this.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.Name = "Text2_ODForm";
            this.Text = "Txt2OD";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox_OD_table;
        private System.Windows.Forms.ComboBox comboBox_OD_field;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_load_OD;
        private System.Windows.Forms.Button button_tranfer_text;
    }
}