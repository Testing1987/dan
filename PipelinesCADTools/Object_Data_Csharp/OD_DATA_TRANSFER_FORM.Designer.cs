namespace Alignment_mdi
{
    partial class OD_DATA_TRANSFER_FORM
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
            this.comboBox_Data_table = new System.Windows.Forms.ComboBox();
            this.comboBox_ATR1 = new System.Windows.Forms.ComboBox();
            this.comboBox_ATR2 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // comboBox_Data_table
            // 
            this.comboBox_Data_table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Data_table.FormattingEnabled = true;
            this.comboBox_Data_table.Location = new System.Drawing.Point(23, 35);
            this.comboBox_Data_table.Name = "comboBox_Data_table";
            this.comboBox_Data_table.Size = new System.Drawing.Size(379, 23);
            this.comboBox_Data_table.TabIndex = 0;
            // 
            // comboBox_ATR1
            // 
            this.comboBox_ATR1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_ATR1.FormattingEnabled = true;
            this.comboBox_ATR1.Location = new System.Drawing.Point(23, 79);
            this.comboBox_ATR1.Name = "comboBox_ATR1";
            this.comboBox_ATR1.Size = new System.Drawing.Size(269, 23);
            this.comboBox_ATR1.TabIndex = 0;
            // 
            // comboBox_ATR2
            // 
            this.comboBox_ATR2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_ATR2.FormattingEnabled = true;
            this.comboBox_ATR2.Location = new System.Drawing.Point(23, 121);
            this.comboBox_ATR2.Name = "comboBox_ATR2";
            this.comboBox_ATR2.Size = new System.Drawing.Size(269, 23);
            this.comboBox_ATR2.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(23, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Object Data Table";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "Object Data Table";
            // 
            // OD_DATA_TRANSFER_FORM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(734, 482);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox_ATR2);
            this.Controls.Add(this.comboBox_ATR1);
            this.Controls.Add(this.comboBox_Data_table);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "OD_DATA_TRANSFER_FORM";
            this.Text = "OD_DATA_TRANSFER_FORM";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox_Data_table;
        private System.Windows.Forms.ComboBox comboBox_ATR1;
        private System.Windows.Forms.ComboBox comboBox_ATR2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}