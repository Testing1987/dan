namespace Alignment_mdi
{
    partial class OD_attach
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
            this.panel_Excel = new System.Windows.Forms.Panel();
            this.button_read_excel_column = new System.Windows.Forms.Button();
            this.textBox_excel_column = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_start = new System.Windows.Forms.TextBox();
            this.textBox_end = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.button_ADD_TO_LAYERS = new System.Windows.Forms.Button();
            this.panel_Excel.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_Excel
            // 
            this.panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel_Excel.Controls.Add(this.button_read_excel_column);
            this.panel_Excel.Controls.Add(this.textBox_excel_column);
            this.panel_Excel.Controls.Add(this.label4);
            this.panel_Excel.Controls.Add(this.label1);
            this.panel_Excel.Controls.Add(this.label3);
            this.panel_Excel.Controls.Add(this.textBox_start);
            this.panel_Excel.Controls.Add(this.textBox_end);
            this.panel_Excel.Controls.Add(this.label2);
            this.panel_Excel.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.panel_Excel.Location = new System.Drawing.Point(12, 1);
            this.panel_Excel.Name = "panel_Excel";
            this.panel_Excel.Size = new System.Drawing.Size(204, 165);
            this.panel_Excel.TabIndex = 6;
            // 
            // button_read_excel_column
            // 
            this.button_read_excel_column.Location = new System.Drawing.Point(6, 110);
            this.button_read_excel_column.Name = "button_read_excel_column";
            this.button_read_excel_column.Size = new System.Drawing.Size(191, 44);
            this.button_read_excel_column.TabIndex = 1;
            this.button_read_excel_column.Text = "Read from Excel";
            this.button_read_excel_column.UseVisualStyleBackColor = true;
            this.button_read_excel_column.Click += new System.EventHandler(this.button_read_excel_column_Click);
            // 
            // textBox_excel_column
            // 
            this.textBox_excel_column.BackColor = System.Drawing.Color.White;
            this.textBox_excel_column.ForeColor = System.Drawing.Color.Black;
            this.textBox_excel_column.Location = new System.Drawing.Point(100, 28);
            this.textBox_excel_column.Name = "textBox_excel_column";
            this.textBox_excel_column.Size = new System.Drawing.Size(46, 21);
            this.textBox_excel_column.TabIndex = 3;
            this.textBox_excel_column.Text = "B";
            this.textBox_excel_column.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(6, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Excel Colum";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label3.Location = new System.Drawing.Point(6, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "End row";
            // 
            // textBox_start
            // 
            this.textBox_start.BackColor = System.Drawing.Color.White;
            this.textBox_start.ForeColor = System.Drawing.Color.Black;
            this.textBox_start.Location = new System.Drawing.Point(100, 55);
            this.textBox_start.Name = "textBox_start";
            this.textBox_start.Size = new System.Drawing.Size(46, 21);
            this.textBox_start.TabIndex = 3;
            this.textBox_start.Text = "7";
            this.textBox_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_end
            // 
            this.textBox_end.BackColor = System.Drawing.Color.White;
            this.textBox_end.ForeColor = System.Drawing.Color.Black;
            this.textBox_end.Location = new System.Drawing.Point(100, 83);
            this.textBox_end.Name = "textBox_end";
            this.textBox_end.Size = new System.Drawing.Size(46, 21);
            this.textBox_end.TabIndex = 3;
            this.textBox_end.Text = "181";
            this.textBox_end.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Location = new System.Drawing.Point(6, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Start row";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Location = new System.Drawing.Point(6, 6);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(153, 17);
            this.label4.TabIndex = 4;
            this.label4.Text = "Read layer list from Excel";
            // 
            // button_ADD_TO_LAYERS
            // 
            this.button_ADD_TO_LAYERS.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_ADD_TO_LAYERS.Location = new System.Drawing.Point(233, 12);
            this.button_ADD_TO_LAYERS.Name = "button_ADD_TO_LAYERS";
            this.button_ADD_TO_LAYERS.Size = new System.Drawing.Size(191, 44);
            this.button_ADD_TO_LAYERS.TabIndex = 1;
            this.button_ADD_TO_LAYERS.Text = "ADD OD TO LAYERS";
            this.button_ADD_TO_LAYERS.UseVisualStyleBackColor = true;
            this.button_ADD_TO_LAYERS.Click += new System.EventHandler(this.button_ADD_TO_LAYERS_Click);
            // 
            // OD_attach
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(499, 400);
            this.Controls.Add(this.button_ADD_TO_LAYERS);
            this.Controls.Add(this.panel_Excel);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "OD_attach";
            this.Text = "OD attach to layers";
            this.panel_Excel.ResumeLayout(false);
            this.panel_Excel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_Excel;
        internal System.Windows.Forms.Button button_read_excel_column;
        private System.Windows.Forms.TextBox textBox_excel_column;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_start;
        private System.Windows.Forms.TextBox textBox_end;
        private System.Windows.Forms.Label label2;
        internal System.Windows.Forms.Button button_ADD_TO_LAYERS;
    }
}