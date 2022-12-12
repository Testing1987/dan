
namespace Alignment_mdi
{
    partial class cs_form
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
            this.panel2 = new System.Windows.Forms.Panel();
            this.button_load_cs = new System.Windows.Forms.Button();
            this.button_convert = new System.Windows.Forms.Button();
            this.comboBox_to = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox_from = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox_xl = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label_brightness = new System.Windows.Forms.Label();
            this.label_contrast = new System.Windows.Forms.Label();
            this.textBox_start = new System.Windows.Forms.TextBox();
            this.textBox_X = new System.Windows.Forms.TextBox();
            this.textBox_Y = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.button_load_cs);
            this.panel2.Controls.Add(this.button_convert);
            this.panel2.Controls.Add(this.comboBox_to);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.comboBox_from);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.comboBox_xl);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label_brightness);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.label_contrast);
            this.panel2.Controls.Add(this.textBox_start);
            this.panel2.Controls.Add(this.textBox_X);
            this.panel2.Controls.Add(this.textBox_Y);
            this.panel2.ForeColor = System.Drawing.Color.White;
            this.panel2.Location = new System.Drawing.Point(3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(425, 232);
            this.panel2.TabIndex = 2207;
            // 
            // button_load_cs
            // 
            this.button_load_cs.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_load_cs.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_load_cs.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_cs.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_load_cs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_cs.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_load_cs.ForeColor = System.Drawing.Color.White;
            this.button_load_cs.Image = global::Alignment_mdi.Properties.Resources.Target;
            this.button_load_cs.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_load_cs.Location = new System.Drawing.Point(315, 47);
            this.button_load_cs.Name = "button_load_cs";
            this.button_load_cs.Size = new System.Drawing.Size(105, 28);
            this.button_load_cs.TabIndex = 2266;
            this.button_load_cs.Text = "Load CS";
            this.button_load_cs.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_load_cs.UseVisualStyleBackColor = false;
            this.button_load_cs.Click += new System.EventHandler(this.button_load_cs_Click);
            // 
            // button_convert
            // 
            this.button_convert.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_convert.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_convert.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_convert.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_convert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_convert.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_convert.ForeColor = System.Drawing.Color.White;
            this.button_convert.Image = global::Alignment_mdi.Properties.Resources.check;
            this.button_convert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_convert.Location = new System.Drawing.Point(235, 199);
            this.button_convert.Name = "button_convert";
            this.button_convert.Size = new System.Drawing.Size(185, 28);
            this.button_convert.TabIndex = 2259;
            this.button_convert.Text = "Convert Coordinates";
            this.button_convert.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_convert.UseVisualStyleBackColor = false;
            this.button_convert.Click += new System.EventHandler(this.button_convert_coordinates_Click);
            // 
            // comboBox_to
            // 
            this.comboBox_to.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_to.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_to.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_to.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.comboBox_to.ForeColor = System.Drawing.Color.White;
            this.comboBox_to.FormattingEnabled = true;
            this.comboBox_to.Location = new System.Drawing.Point(101, 63);
            this.comboBox_to.Name = "comboBox_to";
            this.comboBox_to.Size = new System.Drawing.Size(208, 24);
            this.comboBox_to.TabIndex = 2265;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(6, 175);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 14);
            this.label2.TabIndex = 2263;
            this.label2.Text = "Start Row";
            // 
            // comboBox_from
            // 
            this.comboBox_from.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_from.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_from.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_from.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.comboBox_from.ForeColor = System.Drawing.Color.White;
            this.comboBox_from.FormattingEnabled = true;
            this.comboBox_from.Location = new System.Drawing.Point(101, 33);
            this.comboBox_from.Name = "comboBox_from";
            this.comboBox_from.Size = new System.Drawing.Size(208, 24);
            this.comboBox_from.TabIndex = 2265;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(5, 68);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 14);
            this.label4.TabIndex = 2263;
            this.label4.Text = "Destination CS";
            // 
            // comboBox_xl
            // 
            this.comboBox_xl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_xl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_xl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_xl.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.comboBox_xl.ForeColor = System.Drawing.Color.White;
            this.comboBox_xl.FormattingEnabled = true;
            this.comboBox_xl.Location = new System.Drawing.Point(3, 3);
            this.comboBox_xl.Name = "comboBox_xl";
            this.comboBox_xl.Size = new System.Drawing.Size(417, 24);
            this.comboBox_xl.TabIndex = 2265;
            this.comboBox_xl.DropDown += new System.EventHandler(this.comboBox_xl_DropDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(5, 38);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 14);
            this.label3.TabIndex = 2263;
            this.label3.Text = "Source CS";
            // 
            // label_brightness
            // 
            this.label_brightness.AutoSize = true;
            this.label_brightness.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_brightness.ForeColor = System.Drawing.Color.White;
            this.label_brightness.Location = new System.Drawing.Point(11, 122);
            this.label_brightness.Name = "label_brightness";
            this.label_brightness.Size = new System.Drawing.Size(55, 14);
            this.label_brightness.TabIndex = 2263;
            this.label_brightness.Text = "North [Y]";
            // 
            // label_contrast
            // 
            this.label_contrast.AutoSize = true;
            this.label_contrast.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_contrast.ForeColor = System.Drawing.Color.White;
            this.label_contrast.Location = new System.Drawing.Point(89, 122);
            this.label_contrast.Name = "label_contrast";
            this.label_contrast.Size = new System.Drawing.Size(48, 14);
            this.label_contrast.TabIndex = 2264;
            this.label_contrast.Text = "East [X]";
            // 
            // textBox_start
            // 
            this.textBox_start.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_start.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_start.ForeColor = System.Drawing.Color.White;
            this.textBox_start.Location = new System.Drawing.Point(72, 173);
            this.textBox_start.Name = "textBox_start";
            this.textBox_start.Size = new System.Drawing.Size(45, 20);
            this.textBox_start.TabIndex = 2261;
            this.textBox_start.Text = "2";
            this.textBox_start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_X
            // 
            this.textBox_X.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_X.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_X.ForeColor = System.Drawing.Color.White;
            this.textBox_X.Location = new System.Drawing.Point(92, 139);
            this.textBox_X.Name = "textBox_X";
            this.textBox_X.Size = new System.Drawing.Size(45, 20);
            this.textBox_X.TabIndex = 2262;
            this.textBox_X.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Y
            // 
            this.textBox_Y.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_Y.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_Y.ForeColor = System.Drawing.Color.White;
            this.textBox_Y.Location = new System.Drawing.Point(9, 139);
            this.textBox_Y.Name = "textBox_Y";
            this.textBox_Y.Size = new System.Drawing.Size(45, 20);
            this.textBox_Y.TabIndex = 2261;
            this.textBox_Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.panel2);
            this.panel3.Location = new System.Drawing.Point(8, 11);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(859, 616);
            this.panel3.TabIndex = 2259;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(46, 104);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 14);
            this.label1.TabIndex = 2264;
            this.label1.Text = "Columns";
            // 
            // cs_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(877, 637);
            this.Controls.Add(this.panel3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "cs_form";
            this.Text = "AGEN_16_Toolz";
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button_convert;
        private System.Windows.Forms.Label label_brightness;
        private System.Windows.Forms.Label label_contrast;
        private System.Windows.Forms.TextBox textBox_X;
        private System.Windows.Forms.TextBox textBox_Y;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_start;
        private System.Windows.Forms.ComboBox comboBox_xl;
        private System.Windows.Forms.ComboBox comboBox_to;
        private System.Windows.Forms.ComboBox comboBox_from;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button_load_cs;
        private System.Windows.Forms.Label label1;
    }
}