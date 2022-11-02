namespace Bump
{
    partial class bump_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(bump_form));
            this.panel_header = new System.Windows.Forms.Panel();
            this.button_minimize = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.label_mm = new System.Windows.Forms.Label();
            this.panel8 = new System.Windows.Forms.Panel();
            this.label17 = new System.Windows.Forms.Label();
            this.button_matl_database = new System.Windows.Forms.Button();
            this.panel10 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox_precision = new System.Windows.Forms.ComboBox();
            this.rb_Metric = new System.Windows.Forms.RadioButton();
            this.rb_Imperial = new System.Windows.Forms.RadioButton();
            this.tbox_Bump_Value = new System.Windows.Forms.TextBox();
            this.btn_draw = new System.Windows.Forms.Button();
            this.label29 = new System.Windows.Forms.Label();
            this.panel_header.SuspendLayout();
            this.panel8.SuspendLayout();
            this.panel10.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel_header.Controls.Add(this.button_minimize);
            this.panel_header.Controls.Add(this.button_Exit);
            this.panel_header.Controls.Add(this.label_mm);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(210, 39);
            this.panel_header.TabIndex = 2108;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_header.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_header.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // button_minimize
            // 
            this.button_minimize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_minimize.BackgroundImage")));
            this.button_minimize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_minimize.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.FlatAppearance.BorderSize = 0;
            this.button_minimize.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_minimize.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_minimize.ForeColor = System.Drawing.Color.White;
            this.button_minimize.Location = new System.Drawing.Point(146, 4);
            this.button_minimize.Name = "button_minimize";
            this.button_minimize.Size = new System.Drawing.Size(30, 30);
            this.button_minimize.TabIndex = 162;
            this.button_minimize.UseVisualStyleBackColor = false;
            this.button_minimize.Click += new System.EventHandler(this.button_minimize_Click);
            // 
            // button_Exit
            // 
            this.button_Exit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_Exit.BackgroundImage")));
            this.button_Exit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_Exit.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.FlatAppearance.BorderSize = 0;
            this.button_Exit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Exit.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.ForeColor = System.Drawing.Color.White;
            this.button_Exit.Location = new System.Drawing.Point(182, 4);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(30, 30);
            this.button_Exit.TabIndex = 161;
            this.button_Exit.UseVisualStyleBackColor = false;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // label_mm
            // 
            this.label_mm.AutoSize = true;
            this.label_mm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_mm.ForeColor = System.Drawing.Color.White;
            this.label_mm.Location = new System.Drawing.Point(3, 9);
            this.label_mm.Name = "label_mm";
            this.label_mm.Size = new System.Drawing.Size(137, 20);
            this.label_mm.TabIndex = 3;
            this.label_mm.Text = "Mott Macdonald";
            this.label_mm.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.label_mm.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.label_mm.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // panel8
            // 
            this.panel8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel8.Controls.Add(this.label17);
            this.panel8.Controls.Add(this.button_matl_database);
            this.panel8.Location = new System.Drawing.Point(7, 45);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(193, 25);
            this.panel8.TabIndex = 2145;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.BackColor = System.Drawing.Color.Transparent;
            this.label17.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label17.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label17.Location = new System.Drawing.Point(3, 3);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(50, 18);
            this.label17.TabIndex = 2054;
            this.label17.Text = "BUMP";
            // 
            // button_matl_database
            // 
            this.button_matl_database.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.button_matl_database.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_matl_database.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_matl_database.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_matl_database.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_matl_database.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_matl_database.ForeColor = System.Drawing.Color.White;
            this.button_matl_database.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_matl_database.Location = new System.Drawing.Point(1288, -1);
            this.button_matl_database.Name = "button_matl_database";
            this.button_matl_database.Size = new System.Drawing.Size(92, 25);
            this.button_matl_database.TabIndex = 2325;
            this.button_matl_database.Text = "Export";
            this.button_matl_database.UseVisualStyleBackColor = false;
            this.button_matl_database.Visible = false;
            // 
            // panel10
            // 
            this.panel10.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel10.Controls.Add(this.label6);
            this.panel10.Controls.Add(this.comboBox_precision);
            this.panel10.Controls.Add(this.rb_Metric);
            this.panel10.Controls.Add(this.rb_Imperial);
            this.panel10.Controls.Add(this.tbox_Bump_Value);
            this.panel10.Controls.Add(this.btn_draw);
            this.panel10.Controls.Add(this.label29);
            this.panel10.Location = new System.Drawing.Point(7, 69);
            this.panel10.Name = "panel10";
            this.panel10.Size = new System.Drawing.Size(193, 140);
            this.panel10.TabIndex = 2214;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(6, 59);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(50, 13);
            this.label6.TabIndex = 2331;
            this.label6.Text = "Precision";
            // 
            // comboBox_precision
            // 
            this.comboBox_precision.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_precision.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_precision.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_precision.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.comboBox_precision.ForeColor = System.Drawing.Color.White;
            this.comboBox_precision.FormattingEnabled = true;
            this.comboBox_precision.Location = new System.Drawing.Point(75, 56);
            this.comboBox_precision.Name = "comboBox_precision";
            this.comboBox_precision.Size = new System.Drawing.Size(113, 21);
            this.comboBox_precision.TabIndex = 2330;
            // 
            // rb_Metric
            // 
            this.rb_Metric.AutoSize = true;
            this.rb_Metric.ForeColor = System.Drawing.Color.White;
            this.rb_Metric.Location = new System.Drawing.Point(6, 32);
            this.rb_Metric.Margin = new System.Windows.Forms.Padding(2);
            this.rb_Metric.Name = "rb_Metric";
            this.rb_Metric.Size = new System.Drawing.Size(54, 17);
            this.rb_Metric.TabIndex = 2328;
            this.rb_Metric.Text = "Metric";
            this.rb_Metric.UseVisualStyleBackColor = true;
            // 
            // rb_Imperial
            // 
            this.rb_Imperial.AutoSize = true;
            this.rb_Imperial.Checked = true;
            this.rb_Imperial.ForeColor = System.Drawing.Color.White;
            this.rb_Imperial.Location = new System.Drawing.Point(6, 5);
            this.rb_Imperial.Margin = new System.Windows.Forms.Padding(2);
            this.rb_Imperial.Name = "rb_Imperial";
            this.rb_Imperial.Size = new System.Drawing.Size(61, 17);
            this.rb_Imperial.TabIndex = 2329;
            this.rb_Imperial.TabStop = true;
            this.rb_Imperial.Text = "Imperial";
            this.rb_Imperial.UseVisualStyleBackColor = true;
            // 
            // tbox_Bump_Value
            // 
            this.tbox_Bump_Value.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.tbox_Bump_Value.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbox_Bump_Value.ForeColor = System.Drawing.Color.White;
            this.tbox_Bump_Value.Location = new System.Drawing.Point(75, 82);
            this.tbox_Bump_Value.Name = "tbox_Bump_Value";
            this.tbox_Bump_Value.Size = new System.Drawing.Size(113, 20);
            this.tbox_Bump_Value.TabIndex = 2327;
            this.tbox_Bump_Value.Text = "0";
            this.tbox_Bump_Value.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tbox_Bump_Value.TextChanged += new System.EventHandler(this.tbox_Bump_Value_TextChanged_1);
            // 
            // btn_draw
            // 
            this.btn_draw.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.btn_draw.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.btn_draw.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btn_draw.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.btn_draw.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_draw.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.btn_draw.ForeColor = System.Drawing.Color.White;
            this.btn_draw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_draw.Location = new System.Drawing.Point(3, 108);
            this.btn_draw.Name = "btn_draw";
            this.btn_draw.Size = new System.Drawing.Size(185, 28);
            this.btn_draw.TabIndex = 2326;
            this.btn_draw.Text = "Execute";
            this.btn_draw.UseVisualStyleBackColor = false;
            this.btn_draw.Click += new System.EventHandler(this.btn_bump_Click);
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label29.ForeColor = System.Drawing.Color.White;
            this.label29.Location = new System.Drawing.Point(6, 86);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(64, 13);
            this.label29.TabIndex = 3;
            this.label29.Text = "Bump Value";
            // 
            // bump_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(210, 216);
            this.Controls.Add(this.panel10);
            this.Controls.Add(this.panel8);
            this.Controls.Add(this.panel_header);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "bump_form";
            this.Text = "Form1";
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            this.panel8.ResumeLayout(false);
            this.panel8.PerformLayout();
            this.panel10.ResumeLayout(false);
            this.panel10.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button button_matl_database;
        private System.Windows.Forms.Panel panel10;
        private System.Windows.Forms.TextBox tbox_Bump_Value;
        private System.Windows.Forms.Button btn_draw;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox_precision;
        private System.Windows.Forms.RadioButton rb_Metric;
        private System.Windows.Forms.RadioButton rb_Imperial;
    }
}