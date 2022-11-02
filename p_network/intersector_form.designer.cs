namespace p_network
{
    partial class form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form1));
            this.panel11 = new System.Windows.Forms.Panel();
            this.button_scan = new System.Windows.Forms.Button();
            this.panel_header = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.button_minimize = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.label_mm = new System.Windows.Forms.Label();
            this.panel17 = new System.Windows.Forms.Panel();
            this.radioButton_map = new System.Windows.Forms.RadioButton();
            this.radioButton_civil = new System.Windows.Forms.RadioButton();
            this.panel11.SuspendLayout();
            this.panel_header.SuspendLayout();
            this.panel17.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel11
            // 
            this.panel11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel11.Controls.Add(this.panel17);
            this.panel11.Controls.Add(this.button_scan);
            this.panel11.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel11.Location = new System.Drawing.Point(0, 39);
            this.panel11.Margin = new System.Windows.Forms.Padding(0);
            this.panel11.Name = "panel11";
            this.panel11.Size = new System.Drawing.Size(305, 63);
            this.panel11.TabIndex = 2022;
            // 
            // button_scan
            // 
            this.button_scan.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_scan.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_scan.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_scan.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_scan.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_scan.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_scan.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_scan.ForeColor = System.Drawing.Color.White;
            this.button_scan.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_scan.Location = new System.Drawing.Point(3, 5);
            this.button_scan.Name = "button_scan";
            this.button_scan.Size = new System.Drawing.Size(201, 50);
            this.button_scan.TabIndex = 2146;
            this.button_scan.Text = "Scan";
            this.button_scan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_scan.UseVisualStyleBackColor = false;
            this.button_scan.Click += new System.EventHandler(this.button_scan_Click);
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel_header.Controls.Add(this.label1);
            this.panel_header.Controls.Add(this.button_minimize);
            this.panel_header.Controls.Add(this.button_Exit);
            this.panel_header.Controls.Add(this.label_mm);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Padding = new System.Windows.Forms.Padding(4);
            this.panel_header.Size = new System.Drawing.Size(305, 39);
            this.panel_header.TabIndex = 2220;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_header.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_header.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(10, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(185, 16);
            this.label1.TabIndex = 164;
            this.label1.Text = "Pipe Networks Intersector";
            // 
            // button_minimize
            // 
            this.button_minimize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_minimize.BackgroundImage")));
            this.button_minimize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_minimize.Dock = System.Windows.Forms.DockStyle.Right;
            this.button_minimize.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.FlatAppearance.BorderSize = 0;
            this.button_minimize.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_minimize.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_minimize.ForeColor = System.Drawing.Color.White;
            this.button_minimize.Location = new System.Drawing.Point(239, 4);
            this.button_minimize.Name = "button_minimize";
            this.button_minimize.Size = new System.Drawing.Size(31, 31);
            this.button_minimize.TabIndex = 162;
            this.button_minimize.UseVisualStyleBackColor = false;
            this.button_minimize.Click += new System.EventHandler(this.button_minimize_Click);
            // 
            // button_Exit
            // 
            this.button_Exit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_Exit.BackgroundImage")));
            this.button_Exit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_Exit.Dock = System.Windows.Forms.DockStyle.Right;
            this.button_Exit.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.FlatAppearance.BorderSize = 0;
            this.button_Exit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Exit.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.ForeColor = System.Drawing.Color.White;
            this.button_Exit.Location = new System.Drawing.Point(270, 4);
            this.button_Exit.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Padding = new System.Windows.Forms.Padding(10, 0, 0, 0);
            this.button_Exit.Size = new System.Drawing.Size(31, 31);
            this.button_Exit.TabIndex = 161;
            this.button_Exit.UseVisualStyleBackColor = false;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // label_mm
            // 
            this.label_mm.AutoSize = true;
            this.label_mm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_mm.ForeColor = System.Drawing.Color.White;
            this.label_mm.Location = new System.Drawing.Point(7, 2);
            this.label_mm.Name = "label_mm";
            this.label_mm.Size = new System.Drawing.Size(137, 20);
            this.label_mm.TabIndex = 0;
            this.label_mm.Text = "Mott Macdonald";
            // 
            // panel17
            // 
            this.panel17.Controls.Add(this.radioButton_map);
            this.panel17.Controls.Add(this.radioButton_civil);
            this.panel17.Location = new System.Drawing.Point(210, 5);
            this.panel17.Name = "panel17";
            this.panel17.Size = new System.Drawing.Size(90, 50);
            this.panel17.TabIndex = 2147;
            this.panel17.Visible = false;
            // 
            // radioButton_map
            // 
            this.radioButton_map.AutoSize = true;
            this.radioButton_map.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.radioButton_map.ForeColor = System.Drawing.Color.White;
            this.radioButton_map.Location = new System.Drawing.Point(3, 27);
            this.radioButton_map.Name = "radioButton_map";
            this.radioButton_map.Size = new System.Drawing.Size(64, 18);
            this.radioButton_map.TabIndex = 115;
            this.radioButton_map.Text = "Map 3D";
            this.radioButton_map.UseVisualStyleBackColor = true;
            // 
            // radioButton_civil
            // 
            this.radioButton_civil.AutoSize = true;
            this.radioButton_civil.Checked = true;
            this.radioButton_civil.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.radioButton_civil.ForeColor = System.Drawing.Color.White;
            this.radioButton_civil.Location = new System.Drawing.Point(3, 3);
            this.radioButton_civil.Name = "radioButton_civil";
            this.radioButton_civil.Size = new System.Drawing.Size(48, 18);
            this.radioButton_civil.TabIndex = 114;
            this.radioButton_civil.TabStop = true;
            this.radioButton_civil.Text = "Civil";
            this.radioButton_civil.UseVisualStyleBackColor = true;
            // 
            // form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.ClientSize = new System.Drawing.Size(305, 102);
            this.Controls.Add(this.panel11);
            this.Controls.Add(this.panel_header);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.IsMdiContainer = true;
            this.MaximizeBox = false;
            this.Name = "form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Design Suite";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel11.ResumeLayout(false);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            this.panel17.ResumeLayout(false);
            this.panel17.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel11;
        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.Button button_scan;
        private System.Windows.Forms.Panel panel17;
        private System.Windows.Forms.RadioButton radioButton_map;
        private System.Windows.Forms.RadioButton radioButton_civil;
    }
}
