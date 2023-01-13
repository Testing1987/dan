namespace Alignment_mdi
{
    partial class AGEN_segments_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AGEN_segments_form));
            this.panel_segments = new System.Windows.Forms.Panel();
            this.button_remove_segment = new System.Windows.Forms.Button();
            this.button_add_segment = new System.Windows.Forms.Button();
            this.textBox_name1 = new System.Windows.Forms.TextBox();
            this.label_name1 = new System.Windows.Forms.Label();
            this.button_create = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button_minimize = new System.Windows.Forms.Button();
            this.button_close = new System.Windows.Forms.Button();
            this.label_header2 = new System.Windows.Forms.Label();
            this.label_header1 = new System.Windows.Forms.Label();
            this.panel_button = new System.Windows.Forms.Panel();
            this.panel_segments.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel_button.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_segments
            // 
            this.panel_segments.AutoScroll = true;
            this.panel_segments.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_segments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_segments.Controls.Add(this.textBox_name1);
            this.panel_segments.Controls.Add(this.label_name1);
            this.panel_segments.Location = new System.Drawing.Point(4, 43);
            this.panel_segments.Name = "panel_segments";
            this.panel_segments.Size = new System.Drawing.Size(397, 35);
            this.panel_segments.TabIndex = 2056;
            // 
            // button_remove_segment
            // 
            this.button_remove_segment.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_remove_segment.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_remove_segment.BackgroundImage")));
            this.button_remove_segment.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_remove_segment.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_remove_segment.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_remove_segment.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_remove_segment.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_remove_segment.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_remove_segment.ForeColor = System.Drawing.Color.White;
            this.button_remove_segment.Location = new System.Drawing.Point(30, 6);
            this.button_remove_segment.Name = "button_remove_segment";
            this.button_remove_segment.Size = new System.Drawing.Size(21, 21);
            this.button_remove_segment.TabIndex = 2067;
            this.button_remove_segment.TabStop = false;
            this.button_remove_segment.UseVisualStyleBackColor = false;
            this.button_remove_segment.Click += new System.EventHandler(this.button_remove_Click);
            // 
            // button_add_segment
            // 
            this.button_add_segment.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_add_segment.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_add_segment.BackgroundImage")));
            this.button_add_segment.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_add_segment.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_add_segment.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_add_segment.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_add_segment.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_add_segment.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_add_segment.ForeColor = System.Drawing.Color.White;
            this.button_add_segment.Location = new System.Drawing.Point(3, 6);
            this.button_add_segment.Name = "button_add_segment";
            this.button_add_segment.Size = new System.Drawing.Size(21, 21);
            this.button_add_segment.TabIndex = 2066;
            this.button_add_segment.TabStop = false;
            this.button_add_segment.UseVisualStyleBackColor = false;
            this.button_add_segment.Click += new System.EventHandler(this.button_add_custom_control_Click);
            this.button_add_segment.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.button_add_segment.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.button_add_segment.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // textBox_name1
            // 
            this.textBox_name1.Location = new System.Drawing.Point(155, 6);
            this.textBox_name1.Name = "textBox_name1";
            this.textBox_name1.Size = new System.Drawing.Size(214, 20);
            this.textBox_name1.TabIndex = 101;
            // 
            // label_name1
            // 
            this.label_name1.AutoSize = true;
            this.label_name1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_name1.ForeColor = System.Drawing.Color.White;
            this.label_name1.Location = new System.Drawing.Point(3, 9);
            this.label_name1.Name = "label_name1";
            this.label_name1.Size = new System.Drawing.Size(100, 14);
            this.label_name1.TabIndex = 94;
            this.label_name1.Text = "Segment 1 Name";
            // 
            // button_create
            // 
            this.button_create.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_create.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_create.BackgroundImage")));
            this.button_create.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_create.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_create.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_create.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_create.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_create.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_create.ForeColor = System.Drawing.Color.White;
            this.button_create.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_create.Location = new System.Drawing.Point(57, 5);
            this.button_create.Name = "button_create";
            this.button_create.Size = new System.Drawing.Size(101, 25);
            this.button_create.TabIndex = 93;
            this.button_create.Text = "Create";
            this.button_create.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_create.UseVisualStyleBackColor = false;
            this.button_create.Click += new System.EventHandler(this.button_create_segments_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel1.Controls.Add(this.button_minimize);
            this.panel1.Controls.Add(this.button_close);
            this.panel1.Controls.Add(this.label_header2);
            this.panel1.Controls.Add(this.label_header1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(570, 39);
            this.panel1.TabIndex = 2058;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
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
            this.button_minimize.Location = new System.Drawing.Point(501, 11);
            this.button_minimize.Name = "button_minimize";
            this.button_minimize.Size = new System.Drawing.Size(27, 20);
            this.button_minimize.TabIndex = 166;
            this.button_minimize.UseVisualStyleBackColor = false;
            this.button_minimize.Click += new System.EventHandler(this.button_minimize_Click);
            // 
            // button_close
            // 
            this.button_close.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_close.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_close.BackgroundImage")));
            this.button_close.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_close.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_close.FlatAppearance.BorderSize = 0;
            this.button_close.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_close.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_close.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_close.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_close.ForeColor = System.Drawing.Color.White;
            this.button_close.Location = new System.Drawing.Point(534, 6);
            this.button_close.Name = "button_close";
            this.button_close.Size = new System.Drawing.Size(30, 30);
            this.button_close.TabIndex = 165;
            this.button_close.UseVisualStyleBackColor = false;
            this.button_close.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // label_header2
            // 
            this.label_header2.AutoSize = true;
            this.label_header2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold);
            this.label_header2.ForeColor = System.Drawing.Color.White;
            this.label_header2.Location = new System.Drawing.Point(6, 22);
            this.label_header2.Name = "label_header2";
            this.label_header2.Size = new System.Drawing.Size(44, 16);
            this.label_header2.TabIndex = 164;
            this.label_header2.Text = "AGEN";
            // 
            // label_header1
            // 
            this.label_header1.AutoSize = true;
            this.label_header1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            this.label_header1.ForeColor = System.Drawing.Color.White;
            this.label_header1.Location = new System.Drawing.Point(3, 1);
            this.label_header1.Name = "label_header1";
            this.label_header1.Size = new System.Drawing.Size(150, 19);
            this.label_header1.TabIndex = 3;
            this.label_header1.Text = "Append Segments";
            // 
            // panel_button
            // 
            this.panel_button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_button.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_button.Controls.Add(this.button_remove_segment);
            this.panel_button.Controls.Add(this.button_create);
            this.panel_button.Controls.Add(this.button_add_segment);
            this.panel_button.Location = new System.Drawing.Point(402, 43);
            this.panel_button.Name = "panel_button";
            this.panel_button.Size = new System.Drawing.Size(165, 35);
            this.panel_button.TabIndex = 2056;
            // 
            // AGEN_segments_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(570, 82);
            this.Controls.Add(this.panel_button);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel_segments);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AGEN_segments_form";
            this.Text = "AGEN_0yyy_Inquiry_Tool";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel_segments.ResumeLayout(false);
            this.panel_segments.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel_button.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_segments;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label_header2;
        private System.Windows.Forms.Label label_header1;
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_close;
        private System.Windows.Forms.TextBox textBox_name1;
        private System.Windows.Forms.Label label_name1;
        private System.Windows.Forms.Button button_create;
        private System.Windows.Forms.Button button_remove_segment;
        private System.Windows.Forms.Button button_add_segment;
        private System.Windows.Forms.Panel panel_button;
    }
}