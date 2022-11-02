namespace Alignment_mdi
{
    partial class AGEN_MaterialBand
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AGEN_MaterialBand));
            this.panel_material = new System.Windows.Forms.Panel();
            this.panel_dan = new System.Windows.Forms.Panel();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox_spacing = new System.Windows.Forms.TextBox();
            this.panel9 = new System.Windows.Forms.Panel();
            this.comboBox_segment_name = new System.Windows.Forms.ComboBox();
            this.panel11 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.button_show_mat_counts = new System.Windows.Forms.Button();
            this.dataGridView_materials = new System.Windows.Forms.DataGridView();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label_mat = new System.Windows.Forms.Label();
            this.button_generate_mat_spreadsheets_with_headers_only = new System.Windows.Forms.Button();
            this.button_open_materials = new System.Windows.Forms.Button();
            this.button_load_materials = new System.Windows.Forms.Button();
            this.label_mat_band = new System.Windows.Forms.Label();
            this.button_draw_mat_band = new System.Windows.Forms.Button();
            this.button_scan_heavy_wall = new System.Windows.Forms.Button();
            this.panel_material.SuspendLayout();
            this.panel_dan.SuspendLayout();
            this.panel9.SuspendLayout();
            this.panel11.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_materials)).BeginInit();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_material
            // 
            this.panel_material.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_material.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_material.Controls.Add(this.button_scan_heavy_wall);
            this.panel_material.Controls.Add(this.panel_dan);
            this.panel_material.Controls.Add(this.panel9);
            this.panel_material.Controls.Add(this.panel11);
            this.panel_material.Controls.Add(this.button_show_mat_counts);
            this.panel_material.Controls.Add(this.dataGridView_materials);
            this.panel_material.Controls.Add(this.panel3);
            this.panel_material.Controls.Add(this.button_generate_mat_spreadsheets_with_headers_only);
            this.panel_material.Controls.Add(this.button_open_materials);
            this.panel_material.Controls.Add(this.button_load_materials);
            this.panel_material.Controls.Add(this.label_mat_band);
            this.panel_material.Controls.Add(this.button_draw_mat_band);
            this.panel_material.Location = new System.Drawing.Point(12, 12);
            this.panel_material.Name = "panel_material";
            this.panel_material.Size = new System.Drawing.Size(855, 613);
            this.panel_material.TabIndex = 0;
            // 
            // panel_dan
            // 
            this.panel_dan.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_dan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_dan.Controls.Add(this.label7);
            this.panel_dan.Controls.Add(this.textBox_spacing);
            this.panel_dan.Location = new System.Drawing.Point(507, 507);
            this.panel_dan.Name = "panel_dan";
            this.panel_dan.Size = new System.Drawing.Size(142, 28);
            this.panel_dan.TabIndex = 2257;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(61, 5);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 14);
            this.label7.TabIndex = 2119;
            this.label7.Text = "Min Stretch";
            // 
            // textBox_spacing
            // 
            this.textBox_spacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_spacing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_spacing.ForeColor = System.Drawing.Color.White;
            this.textBox_spacing.Location = new System.Drawing.Point(4, 3);
            this.textBox_spacing.Name = "textBox_spacing";
            this.textBox_spacing.Size = new System.Drawing.Size(45, 20);
            this.textBox_spacing.TabIndex = 2117;
            this.textBox_spacing.Text = "0.5";
            this.textBox_spacing.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // panel9
            // 
            this.panel9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel9.Controls.Add(this.comboBox_segment_name);
            this.panel9.Location = new System.Drawing.Point(1, 28);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(186, 34);
            this.panel9.TabIndex = 2255;
            // 
            // comboBox_segment_name
            // 
            this.comboBox_segment_name.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_segment_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_segment_name.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_segment_name.ForeColor = System.Drawing.Color.White;
            this.comboBox_segment_name.FormattingEnabled = true;
            this.comboBox_segment_name.Location = new System.Drawing.Point(3, 5);
            this.comboBox_segment_name.Name = "comboBox_segment_name";
            this.comboBox_segment_name.Size = new System.Drawing.Size(175, 21);
            this.comboBox_segment_name.TabIndex = 2143;
            this.comboBox_segment_name.SelectedIndexChanged += new System.EventHandler(this.ComboBox_segment_name_SelectedIndexChanged);
            // 
            // panel11
            // 
            this.panel11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel11.Controls.Add(this.label2);
            this.panel11.Location = new System.Drawing.Point(1, 3);
            this.panel11.Name = "panel11";
            this.panel11.Size = new System.Drawing.Size(186, 25);
            this.panel11.TabIndex = 2256;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label2.Location = new System.Drawing.Point(3, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(118, 18);
            this.label2.TabIndex = 2054;
            this.label2.Text = "Segment Name";
            // 
            // button_show_mat_counts
            // 
            this.button_show_mat_counts.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_show_mat_counts.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_show_mat_counts.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_show_mat_counts.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_show_mat_counts.ForeColor = System.Drawing.Color.White;
            this.button_show_mat_counts.Image = ((System.Drawing.Image)(resources.GetObject("button_show_mat_counts.Image")));
            this.button_show_mat_counts.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_show_mat_counts.Location = new System.Drawing.Point(655, 578);
            this.button_show_mat_counts.Name = "button_show_mat_counts";
            this.button_show_mat_counts.Size = new System.Drawing.Size(195, 28);
            this.button_show_mat_counts.TabIndex = 2254;
            this.button_show_mat_counts.Text = "Continue to Material Counts";
            this.button_show_mat_counts.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_show_mat_counts.UseVisualStyleBackColor = false;
            this.button_show_mat_counts.Click += new System.EventHandler(this.button_show_mat_counts_Click);
            // 
            // dataGridView_materials
            // 
            this.dataGridView_materials.AllowUserToAddRows = false;
            this.dataGridView_materials.AllowUserToDeleteRows = false;
            this.dataGridView_materials.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView_materials.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.dataGridView_materials.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView_materials.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.dataGridView_materials.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.dataGridView_materials.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_materials.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_materials.GridColor = System.Drawing.Color.LightGray;
            this.dataGridView_materials.Location = new System.Drawing.Point(3, 144);
            this.dataGridView_materials.Name = "dataGridView_materials";
            this.dataGridView_materials.RowHeadersVisible = false;
            this.dataGridView_materials.Size = new System.Drawing.Size(847, 357);
            this.dataGridView_materials.TabIndex = 2252;
            this.dataGridView_materials.TabStop = false;
            this.dataGridView_materials.VirtualMode = true;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label_mat);
            this.panel3.Location = new System.Drawing.Point(3, 119);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(847, 25);
            this.panel3.TabIndex = 2253;
            // 
            // label_mat
            // 
            this.label_mat.AutoSize = true;
            this.label_mat.BackColor = System.Drawing.Color.Transparent;
            this.label_mat.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_mat.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label_mat.Location = new System.Drawing.Point(3, 3);
            this.label_mat.Name = "label_mat";
            this.label_mat.Size = new System.Drawing.Size(99, 18);
            this.label_mat.TabIndex = 125;
            this.label_mat.Text = "Material Info";
            // 
            // button_generate_mat_spreadsheets_with_headers_only
            // 
            this.button_generate_mat_spreadsheets_with_headers_only.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_generate_mat_spreadsheets_with_headers_only.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_generate_mat_spreadsheets_with_headers_only.BackgroundImage")));
            this.button_generate_mat_spreadsheets_with_headers_only.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_generate_mat_spreadsheets_with_headers_only.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_generate_mat_spreadsheets_with_headers_only.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_generate_mat_spreadsheets_with_headers_only.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_generate_mat_spreadsheets_with_headers_only.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_generate_mat_spreadsheets_with_headers_only.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_generate_mat_spreadsheets_with_headers_only.ForeColor = System.Drawing.Color.White;
            this.button_generate_mat_spreadsheets_with_headers_only.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_generate_mat_spreadsheets_with_headers_only.Location = new System.Drawing.Point(6, 85);
            this.button_generate_mat_spreadsheets_with_headers_only.Name = "button_generate_mat_spreadsheets_with_headers_only";
            this.button_generate_mat_spreadsheets_with_headers_only.Size = new System.Drawing.Size(196, 28);
            this.button_generate_mat_spreadsheets_with_headers_only.TabIndex = 2250;
            this.button_generate_mat_spreadsheets_with_headers_only.Text = "New Material Spreadsheet";
            this.button_generate_mat_spreadsheets_with_headers_only.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_generate_mat_spreadsheets_with_headers_only.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_generate_mat_spreadsheets_with_headers_only.UseVisualStyleBackColor = false;
            this.button_generate_mat_spreadsheets_with_headers_only.Click += new System.EventHandler(this.Button_generate_mat_spreadsheets_with_headers_only_Click);
            // 
            // button_open_materials
            // 
            this.button_open_materials.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_open_materials.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_open_materials.BackgroundImage")));
            this.button_open_materials.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_open_materials.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_open_materials.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_open_materials.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_open_materials.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_open_materials.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_open_materials.ForeColor = System.Drawing.Color.White;
            this.button_open_materials.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_open_materials.Location = new System.Drawing.Point(6, 510);
            this.button_open_materials.Name = "button_open_materials";
            this.button_open_materials.Size = new System.Drawing.Size(196, 28);
            this.button_open_materials.TabIndex = 2250;
            this.button_open_materials.Text = "Open Material Spreadsheet";
            this.button_open_materials.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_open_materials.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_open_materials.UseVisualStyleBackColor = false;
            this.button_open_materials.Click += new System.EventHandler(this.button_open_mat_linear_Click);
            // 
            // button_load_materials
            // 
            this.button_load_materials.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_load_materials.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_load_materials.BackgroundImage")));
            this.button_load_materials.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_load_materials.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_load_materials.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_materials.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_load_materials.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_materials.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_load_materials.ForeColor = System.Drawing.Color.White;
            this.button_load_materials.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_load_materials.Location = new System.Drawing.Point(208, 85);
            this.button_load_materials.Name = "button_load_materials";
            this.button_load_materials.Size = new System.Drawing.Size(226, 28);
            this.button_load_materials.TabIndex = 0;
            this.button_load_materials.Text = "Load Material Spreadsheets";
            this.button_load_materials.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_load_materials.UseVisualStyleBackColor = false;
            this.button_load_materials.Click += new System.EventHandler(this.button_load_materials_Click);
            // 
            // label_mat_band
            // 
            this.label_mat_band.AutoSize = true;
            this.label_mat_band.BackColor = System.Drawing.Color.Transparent;
            this.label_mat_band.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_mat_band.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label_mat_band.Location = new System.Drawing.Point(3, 64);
            this.label_mat_band.Name = "label_mat_band";
            this.label_mat_band.Size = new System.Drawing.Size(108, 18);
            this.label_mat_band.TabIndex = 2070;
            this.label_mat_band.Text = "Material Band";
            // 
            // button_draw_mat_band
            // 
            this.button_draw_mat_band.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_draw_mat_band.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_draw_mat_band.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_draw_mat_band.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_draw_mat_band.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_draw_mat_band.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_draw_mat_band.ForeColor = System.Drawing.Color.White;
            this.button_draw_mat_band.Image = ((System.Drawing.Image)(resources.GetObject("button_draw_mat_band.Image")));
            this.button_draw_mat_band.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_draw_mat_band.Location = new System.Drawing.Point(655, 507);
            this.button_draw_mat_band.Name = "button_draw_mat_band";
            this.button_draw_mat_band.Size = new System.Drawing.Size(195, 28);
            this.button_draw_mat_band.TabIndex = 40;
            this.button_draw_mat_band.Text = "Draw Material Band";
            this.button_draw_mat_band.UseVisualStyleBackColor = false;
            this.button_draw_mat_band.Click += new System.EventHandler(this.button_draw_mat_band_Click);
            // 
            // button_scan_heavy_wall
            // 
            this.button_scan_heavy_wall.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_scan_heavy_wall.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_scan_heavy_wall.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_scan_heavy_wall.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_scan_heavy_wall.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_scan_heavy_wall.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_scan_heavy_wall.ForeColor = System.Drawing.Color.White;
            this.button_scan_heavy_wall.Image = ((System.Drawing.Image)(resources.GetObject("button_scan_heavy_wall.Image")));
            this.button_scan_heavy_wall.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_scan_heavy_wall.Location = new System.Drawing.Point(3, 578);
            this.button_scan_heavy_wall.Name = "button_scan_heavy_wall";
            this.button_scan_heavy_wall.Size = new System.Drawing.Size(199, 28);
            this.button_scan_heavy_wall.TabIndex = 2258;
            this.button_scan_heavy_wall.Text = "Scan Materials";
            this.button_scan_heavy_wall.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_scan_heavy_wall.UseVisualStyleBackColor = false;
            this.button_scan_heavy_wall.Click += new System.EventHandler(this.button_scan_heavy_wall_Click);
            // 
            // AGEN_MaterialBand
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(877, 637);
            this.Controls.Add(this.panel_material);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AGEN_MaterialBand";
            this.Text = "AGEN_0402_MaterialScanAndDraw";
            this.panel_material.ResumeLayout(false);
            this.panel_material.PerformLayout();
            this.panel_dan.ResumeLayout(false);
            this.panel_dan.PerformLayout();
            this.panel9.ResumeLayout(false);
            this.panel11.ResumeLayout(false);
            this.panel11.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_materials)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_material;
        private System.Windows.Forms.Label label_mat_band;
        private System.Windows.Forms.Button button_draw_mat_band;
        private System.Windows.Forms.Button button_load_materials;
        private System.Windows.Forms.Button button_open_materials;
        private System.Windows.Forms.DataGridView dataGridView_materials;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label_mat;
        private System.Windows.Forms.Button button_show_mat_counts;
        private System.Windows.Forms.Button button_generate_mat_spreadsheets_with_headers_only;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.ComboBox comboBox_segment_name;
        private System.Windows.Forms.Panel panel11;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel_dan;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox_spacing;
        private System.Windows.Forms.Button button_scan_heavy_wall;
    }
}