namespace Alignment_mdi
{
    partial class AGEN_MaterialCount
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AGEN_MaterialCount));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel_material = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.label_open_files_canada = new System.Windows.Forms.Label();
            this.comboBox_excel_files = new System.Windows.Forms.ComboBox();
            this.button_counts2TBLK = new System.Windows.Forms.Button();
            this.button_show_mat_draw = new System.Windows.Forms.Button();
            this.dataGridView_mat_totals = new System.Windows.Forms.DataGridView();
            this.button_calculate_totals = new System.Windows.Forms.Button();
            this.panel7 = new System.Windows.Forms.Panel();
            this.label12 = new System.Windows.Forms.Label();
            this.panel_material.SuspendLayout();
            this.panel6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_mat_totals)).BeginInit();
            this.panel7.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_material
            // 
            this.panel_material.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_material.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_material.Controls.Add(this.panel6);
            this.panel_material.Controls.Add(this.panel7);
            this.panel_material.Location = new System.Drawing.Point(12, 12);
            this.panel_material.Name = "panel_material";
            this.panel_material.Size = new System.Drawing.Size(855, 613);
            this.panel_material.TabIndex = 0;
            // 
            // panel6
            // 
            this.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel6.Controls.Add(this.label_open_files_canada);
            this.panel6.Controls.Add(this.comboBox_excel_files);
            this.panel6.Controls.Add(this.button_counts2TBLK);
            this.panel6.Controls.Add(this.button_show_mat_draw);
            this.panel6.Controls.Add(this.dataGridView_mat_totals);
            this.panel6.Controls.Add(this.button_calculate_totals);
            this.panel6.Location = new System.Drawing.Point(3, 28);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(843, 581);
            this.panel6.TabIndex = 2250;
            this.panel6.Click += new System.EventHandler(this.panel7_Click);
            // 
            // label_open_files_canada
            // 
            this.label_open_files_canada.AutoSize = true;
            this.label_open_files_canada.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_open_files_canada.ForeColor = System.Drawing.Color.Yellow;
            this.label_open_files_canada.Location = new System.Drawing.Point(374, 538);
            this.label_open_files_canada.Name = "label_open_files_canada";
            this.label_open_files_canada.Size = new System.Drawing.Size(142, 14);
            this.label_open_files_canada.TabIndex = 2483;
            this.label_open_files_canada.Text = "**Select From Open Files";
            this.label_open_files_canada.Visible = false;
            // 
            // comboBox_excel_files
            // 
            this.comboBox_excel_files.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_excel_files.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_excel_files.DropDownWidth = 179;
            this.comboBox_excel_files.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_excel_files.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.comboBox_excel_files.ForeColor = System.Drawing.Color.White;
            this.comboBox_excel_files.FormattingEnabled = true;
            this.comboBox_excel_files.Location = new System.Drawing.Point(374, 552);
            this.comboBox_excel_files.Margin = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.comboBox_excel_files.Name = "comboBox_excel_files";
            this.comboBox_excel_files.Size = new System.Drawing.Size(192, 22);
            this.comboBox_excel_files.TabIndex = 2482;
            this.comboBox_excel_files.Visible = false;
            this.comboBox_excel_files.DropDown += new System.EventHandler(this.comboBox_excel_files_DropDown);
            // 
            // button_counts2TBLK
            // 
            this.button_counts2TBLK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_counts2TBLK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_counts2TBLK.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_counts2TBLK.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_counts2TBLK.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_counts2TBLK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_counts2TBLK.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_counts2TBLK.ForeColor = System.Drawing.Color.White;
            this.button_counts2TBLK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_counts2TBLK.Location = new System.Drawing.Point(572, 548);
            this.button_counts2TBLK.Name = "button_counts2TBLK";
            this.button_counts2TBLK.Size = new System.Drawing.Size(266, 28);
            this.button_counts2TBLK.TabIndex = 2480;
            this.button_counts2TBLK.Text = "Tansfer Counts to TBLK format";
            this.button_counts2TBLK.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_counts2TBLK.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_counts2TBLK.UseVisualStyleBackColor = false;
            this.button_counts2TBLK.Visible = false;
            this.button_counts2TBLK.Click += new System.EventHandler(this.button_counts2TBLK_Click);
            // 
            // button_show_mat_draw
            // 
            this.button_show_mat_draw.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_show_mat_draw.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_show_mat_draw.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_show_mat_draw.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_show_mat_draw.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_show_mat_draw.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_show_mat_draw.ForeColor = System.Drawing.Color.White;
            this.button_show_mat_draw.Image = ((System.Drawing.Image)(resources.GetObject("button_show_mat_draw.Image")));
            this.button_show_mat_draw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_show_mat_draw.Location = new System.Drawing.Point(6, 548);
            this.button_show_mat_draw.Name = "button_show_mat_draw";
            this.button_show_mat_draw.Size = new System.Drawing.Size(170, 28);
            this.button_show_mat_draw.TabIndex = 2142;
            this.button_show_mat_draw.Text = "Back to material draw";
            this.button_show_mat_draw.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_show_mat_draw.UseVisualStyleBackColor = false;
            this.button_show_mat_draw.Click += new System.EventHandler(this.button_show_mat_draw_Click);
            // 
            // dataGridView_mat_totals
            // 
            this.dataGridView_mat_totals.AllowUserToAddRows = false;
            this.dataGridView_mat_totals.AllowUserToDeleteRows = false;
            this.dataGridView_mat_totals.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView_mat_totals.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.dataGridView_mat_totals.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView_mat_totals.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.dataGridView_mat_totals.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.dataGridView_mat_totals.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_mat_totals.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView_mat_totals.GridColor = System.Drawing.Color.LightGray;
            this.dataGridView_mat_totals.Location = new System.Drawing.Point(5, -1);
            this.dataGridView_mat_totals.Name = "dataGridView_mat_totals";
            this.dataGridView_mat_totals.RowHeadersVisible = false;
            this.dataGridView_mat_totals.Size = new System.Drawing.Size(843, 509);
            this.dataGridView_mat_totals.TabIndex = 2093;
            this.dataGridView_mat_totals.TabStop = false;
            this.dataGridView_mat_totals.VirtualMode = true;
            // 
            // button_calculate_totals
            // 
            this.button_calculate_totals.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_calculate_totals.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_calculate_totals.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_calculate_totals.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_calculate_totals.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_calculate_totals.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_calculate_totals.ForeColor = System.Drawing.Color.White;
            this.button_calculate_totals.Image = ((System.Drawing.Image)(resources.GetObject("button_calculate_totals.Image")));
            this.button_calculate_totals.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_calculate_totals.Location = new System.Drawing.Point(572, 514);
            this.button_calculate_totals.Name = "button_calculate_totals";
            this.button_calculate_totals.Size = new System.Drawing.Size(266, 28);
            this.button_calculate_totals.TabIndex = 40;
            this.button_calculate_totals.Text = "Calculate Totals";
            this.button_calculate_totals.UseVisualStyleBackColor = false;
            this.button_calculate_totals.Click += new System.EventHandler(this.button_calculate_totals_Click);
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.label12);
            this.panel7.Location = new System.Drawing.Point(3, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(843, 34);
            this.panel7.TabIndex = 2249;
            this.panel7.Click += new System.EventHandler(this.panel7_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.BackColor = System.Drawing.Color.Transparent;
            this.label12.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label12.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label12.Location = new System.Drawing.Point(3, 3);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(115, 18);
            this.label12.TabIndex = 2054;
            this.label12.Text = "Material Totals";
            this.label12.Click += new System.EventHandler(this.panel7_Click);
            // 
            // AGEN_MaterialCount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(877, 637);
            this.Controls.Add(this.panel_material);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AGEN_MaterialCount";
            this.Text = "AGEN_0402_MaterialScanAndDraw";
            this.panel_material.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_mat_totals)).EndInit();
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_material;
        private System.Windows.Forms.Button button_calculate_totals;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dataGridView_mat_totals;
        private System.Windows.Forms.Button button_show_mat_draw;
        private System.Windows.Forms.Button button_counts2TBLK;
        private System.Windows.Forms.Label label_open_files_canada;
        private System.Windows.Forms.ComboBox comboBox_excel_files;
    }
}