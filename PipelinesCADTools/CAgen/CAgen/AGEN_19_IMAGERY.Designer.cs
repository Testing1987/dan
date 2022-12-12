
namespace Alignment_mdi
{
    partial class imagery_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(cs_form));
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView_layout = new System.Windows.Forms.DataGridView();
            this.button_set_image = new System.Windows.Forms.Button();
            this.button_select_drawings = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button_remove_all_images = new System.Windows.Forms.Button();
            this.button_attach_images = new System.Windows.Forms.Button();
            this.button_load_dwg_and_images = new System.Windows.Forms.Button();
            this.label_brightness = new System.Windows.Forms.Label();
            this.label_contrast = new System.Windows.Forms.Label();
            this.textBox_image_contrast = new System.Windows.Forms.TextBox();
            this.textBox_image_brightness = new System.Windows.Forms.TextBox();
            this.button_adjust_images = new System.Windows.Forms.Button();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_layout)).BeginInit();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.dataGridView_layout);
            this.panel2.ForeColor = System.Drawing.Color.White;
            this.panel2.Location = new System.Drawing.Point(3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(851, 473);
            this.panel2.TabIndex = 2207;
            // 
            // dataGridView_layout
            // 
            this.dataGridView_layout.AllowUserToAddRows = false;
            this.dataGridView_layout.AllowUserToDeleteRows = false;
            this.dataGridView_layout.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView_layout.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.dataGridView_layout.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView_layout.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.dataGridView_layout.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.dataGridView_layout.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_layout.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_layout.GridColor = System.Drawing.Color.LightGray;
            this.dataGridView_layout.Location = new System.Drawing.Point(0, 0);
            this.dataGridView_layout.Name = "dataGridView_layout";
            this.dataGridView_layout.RowHeadersVisible = false;
            this.dataGridView_layout.Size = new System.Drawing.Size(849, 471);
            this.dataGridView_layout.TabIndex = 21;
            this.dataGridView_layout.TabStop = false;
            this.dataGridView_layout.VirtualMode = true;
            this.dataGridView_layout.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_layout_CellMouseClick);
            // 
            // button_set_image
            // 
            this.button_set_image.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_set_image.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_set_image.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_set_image.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_set_image.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_set_image.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_set_image.ForeColor = System.Drawing.Color.White;
            this.button_set_image.Image = ((System.Drawing.Image)(resources.GetObject("button_set_image.Image")));
            this.button_set_image.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_set_image.Location = new System.Drawing.Point(4, 583);
            this.button_set_image.Name = "button_set_image";
            this.button_set_image.Size = new System.Drawing.Size(185, 28);
            this.button_set_image.TabIndex = 2258;
            this.button_set_image.Text = "Set Image Frame to ZERO";
            this.button_set_image.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_set_image.UseVisualStyleBackColor = false;
            this.button_set_image.Click += new System.EventHandler(this.button_set_imageframe_to_zero_Click);
            // 
            // button_select_drawings
            // 
            this.button_select_drawings.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_select_drawings.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_select_drawings.BackgroundImage")));
            this.button_select_drawings.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_select_drawings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_select_drawings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_select_drawings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_select_drawings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_select_drawings.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_select_drawings.ForeColor = System.Drawing.Color.White;
            this.button_select_drawings.Location = new System.Drawing.Point(5, 481);
            this.button_select_drawings.Name = "button_select_drawings";
            this.button_select_drawings.Size = new System.Drawing.Size(186, 28);
            this.button_select_drawings.TabIndex = 2078;
            this.button_select_drawings.Text = "Select Drawings";
            this.button_select_drawings.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_select_drawings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_select_drawings.UseVisualStyleBackColor = false;
            this.button_select_drawings.Click += new System.EventHandler(this.button_select_drawings_Click);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label_brightness);
            this.panel3.Controls.Add(this.label_contrast);
            this.panel3.Controls.Add(this.textBox_image_contrast);
            this.panel3.Controls.Add(this.textBox_image_brightness);
            this.panel3.Controls.Add(this.button_adjust_images);
            this.panel3.Controls.Add(this.button_remove_all_images);
            this.panel3.Controls.Add(this.button_attach_images);
            this.panel3.Controls.Add(this.button_load_dwg_and_images);
            this.panel3.Controls.Add(this.button_select_drawings);
            this.panel3.Controls.Add(this.button_set_image);
            this.panel3.Controls.Add(this.panel2);
            this.panel3.Location = new System.Drawing.Point(8, 11);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(859, 616);
            this.panel3.TabIndex = 2259;
            // 
            // button_remove_all_images
            // 
            this.button_remove_all_images.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_remove_all_images.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_remove_all_images.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_remove_all_images.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_remove_all_images.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_remove_all_images.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_remove_all_images.ForeColor = System.Drawing.Color.White;
            this.button_remove_all_images.Image = global::Alignment_mdi.Properties.Resources.X_Icon_New_Small;
            this.button_remove_all_images.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_remove_all_images.Location = new System.Drawing.Point(4, 549);
            this.button_remove_all_images.Name = "button_remove_all_images";
            this.button_remove_all_images.Size = new System.Drawing.Size(185, 28);
            this.button_remove_all_images.TabIndex = 2259;
            this.button_remove_all_images.Text = "Remove all images";
            this.button_remove_all_images.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_remove_all_images.UseVisualStyleBackColor = false;
            this.button_remove_all_images.Click += new System.EventHandler(this.button_remove_all_images_Click);
            // 
            // button_attach_images
            // 
            this.button_attach_images.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_attach_images.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_attach_images.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_attach_images.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_attach_images.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_attach_images.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_attach_images.ForeColor = System.Drawing.Color.White;
            this.button_attach_images.Image = global::Alignment_mdi.Properties.Resources.check;
            this.button_attach_images.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_attach_images.Location = new System.Drawing.Point(668, 583);
            this.button_attach_images.Name = "button_attach_images";
            this.button_attach_images.Size = new System.Drawing.Size(185, 28);
            this.button_attach_images.TabIndex = 2259;
            this.button_attach_images.Text = "Attach images";
            this.button_attach_images.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_attach_images.UseVisualStyleBackColor = false;
            this.button_attach_images.Click += new System.EventHandler(this.button_attach_images_Click);
            // 
            // button_load_dwg_and_images
            // 
            this.button_load_dwg_and_images.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_load_dwg_and_images.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_load_dwg_and_images.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_dwg_and_images.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_load_dwg_and_images.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_dwg_and_images.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_load_dwg_and_images.ForeColor = System.Drawing.Color.White;
            this.button_load_dwg_and_images.Image = global::Alignment_mdi.Properties.Resources.Target;
            this.button_load_dwg_and_images.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_load_dwg_and_images.Location = new System.Drawing.Point(668, 549);
            this.button_load_dwg_and_images.Name = "button_load_dwg_and_images";
            this.button_load_dwg_and_images.Size = new System.Drawing.Size(185, 28);
            this.button_load_dwg_and_images.TabIndex = 2260;
            this.button_load_dwg_and_images.Text = "Load imagery.xls";
            this.button_load_dwg_and_images.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_load_dwg_and_images.UseVisualStyleBackColor = false;
            this.button_load_dwg_and_images.Click += new System.EventHandler(this.button_load_dwg_and_images_Click);
            // 
            // label_brightness
            // 
            this.label_brightness.AutoSize = true;
            this.label_brightness.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_brightness.ForeColor = System.Drawing.Color.White;
            this.label_brightness.Location = new System.Drawing.Point(197, 503);
            this.label_brightness.Name = "label_brightness";
            this.label_brightness.Size = new System.Drawing.Size(68, 14);
            this.label_brightness.TabIndex = 2263;
            this.label_brightness.Text = "Brightness";
            // 
            // label_contrast
            // 
            this.label_contrast.AutoSize = true;
            this.label_contrast.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_contrast.ForeColor = System.Drawing.Color.White;
            this.label_contrast.Location = new System.Drawing.Point(276, 503);
            this.label_contrast.Name = "label_contrast";
            this.label_contrast.Size = new System.Drawing.Size(55, 14);
            this.label_contrast.TabIndex = 2264;
            this.label_contrast.Text = "Contrast";
            // 
            // textBox_image_contrast
            // 
            this.textBox_image_contrast.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_image_contrast.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_image_contrast.ForeColor = System.Drawing.Color.White;
            this.textBox_image_contrast.Location = new System.Drawing.Point(282, 520);
            this.textBox_image_contrast.Name = "textBox_image_contrast";
            this.textBox_image_contrast.Size = new System.Drawing.Size(45, 20);
            this.textBox_image_contrast.TabIndex = 2262;
            this.textBox_image_contrast.Text = "50";
            this.textBox_image_contrast.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_image_brightness
            // 
            this.textBox_image_brightness.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_image_brightness.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_image_brightness.ForeColor = System.Drawing.Color.White;
            this.textBox_image_brightness.Location = new System.Drawing.Point(209, 520);
            this.textBox_image_brightness.Name = "textBox_image_brightness";
            this.textBox_image_brightness.Size = new System.Drawing.Size(45, 20);
            this.textBox_image_brightness.TabIndex = 2261;
            this.textBox_image_brightness.Text = "50";
            this.textBox_image_brightness.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button_adjust_images
            // 
            this.button_adjust_images.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_adjust_images.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_adjust_images.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_adjust_images.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_adjust_images.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_adjust_images.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_adjust_images.ForeColor = System.Drawing.Color.White;
            this.button_adjust_images.Image = global::Alignment_mdi.Properties.Resources.Target;
            this.button_adjust_images.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_adjust_images.Location = new System.Drawing.Point(4, 515);
            this.button_adjust_images.Name = "button_adjust_images";
            this.button_adjust_images.Size = new System.Drawing.Size(185, 28);
            this.button_adjust_images.TabIndex = 2259;
            this.button_adjust_images.Text = "Adjust all images";
            this.button_adjust_images.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_adjust_images.UseVisualStyleBackColor = false;
            this.button_adjust_images.Click += new System.EventHandler(this.button_adjust_images_Click);
            // 
            // image_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(877, 637);
            this.Controls.Add(this.panel3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "image_form";
            this.Text = "AGEN_16_Toolz";
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_layout)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView_layout;
        private System.Windows.Forms.Button button_set_image;
        private System.Windows.Forms.Button button_select_drawings;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button_attach_images;
        private System.Windows.Forms.Button button_load_dwg_and_images;
        private System.Windows.Forms.Button button_remove_all_images;
        private System.Windows.Forms.Label label_brightness;
        private System.Windows.Forms.Label label_contrast;
        private System.Windows.Forms.TextBox textBox_image_contrast;
        private System.Windows.Forms.TextBox textBox_image_brightness;
        private System.Windows.Forms.Button button_adjust_images;
    }
}