
namespace Alignment_mdi
{
    partial class laypurge_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(laypurge_form));
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView_layout = new System.Windows.Forms.DataGridView();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button_select_drawings = new System.Windows.Forms.Button();
            this.button_remove_empty_layers_list = new System.Windows.Forms.Button();
            this.button_remove_empty_layers_dwg = new System.Windows.Forms.Button();
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
            this.panel2.Location = new System.Drawing.Point(3, 34);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(851, 541);
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
            this.dataGridView_layout.Size = new System.Drawing.Size(849, 539);
            this.dataGridView_layout.TabIndex = 21;
            this.dataGridView_layout.TabStop = false;
            this.dataGridView_layout.VirtualMode = true;
            this.dataGridView_layout.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_layout_CellMouseClick);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.button_select_drawings);
            this.panel3.Controls.Add(this.button_remove_empty_layers_dwg);
            this.panel3.Controls.Add(this.button_remove_empty_layers_list);
            this.panel3.Controls.Add(this.panel2);
            this.panel3.Location = new System.Drawing.Point(8, 11);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(859, 616);
            this.panel3.TabIndex = 2259;
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
            this.button_select_drawings.Location = new System.Drawing.Point(3, 3);
            this.button_select_drawings.Name = "button_select_drawings";
            this.button_select_drawings.Size = new System.Drawing.Size(140, 28);
            this.button_select_drawings.TabIndex = 2078;
            this.button_select_drawings.Text = "Select Drawings";
            this.button_select_drawings.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_select_drawings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_select_drawings.UseVisualStyleBackColor = false;
            this.button_select_drawings.Click += new System.EventHandler(this.button_select_drawings_Click);
            // 
            // button_remove_empty_layers_list
            // 
            this.button_remove_empty_layers_list.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_remove_empty_layers_list.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_remove_empty_layers_list.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_remove_empty_layers_list.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_remove_empty_layers_list.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_remove_empty_layers_list.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_remove_empty_layers_list.ForeColor = System.Drawing.Color.White;
            this.button_remove_empty_layers_list.Image = ((System.Drawing.Image)(resources.GetObject("button_remove_empty_layers_list.Image")));
            this.button_remove_empty_layers_list.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_remove_empty_layers_list.Location = new System.Drawing.Point(638, 581);
            this.button_remove_empty_layers_list.Name = "button_remove_empty_layers_list";
            this.button_remove_empty_layers_list.Size = new System.Drawing.Size(215, 28);
            this.button_remove_empty_layers_list.TabIndex = 2258;
            this.button_remove_empty_layers_list.Text = "Remove empty layers from list";
            this.button_remove_empty_layers_list.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_remove_empty_layers_list.UseVisualStyleBackColor = false;
            this.button_remove_empty_layers_list.Click += new System.EventHandler(this.button_vp2poly_Click);
            // 
            // button_remove_empty_layers_dwg
            // 
            this.button_remove_empty_layers_dwg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_remove_empty_layers_dwg.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_remove_empty_layers_dwg.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_remove_empty_layers_dwg.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_remove_empty_layers_dwg.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_remove_empty_layers_dwg.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_remove_empty_layers_dwg.ForeColor = System.Drawing.Color.White;
            this.button_remove_empty_layers_dwg.Image = ((System.Drawing.Image)(resources.GetObject("button_remove_empty_layers_dwg.Image")));
            this.button_remove_empty_layers_dwg.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_remove_empty_layers_dwg.Location = new System.Drawing.Point(4, 580);
            this.button_remove_empty_layers_dwg.Name = "button_remove_empty_layers_dwg";
            this.button_remove_empty_layers_dwg.Size = new System.Drawing.Size(265, 28);
            this.button_remove_empty_layers_dwg.TabIndex = 2258;
            this.button_remove_empty_layers_dwg.Text = "Remove empty layers from curent dwg";
            this.button_remove_empty_layers_dwg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_remove_empty_layers_dwg.UseVisualStyleBackColor = false;
            this.button_remove_empty_layers_dwg.Click += new System.EventHandler(this.button_remove_empty_layers_current_dwg_Click);
            // 
            // laypurge_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(877, 637);
            this.Controls.Add(this.panel3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "laypurge_form";
            this.Text = "AGEN_16_Toolz";
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_layout)).EndInit();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView_layout;
        private System.Windows.Forms.Button button_remove_empty_layers_list;
        private System.Windows.Forms.Button button_select_drawings;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button_remove_empty_layers_dwg;
    }
}