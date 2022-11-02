namespace Alignment_mdi
{
    partial class Filter_Box
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView_Filter = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.checkBox_Filter_Select_All = new System.Windows.Forms.CheckBox();
            this.button_Filter = new System.Windows.Forms.Button();
            this.panel_filter = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Filter)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel_filter.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView_Filter
            // 
            this.dataGridView_Filter.AllowUserToAddRows = false;
            this.dataGridView_Filter.AllowUserToDeleteRows = false;
            this.dataGridView_Filter.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.dataGridView_Filter.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView_Filter.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_Filter.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_Filter.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_Filter.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView_Filter.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridView_Filter.Location = new System.Drawing.Point(0, 53);
            this.dataGridView_Filter.Name = "dataGridView_Filter";
            this.dataGridView_Filter.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dataGridView_Filter.RowHeadersVisible = false;
            this.dataGridView_Filter.RowHeadersWidth = 20;
            this.dataGridView_Filter.Size = new System.Drawing.Size(172, 104);
            this.dataGridView_Filter.TabIndex = 2216;
            this.dataGridView_Filter.CellMouseMove += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_Filter_CellMouseMove);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel2.Controls.Add(this.checkBox_Filter_Select_All);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 27);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(172, 26);
            this.panel2.TabIndex = 2218;
            // 
            // checkBox_Filter_Select_All
            // 
            this.checkBox_Filter_Select_All.AutoSize = true;
            this.checkBox_Filter_Select_All.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkBox_Filter_Select_All.ForeColor = System.Drawing.Color.White;
            this.checkBox_Filter_Select_All.Location = new System.Drawing.Point(3, 4);
            this.checkBox_Filter_Select_All.Name = "checkBox_Filter_Select_All";
            this.checkBox_Filter_Select_All.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkBox_Filter_Select_All.Size = new System.Drawing.Size(70, 17);
            this.checkBox_Filter_Select_All.TabIndex = 2472;
            this.checkBox_Filter_Select_All.Text = "Select All";
            this.checkBox_Filter_Select_All.UseVisualStyleBackColor = true;
            this.checkBox_Filter_Select_All.Click += new System.EventHandler(this.checkBox_Filter_Select_All_Click);
            // 
            // button_Filter
            // 
            this.button_Filter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.button_Filter.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_Filter.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_Filter.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_Filter.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_Filter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Filter.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_Filter.ForeColor = System.Drawing.Color.White;
            this.button_Filter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_Filter.Location = new System.Drawing.Point(0, 0);
            this.button_Filter.Name = "button_Filter";
            this.button_Filter.Size = new System.Drawing.Size(172, 27);
            this.button_Filter.TabIndex = 2223;
            this.button_Filter.Text = "Filter";
            this.button_Filter.UseVisualStyleBackColor = false;
            this.button_Filter.Click += new System.EventHandler(this.button_Filter_Click);
            // 
            // panel_filter
            // 
            this.panel_filter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_filter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_filter.Controls.Add(this.dataGridView_Filter);
            this.panel_filter.Controls.Add(this.panel2);
            this.panel_filter.Controls.Add(this.button_Filter);
            this.panel_filter.Location = new System.Drawing.Point(0, 0);
            this.panel_filter.Name = "panel_filter";
            this.panel_filter.Size = new System.Drawing.Size(174, 191);
            this.panel_filter.TabIndex = 2218;
            // 
            // Filter_Box
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.ClientSize = new System.Drawing.Size(174, 193);
            this.Controls.Add(this.panel_filter);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Filter_Box";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Filter)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel_filter.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridView_Filter;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckBox checkBox_Filter_Select_All;
        private System.Windows.Forms.Button button_Filter;
        private System.Windows.Forms.Panel panel_filter;
    }
}