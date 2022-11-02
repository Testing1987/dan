namespace Alignment_mdi
{
    partial class AGEN_band_filter
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
            this.panel_segments = new System.Windows.Forms.Panel();
            this.checkBox_pi_show_stations = new System.Windows.Forms.CheckBox();
            this.button_remove_band = new System.Windows.Forms.Button();
            this.panel_segments.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_segments
            // 
            this.panel_segments.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_segments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_segments.Controls.Add(this.button_remove_band);
            this.panel_segments.Controls.Add(this.checkBox_pi_show_stations);
            this.panel_segments.Location = new System.Drawing.Point(3, 4);
            this.panel_segments.Name = "panel_segments";
            this.panel_segments.Size = new System.Drawing.Size(196, 112);
            this.panel_segments.TabIndex = 2056;
            // 
            // checkBox_pi_show_stations
            // 
            this.checkBox_pi_show_stations.AutoSize = true;
            this.checkBox_pi_show_stations.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.checkBox_pi_show_stations.ForeColor = System.Drawing.Color.White;
            this.checkBox_pi_show_stations.Location = new System.Drawing.Point(3, 3);
            this.checkBox_pi_show_stations.Name = "checkBox_pi_show_stations";
            this.checkBox_pi_show_stations.Size = new System.Drawing.Size(117, 18);
            this.checkBox_pi_show_stations.TabIndex = 2138;
            this.checkBox_pi_show_stations.Text = "Ownership band";
            this.checkBox_pi_show_stations.UseVisualStyleBackColor = true;
            // 
            // button_remove_band
            // 
            this.button_remove_band.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_remove_band.BackgroundImage = global::Alignment_mdi.Properties.Resources.X_Icon_New_Small;
            this.button_remove_band.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_remove_band.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gold;
            this.button_remove_band.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.button_remove_band.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_remove_band.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_remove_band.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_remove_band.Location = new System.Drawing.Point(168, 86);
            this.button_remove_band.Name = "button_remove_band";
            this.button_remove_band.Size = new System.Drawing.Size(21, 21);
            this.button_remove_band.TabIndex = 2139;
            this.button_remove_band.TabStop = false;
            this.button_remove_band.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_remove_band.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_remove_band.UseVisualStyleBackColor = false;
            this.button_remove_band.Click += new System.EventHandler(this.button_remove_band_Click);
            // 
            // AGEN_band_filter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(204, 121);
            this.Controls.Add(this.panel_segments);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AGEN_band_filter";
            this.Text = "AGEN_band_field";
            this.panel_segments.ResumeLayout(false);
            this.panel_segments.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_segments;
        private System.Windows.Forms.CheckBox checkBox_pi_show_stations;
        private System.Windows.Forms.Button button_remove_band;
    }
}