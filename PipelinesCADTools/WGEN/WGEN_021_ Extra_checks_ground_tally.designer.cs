using System.Drawing;

namespace Alignment_mdi
{
    partial class Wgen_extra_checks
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
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.button_x = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.panel14 = new System.Windows.Forms.Panel();
            this.checkBox_dj_vs_x_ray = new System.Windows.Forms.CheckBox();
            this.panel3.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel14.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.panel7);
            this.panel3.Controls.Add(this.panel14);
            this.panel3.Location = new System.Drawing.Point(2, 1);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(279, 59);
            this.panel3.TabIndex = 0;
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.button_x);
            this.panel7.Controls.Add(this.label2);
            this.panel7.Location = new System.Drawing.Point(3, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(270, 25);
            this.panel7.TabIndex = 2135;
            // 
            // button_x
            // 
            this.button_x.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_x.Image = global::Alignment_mdi.Properties.Resources.close;
            this.button_x.Location = new System.Drawing.Point(246, 0);
            this.button_x.Name = "button_x";
            this.button_x.Size = new System.Drawing.Size(20, 20);
            this.button_x.TabIndex = 2136;
            this.button_x.UseVisualStyleBackColor = true;
            this.button_x.Click += new System.EventHandler(this.button_x_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label2.Location = new System.Drawing.Point(3, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 18);
            this.label2.TabIndex = 2054;
            this.label2.Text = "Extra checks";
            // 
            // panel14
            // 
            this.panel14.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel14.Controls.Add(this.checkBox_dj_vs_x_ray);
            this.panel14.Location = new System.Drawing.Point(3, 27);
            this.panel14.Name = "panel14";
            this.panel14.Size = new System.Drawing.Size(270, 27);
            this.panel14.TabIndex = 0;
            // 
            // checkBox_dj_vs_x_ray
            // 
            this.checkBox_dj_vs_x_ray.AutoSize = true;
            this.checkBox_dj_vs_x_ray.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.checkBox_dj_vs_x_ray.ForeColor = System.Drawing.Color.White;
            this.checkBox_dj_vs_x_ray.Location = new System.Drawing.Point(6, 5);
            this.checkBox_dj_vs_x_ray.Name = "checkBox_dj_vs_x_ray";
            this.checkBox_dj_vs_x_ray.Size = new System.Drawing.Size(209, 18);
            this.checkBox_dj_vs_x_ray.TabIndex = 2137;
            this.checkBox_dj_vs_x_ray.Text = "Perform check Dj # against Xray #";
            this.checkBox_dj_vs_x_ray.UseVisualStyleBackColor = true;
            this.checkBox_dj_vs_x_ray.CheckedChanged += new System.EventHandler(this.checkBox_dj_vs_x_ray_CheckedChanged);
            // 
            // Wgen_extra_checks
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(281, 60);
            this.Controls.Add(this.panel3);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Wgen_extra_checks";
            this.panel3.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.panel14.ResumeLayout(false);
            this.panel14.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBox_dj_vs_x_ray;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Panel panel14;
        private System.Windows.Forms.Button button_x;
    }
}
