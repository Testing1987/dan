namespace Alignment_mdi
{
    partial class Wgen_Blank_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Wgen_Blank_form));
            this.panel_logo = new System.Windows.Forms.Panel();
            this.label_wait = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // panel_logo
            // 
            this.panel_logo.BackColor = System.Drawing.Color.White;
            this.panel_logo.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("panel_logo.BackgroundImage")));
            this.panel_logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel_logo.Location = new System.Drawing.Point(12, 12);
            this.panel_logo.Name = "panel_logo";
            this.panel_logo.Size = new System.Drawing.Size(127, 100);
            this.panel_logo.TabIndex = 35;
            // 
            // label_wait
            // 
            this.label_wait.AutoSize = true;
            this.label_wait.Font = new System.Drawing.Font("Arial Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_wait.ForeColor = System.Drawing.Color.White;
            this.label_wait.Location = new System.Drawing.Point(243, 200);
            this.label_wait.Name = "label_wait";
            this.label_wait.Size = new System.Drawing.Size(224, 108);
            this.label_wait.TabIndex = 36;
            this.label_wait.Text = "Wait please....\r\n   Wait please....\r\n       Wait please....\r\n           Wait plea" +
    "se....";
            this.label_wait.Visible = false;
            // 
            // Wgen_Blank_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(1150, 637);
            this.Controls.Add(this.label_wait);
            this.Controls.Add(this.panel_logo);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Wgen_Blank_form";
            this.Text = "Home";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel_logo;
        private System.Windows.Forms.Label label_wait;
    }
}