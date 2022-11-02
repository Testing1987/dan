namespace Alignment_mdi
{
    partial class pgen_vp2ms
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
            this.Button_pick_vp = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel_err = new System.Windows.Forms.Panel();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button_draw_poly = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel_err.SuspendLayout();
            this.SuspendLayout();
            // 
            // Button_pick_vp
            // 
            this.Button_pick_vp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.Button_pick_vp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.Button_pick_vp.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.Button_pick_vp.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.Button_pick_vp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.Button_pick_vp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_pick_vp.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.Button_pick_vp.ForeColor = System.Drawing.Color.White;
            this.Button_pick_vp.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Button_pick_vp.Location = new System.Drawing.Point(11, 5);
            this.Button_pick_vp.Name = "Button_pick_vp";
            this.Button_pick_vp.Size = new System.Drawing.Size(242, 31);
            this.Button_pick_vp.TabIndex = 2200;
            this.Button_pick_vp.Text = "Pick Viewport";
            this.Button_pick_vp.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Button_pick_vp.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.Button_pick_vp.UseVisualStyleBackColor = false;
            this.Button_pick_vp.Click += new System.EventHandler(this.Button_pick_vp_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.panel_err);
            this.panel1.Controls.Add(this.button_draw_poly);
            this.panel1.Controls.Add(this.Button_pick_vp);
            this.panel1.Location = new System.Drawing.Point(2, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(317, 618);
            this.panel1.TabIndex = 2201;
            // 
            // panel_err
            // 
            this.panel_err.AutoScroll = true;
            this.panel_err.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_err.Controls.Add(this.textBox1);
            this.panel_err.Location = new System.Drawing.Point(12, 42);
            this.panel_err.Name = "panel_err";
            this.panel_err.Size = new System.Drawing.Size(289, 519);
            this.panel_err.TabIndex = 2202;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.textBox1.ForeColor = System.Drawing.Color.DarkTurquoise;
            this.textBox1.Location = new System.Drawing.Point(3, 4);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(225, 23);
            this.textBox1.TabIndex = 1000;
            this.textBox1.TabStop = false;
            this.textBox1.Visible = false;
            // 
            // button_draw_poly
            // 
            this.button_draw_poly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_draw_poly.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_draw_poly.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_draw_poly.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_draw_poly.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_draw_poly.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_draw_poly.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_draw_poly.ForeColor = System.Drawing.Color.White;
            this.button_draw_poly.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_draw_poly.Location = new System.Drawing.Point(12, 578);
            this.button_draw_poly.Name = "button_draw_poly";
            this.button_draw_poly.Size = new System.Drawing.Size(242, 31);
            this.button_draw_poly.TabIndex = 2200;
            this.button_draw_poly.Text = "Draw Model Space";
            this.button_draw_poly.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_draw_poly.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_draw_poly.UseVisualStyleBackColor = false;
            this.button_draw_poly.Click += new System.EventHandler(this.button_draw_poly_Click);
            // 
            // pgen_vp2ms
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(62)))));
            this.ClientSize = new System.Drawing.Size(445, 625);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "pgen_vp2ms";
            this.panel1.ResumeLayout(false);
            this.panel_err.ResumeLayout(false);
            this.panel_err.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button Button_pick_vp;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button_draw_poly;
        private System.Windows.Forms.Panel panel_err;
        private System.Windows.Forms.TextBox textBox1;
    }
}