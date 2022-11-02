namespace Alignment_mdi
{
    partial class Igen__Start_Page_form
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
            this.label1 = new System.Windows.Forms.Label();
            this.button_agen = new System.Windows.Forms.Button();
            this.button_pt_inq = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(299, 19);
            this.label1.TabIndex = 42;
            this.label1.Text = "To start, please select an option below";
            // 
            // button_agen
            // 
            this.button_agen.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(102)))), ((int)(((byte)(204)))));
            this.button_agen.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_agen.ForeColor = System.Drawing.Color.White;
            this.button_agen.Location = new System.Drawing.Point(15, 43);
            this.button_agen.Name = "button_agen";
            this.button_agen.Size = new System.Drawing.Size(225, 100);
            this.button_agen.TabIndex = 43;
            this.button_agen.Text = "Alignment Sheet Production";
            this.button_agen.UseVisualStyleBackColor = false;
            this.button_agen.Click += new System.EventHandler(this.button_agen_Click);
            // 
            // button_pt_inq
            // 
            this.button_pt_inq.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(102)))), ((int)(((byte)(204)))));
            this.button_pt_inq.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_pt_inq.ForeColor = System.Drawing.Color.White;
            this.button_pt_inq.Location = new System.Drawing.Point(246, 43);
            this.button_pt_inq.Name = "button_pt_inq";
            this.button_pt_inq.Size = new System.Drawing.Size(225, 100);
            this.button_pt_inq.TabIndex = 44;
            this.button_pt_inq.Text = "Data Review";
            this.button_pt_inq.UseVisualStyleBackColor = false;
            this.button_pt_inq.Click += new System.EventHandler(this.button_pt_inq_Click);
            // 
            // _Start_Page_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(1150, 637);
            this.Controls.Add(this.button_pt_inq);
            this.Controls.Add(this.button_agen);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "_Start_Page_form";
            this.Text = "AGEN_0000AAA_User_Pick_Screen";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_agen;
        private System.Windows.Forms.Button button_pt_inq;
    }
}