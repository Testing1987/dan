namespace Alignment_mdi
{
    partial class AGEN_dwg_selection
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
            this.panel_dwg = new System.Windows.Forms.Panel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_generate_list = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.label40 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button_apply = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel_dwg.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_dwg
            // 
            this.panel_dwg.AutoScroll = true;
            this.panel_dwg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_dwg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_dwg.Controls.Add(this.checkBox1);
            this.panel_dwg.Controls.Add(this.label1);
            this.panel_dwg.Location = new System.Drawing.Point(4, 71);
            this.panel_dwg.Name = "panel_dwg";
            this.panel_dwg.Size = new System.Drawing.Size(270, 320);
            this.panel_dwg.TabIndex = 2056;
            this.panel_dwg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_dwg.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_dwg.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.checkBox1.ForeColor = System.Drawing.Color.White;
            this.checkBox1.Location = new System.Drawing.Point(7, 7);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(15, 14);
            this.checkBox1.TabIndex = 2078;
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(28, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 14);
            this.label1.TabIndex = 94;
            this.label1.Text = "123-123-123-123";
            // 
            // button_generate_list
            // 
            this.button_generate_list.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_generate_list.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_generate_list.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_generate_list.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_generate_list.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_generate_list.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_generate_list.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_generate_list.ForeColor = System.Drawing.Color.White;
            this.button_generate_list.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_generate_list.Location = new System.Drawing.Point(4, 395);
            this.button_generate_list.Name = "button_generate_list";
            this.button_generate_list.Size = new System.Drawing.Size(270, 29);
            this.button_generate_list.TabIndex = 93;
            this.button_generate_list.Text = "Create List!";
            this.button_generate_list.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_generate_list.UseVisualStyleBackColor = false;
            this.button_generate_list.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // button_cancel
            // 
            this.button_cancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_cancel.BackgroundImage = global::Alignment_mdi.Properties.Resources.X_Icon_New_Small;
            this.button_cancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_cancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_cancel.FlatAppearance.BorderSize = 4;
            this.button_cancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_cancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_cancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_cancel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_cancel.ForeColor = System.Drawing.Color.White;
            this.button_cancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_cancel.Location = new System.Drawing.Point(239, 2);
            this.button_cancel.Margin = new System.Windows.Forms.Padding(1);
            this.button_cancel.Name = "button_cancel";
            this.button_cancel.Size = new System.Drawing.Size(30, 30);
            this.button_cancel.TabIndex = 2092;
            this.button_cancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_cancel.UseVisualStyleBackColor = false;
            this.button_cancel.Click += new System.EventHandler(this.button_cancel_Click);
            // 
            // label40
            // 
            this.label40.AutoSize = true;
            this.label40.BackColor = System.Drawing.Color.Transparent;
            this.label40.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label40.ForeColor = System.Drawing.Color.White;
            this.label40.Location = new System.Drawing.Point(3, 3);
            this.label40.Name = "label40";
            this.label40.Size = new System.Drawing.Size(41, 14);
            this.label40.TabIndex = 2094;
            this.label40.Text = "Select";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(3, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 14);
            this.label2.TabIndex = 2094;
            this.label2.Text = "From";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(120, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 14);
            this.label3.TabIndex = 2094;
            this.label3.Text = "To";
            // 
            // button_apply
            // 
            this.button_apply.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_apply.BackgroundImage = global::Alignment_mdi.Properties.Resources.check;
            this.button_apply.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_apply.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button_apply.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_apply.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_apply.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_apply.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_apply.ForeColor = System.Drawing.Color.White;
            this.button_apply.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_apply.Location = new System.Drawing.Point(244, 40);
            this.button_apply.Margin = new System.Windows.Forms.Padding(1);
            this.button_apply.Name = "button_apply";
            this.button_apply.Size = new System.Drawing.Size(20, 20);
            this.button_apply.TabIndex = 2092;
            this.button_apply.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_apply.UseVisualStyleBackColor = false;
            this.button_apply.Click += new System.EventHandler(this.button_apply_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox1.Font = new System.Drawing.Font("Arial", 7F);
            this.comboBox1.ForeColor = System.Drawing.Color.White;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(3, 40);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(115, 20);
            this.comboBox1.TabIndex = 2095;
            // 
            // comboBox2
            // 
            this.comboBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox2.Font = new System.Drawing.Font("Arial", 7F);
            this.comboBox2.ForeColor = System.Drawing.Color.White;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(123, 40);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(115, 20);
            this.comboBox2.TabIndex = 2095;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.button_cancel);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label40);
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.button_apply);
            this.panel1.Location = new System.Drawing.Point(4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(270, 65);
            this.panel1.TabIndex = 2056;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // AGEN_dwg_selection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(277, 428);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel_dwg);
            this.Controls.Add(this.button_generate_list);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AGEN_dwg_selection";
            this.Text = "AGEN_0yyy_Inquiry_Tool";
            this.panel_dwg.ResumeLayout(false);
            this.panel_dwg.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_dwg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button button_generate_list;
        private System.Windows.Forms.Button button_cancel;
        private System.Windows.Forms.Label label40;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button_apply;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Panel panel1;
    }
}