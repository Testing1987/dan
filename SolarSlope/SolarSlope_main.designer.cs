namespace Alignment_mdi
{
    partial class Solar_main_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Solar_main_form));
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Slope Analizer");
            this.panel_header = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label_4563476 = new System.Windows.Forms.Label();
            this.panel_inq = new System.Windows.Forms.Panel();
            this.label16 = new System.Windows.Forms.Label();
            this.treeView_inquiry = new System.Windows.Forms.TreeView();
            this.label_iq_treeviewnav = new System.Windows.Forms.Label();
            this.panel_header.SuspendLayout();
            this.panel_inq.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel_header.Controls.Add(this.button2);
            this.panel_header.Controls.Add(this.button3);
            this.panel_header.Controls.Add(this.label_4563476);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(1033, 39);
            this.panel_header.TabIndex = 2058;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_header.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_header.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button2.BackgroundImage")));
            this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button2.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Location = new System.Drawing.Point(965, 9);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(27, 20);
            this.button2.TabIndex = 166;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button_minimize_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button3.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button3.BackgroundImage")));
            this.button3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button3.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button3.FlatAppearance.BorderSize = 0;
            this.button3.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.White;
            this.button3.Location = new System.Drawing.Point(998, 4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(30, 30);
            this.button3.TabIndex = 165;
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // label_4563476
            // 
            this.label_4563476.AutoSize = true;
            this.label_4563476.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.label_4563476.ForeColor = System.Drawing.Color.White;
            this.label_4563476.Location = new System.Drawing.Point(3, 2);
            this.label_4563476.Name = "label_4563476";
            this.label_4563476.Size = new System.Drawing.Size(110, 32);
            this.label_4563476.TabIndex = 3;
            this.label_4563476.Text = "Cogo Points\r\nSlope Analyzer";
            // 
            // panel_inq
            // 
            this.panel_inq.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_inq.Controls.Add(this.label16);
            this.panel_inq.Controls.Add(this.treeView_inquiry);
            this.panel_inq.Controls.Add(this.label_iq_treeviewnav);
            this.panel_inq.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel_inq.Location = new System.Drawing.Point(0, 39);
            this.panel_inq.Name = "panel_inq";
            this.panel_inq.Size = new System.Drawing.Size(153, 600);
            this.panel_inq.TabIndex = 2059;
            this.panel_inq.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_inq.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_inq.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.BackColor = System.Drawing.Color.Transparent;
            this.label16.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label16.ForeColor = System.Drawing.Color.Red;
            this.label16.Location = new System.Drawing.Point(59, 575);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(89, 18);
            this.label16.TabIndex = 2055;
            this.label16.Text = "v Beta-1.01";
            // 
            // treeView_inquiry
            // 
            this.treeView_inquiry.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.treeView_inquiry.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeView_inquiry.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeView_inquiry.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.treeView_inquiry.Indent = 10;
            this.treeView_inquiry.ItemHeight = 30;
            this.treeView_inquiry.Location = new System.Drawing.Point(3, 19);
            this.treeView_inquiry.Name = "treeView_inquiry";
            treeNode1.Name = "Node0";
            treeNode1.Text = "Slope Analizer";
            this.treeView_inquiry.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.treeView_inquiry.ShowLines = false;
            this.treeView_inquiry.Size = new System.Drawing.Size(147, 68);
            this.treeView_inquiry.TabIndex = 47;
            this.treeView_inquiry.BeforeSelect += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView1_BeforeSelect);
            this.treeView_inquiry.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView_inquiry_AfterSelect);
            this.treeView_inquiry.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.treeView_inquiry.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.treeView_inquiry.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // label_iq_treeviewnav
            // 
            this.label_iq_treeviewnav.AutoSize = true;
            this.label_iq_treeviewnav.Dock = System.Windows.Forms.DockStyle.Top;
            this.label_iq_treeviewnav.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.label_iq_treeviewnav.ForeColor = System.Drawing.Color.White;
            this.label_iq_treeviewnav.Location = new System.Drawing.Point(0, 0);
            this.label_iq_treeviewnav.Name = "label_iq_treeviewnav";
            this.label_iq_treeviewnav.Size = new System.Drawing.Size(110, 16);
            this.label_iq_treeviewnav.TabIndex = 44;
            this.label_iq_treeviewnav.Text = "Navigation bar";
            // 
            // Solar_main_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(1033, 639);
            this.Controls.Add(this.panel_inq);
            this.Controls.Add(this.panel_header);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.IsMdiContainer = true;
            this.Name = "Solar_main_form";
            this.Text = "WGEN";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            this.panel_inq.ResumeLayout(false);
            this.panel_inq.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Label label_4563476;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Panel panel_inq;
        private System.Windows.Forms.TreeView treeView_inquiry;
        private System.Windows.Forms.Label label_iq_treeviewnav;
        private System.Windows.Forms.Label label16;
    }
}



