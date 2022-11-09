namespace Alignment_mdi
{
    partial class _AGEN_mainform
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(_AGEN_mainform));
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Settings");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Border Definition");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Block Attributes");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Station Equations");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Project", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4});
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Sheet Index Setup");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Plan View", new System.Windows.Forms.TreeNode[] {
            treeNode6});
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Ownership");
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("Crossing");
            System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("Material");
            System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("Profile");
            System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("Custom");
            System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("Band Builder", new System.Windows.Forms.TreeNode[] {
            treeNode8,
            treeNode9,
            treeNode10,
            treeNode11,
            treeNode12});
            System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("Create Alignment Sheets");
            System.Windows.Forms.TreeNode treeNode15 = new System.Windows.Forms.TreeNode("Sheet Generation", new System.Windows.Forms.TreeNode[] {
            treeNode14});
            System.Windows.Forms.TreeNode treeNode16 = new System.Windows.Forms.TreeNode("Rename Layout");
            System.Windows.Forms.TreeNode treeNode17 = new System.Windows.Forms.TreeNode("Viewport to poly");
            System.Windows.Forms.TreeNode treeNode18 = new System.Windows.Forms.TreeNode("IMAGERY");
            System.Windows.Forms.TreeNode treeNode19 = new System.Windows.Forms.TreeNode("Extra tools", new System.Windows.Forms.TreeNode[] {
            treeNode16,
            treeNode17,
            treeNode18});
            this.label_projectpath = new System.Windows.Forms.Label();
            this.textBox_config_file_path = new System.Windows.Forms.TextBox();
            this.panel_header = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.button_minimize = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.label_mm = new System.Windows.Forms.Label();
            this.panel11 = new System.Windows.Forms.Panel();
            this.label16 = new System.Windows.Forms.Label();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.label_name_of_treeview = new System.Windows.Forms.Label();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.panel_header.SuspendLayout();
            this.panel11.SuspendLayout();
            this.SuspendLayout();
            // 
            // label_projectpath
            // 
            this.label_projectpath.AutoSize = true;
            this.label_projectpath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label_projectpath.ForeColor = System.Drawing.Color.White;
            this.label_projectpath.Location = new System.Drawing.Point(203, 15);
            this.label_projectpath.Name = "label_projectpath";
            this.label_projectpath.Size = new System.Drawing.Size(107, 13);
            this.label_projectpath.TabIndex = 20;
            this.label_projectpath.Text = "Loaded Project Path:";
            // 
            // textBox_config_file_path
            // 
            this.textBox_config_file_path.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.textBox_config_file_path.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_config_file_path.ForeColor = System.Drawing.Color.White;
            this.textBox_config_file_path.Location = new System.Drawing.Point(316, 3);
            this.textBox_config_file_path.Multiline = true;
            this.textBox_config_file_path.Name = "textBox_config_file_path";
            this.textBox_config_file_path.ReadOnly = true;
            this.textBox_config_file_path.Size = new System.Drawing.Size(692, 27);
            this.textBox_config_file_path.TabIndex = 19;
            this.textBox_config_file_path.Text = "***";
            this.textBox_config_file_path.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.textBox_config_file_path.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.textBox_config_file_path.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel_header.Controls.Add(this.label1);
            this.panel_header.Controls.Add(this.label_projectpath);
            this.panel_header.Controls.Add(this.textBox_config_file_path);
            this.panel_header.Controls.Add(this.button_minimize);
            this.panel_header.Controls.Add(this.button_Exit);
            this.panel_header.Controls.Add(this.label_mm);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(1088, 39);
            this.panel_header.TabIndex = 41;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_header.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_header.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 16);
            this.label1.TabIndex = 164;
            this.label1.Text = "AGEN";
            // 
            // button_minimize
            // 
            this.button_minimize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_minimize.BackgroundImage")));
            this.button_minimize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_minimize.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.FlatAppearance.BorderSize = 0;
            this.button_minimize.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_minimize.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_minimize.ForeColor = System.Drawing.Color.White;
            this.button_minimize.Location = new System.Drawing.Point(1014, 6);
            this.button_minimize.Name = "button_minimize";
            this.button_minimize.Size = new System.Drawing.Size(30, 30);
            this.button_minimize.TabIndex = 162;
            this.button_minimize.UseVisualStyleBackColor = false;
            this.button_minimize.Click += new System.EventHandler(this.button_minimize_Click);
            // 
            // button_Exit
            // 
            this.button_Exit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_Exit.BackgroundImage")));
            this.button_Exit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_Exit.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.FlatAppearance.BorderSize = 0;
            this.button_Exit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Exit.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.ForeColor = System.Drawing.Color.White;
            this.button_Exit.Location = new System.Drawing.Point(1050, 7);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(30, 30);
            this.button_Exit.TabIndex = 161;
            this.button_Exit.UseVisualStyleBackColor = false;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // label_mm
            // 
            this.label_mm.AutoSize = true;
            this.label_mm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_mm.ForeColor = System.Drawing.Color.White;
            this.label_mm.Location = new System.Drawing.Point(3, 1);
            this.label_mm.Name = "label_mm";
            this.label_mm.Size = new System.Drawing.Size(137, 20);
            this.label_mm.TabIndex = 3;
            this.label_mm.Text = "Mott Macdonald";
            // 
            // panel11
            // 
            this.panel11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel11.Controls.Add(this.label16);
            this.panel11.Controls.Add(this.treeView1);
            this.panel11.Controls.Add(this.label_name_of_treeview);
            this.panel11.Controls.Add(this.linkLabel2);
            this.panel11.Controls.Add(this.linkLabel1);
            this.panel11.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel11.Location = new System.Drawing.Point(0, 39);
            this.panel11.Name = "panel11";
            this.panel11.Size = new System.Drawing.Size(211, 637);
            this.panel11.TabIndex = 2022;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.BackColor = System.Drawing.Color.Red;
            this.label16.Font = new System.Drawing.Font("Arial Black", 12F);
            this.label16.ForeColor = System.Drawing.Color.White;
            this.label16.Location = new System.Drawing.Point(143, 609);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(65, 23);
            this.label16.TabIndex = 2055;
            this.label16.Text = "V 4.48";
            // 
            // treeView1
            // 
            this.treeView1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.treeView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeView1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeView1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.treeView1.Indent = 10;
            this.treeView1.ItemHeight = 30;
            this.treeView1.Location = new System.Drawing.Point(0, 27);
            this.treeView1.Name = "treeView1";
            treeNode1.ForeColor = System.Drawing.Color.White;
            treeNode1.Name = "Node0";
            treeNode1.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            treeNode1.Text = "Settings";
            treeNode2.ForeColor = System.Drawing.Color.White;
            treeNode2.Name = "Node1";
            treeNode2.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode2.Text = "Border Definition";
            treeNode3.ForeColor = System.Drawing.Color.White;
            treeNode3.Name = "Node2";
            treeNode3.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode3.Text = "Block Attributes";
            treeNode4.ForeColor = System.Drawing.Color.White;
            treeNode4.Name = "Node4";
            treeNode4.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode4.Text = "Station Equations";
            treeNode5.Name = "Node0";
            treeNode5.Text = "Project";
            treeNode6.ForeColor = System.Drawing.Color.White;
            treeNode6.Name = "Node0";
            treeNode6.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode6.Text = "Sheet Index Setup";
            treeNode7.Name = "Node1";
            treeNode7.Text = "Plan View";
            treeNode8.ForeColor = System.Drawing.Color.White;
            treeNode8.Name = "Node0";
            treeNode8.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode8.Text = "Ownership";
            treeNode9.ForeColor = System.Drawing.Color.White;
            treeNode9.Name = "Node1";
            treeNode9.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode9.Text = "Crossing";
            treeNode10.ForeColor = System.Drawing.Color.White;
            treeNode10.Name = "Node2";
            treeNode10.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode10.Text = "Material";
            treeNode11.ForeColor = System.Drawing.Color.White;
            treeNode11.Name = "Node3";
            treeNode11.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode11.Text = "Profile";
            treeNode12.ForeColor = System.Drawing.Color.White;
            treeNode12.Name = "Node4";
            treeNode12.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode12.Text = "Custom";
            treeNode13.Name = "Node2";
            treeNode13.Text = "Band Builder";
            treeNode14.ForeColor = System.Drawing.Color.White;
            treeNode14.Name = "Node31";
            treeNode14.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode14.Text = "Create Alignment Sheets";
            treeNode15.Name = "Node3";
            treeNode15.Text = "Sheet Generation";
            treeNode16.ForeColor = System.Drawing.Color.White;
            treeNode16.Name = "Node41";
            treeNode16.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode16.Text = "Rename Layout";
            treeNode17.ForeColor = System.Drawing.Color.White;
            treeNode17.Name = "Node42";
            treeNode17.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode17.Text = "Viewport to poly";
            treeNode18.ForeColor = System.Drawing.Color.White;
            treeNode18.Name = "Node43";
            treeNode18.NodeFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            treeNode18.Text = "IMAGERY";
            treeNode19.Name = "Node4";
            treeNode19.Text = "Extra tools";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode5,
            treeNode7,
            treeNode13,
            treeNode15,
            treeNode19});
            this.treeView1.ShowLines = false;
            this.treeView1.Size = new System.Drawing.Size(201, 482);
            this.treeView1.TabIndex = 47;
            this.treeView1.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseClick);
            // 
            // label_name_of_treeview
            // 
            this.label_name_of_treeview.AutoSize = true;
            this.label_name_of_treeview.Dock = System.Windows.Forms.DockStyle.Top;
            this.label_name_of_treeview.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_name_of_treeview.ForeColor = System.Drawing.Color.White;
            this.label_name_of_treeview.Location = new System.Drawing.Point(0, 0);
            this.label_name_of_treeview.Name = "label_name_of_treeview";
            this.label_name_of_treeview.Size = new System.Drawing.Size(156, 24);
            this.label_name_of_treeview.TabIndex = 44;
            this.label_name_of_treeview.Text = "Navigation bar";
            this.label_name_of_treeview.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.label_name_of_treeview.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.label_name_of_treeview.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // linkLabel2
            // 
            this.linkLabel2.ActiveLinkColor = System.Drawing.Color.Yellow;
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.linkLabel2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.linkLabel2.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline;
            this.linkLabel2.LinkColor = System.Drawing.Color.White;
            this.linkLabel2.Location = new System.Drawing.Point(0, 609);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(31, 14);
            this.linkLabel2.TabIndex = 13;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "Help";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // linkLabel1
            // 
            this.linkLabel1.ActiveLinkColor = System.Drawing.Color.Yellow;
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.linkLabel1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline;
            this.linkLabel1.LinkColor = System.Drawing.Color.White;
            this.linkLabel1.Location = new System.Drawing.Point(0, 623);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(96, 14);
            this.linkLabel1.TabIndex = 12;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Contact Support";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // _AGEN_mainform
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.ClientSize = new System.Drawing.Size(1088, 676);
            this.Controls.Add(this.panel11);
            this.Controls.Add(this.panel_header);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.IsMdiContainer = true;
            this.MaximizeBox = false;
            this.Name = "_AGEN_mainform";
            this.Text = "t";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            this.panel11.ResumeLayout(false);
            this.panel11.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TextBox textBox_config_file_path;
        private System.Windows.Forms.Label label_projectpath;
        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.Panel panel11;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label label_name_of_treeview;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Label label16;
    }
}
