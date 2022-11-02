namespace Dimensioning
{
    partial class Bearing_and_dist_form
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
            this.radioButton_BD = new System.Windows.Forms.RadioButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.checkBox_thousand_sep = new System.Windows.Forms.CheckBox();
            this.checkBox_no_space_in_bearing = new System.Windows.Forms.CheckBox();
            this.checkBox_tolerance = new System.Windows.Forms.CheckBox();
            this.checkBox_background_mask = new System.Windows.Forms.CheckBox();
            this.comboBox_Label_Position = new System.Windows.Forms.ComboBox();
            this.comboBox_coord_system = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label_precision = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox_Precision = new System.Windows.Forms.ComboBox();
            this.comboBox_Scale = new System.Windows.Forms.ComboBox();
            this.label_textheight = new System.Windows.Forms.Label();
            this.textBox_textheight = new System.Windows.Forms.TextBox();
            this.button_Label = new System.Windows.Forms.Button();
            this.textBox_CT_Index = new System.Windows.Forms.TextBox();
            this.radioButton_add_0_ang_dim = new System.Windows.Forms.RadioButton();
            this.radioButton_draw_arcL = new System.Windows.Forms.RadioButton();
            this.radioButton_PI = new System.Windows.Forms.RadioButton();
            this.radioButton_NE = new System.Windows.Forms.RadioButton();
            this.textBox_LT_Index = new System.Windows.Forms.TextBox();
            this.label_Index = new System.Windows.Forms.Label();
            this.radioButton_CT = new System.Windows.Forms.RadioButton();
            this.radioButton_B = new System.Windows.Forms.RadioButton();
            this.radioButton_tie_acholade = new System.Windows.Forms.RadioButton();
            this.radioButton_D = new System.Windows.Forms.RadioButton();
            this.radioButton_LT = new System.Windows.Forms.RadioButton();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button_label_on_poly = new System.Windows.Forms.Button();
            this.comboBox_blocks = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_current_number = new System.Windows.Forms.TextBox();
            this.textBox_block_height = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // radioButton_BD
            // 
            this.radioButton_BD.AutoSize = true;
            this.radioButton_BD.Checked = true;
            this.radioButton_BD.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_BD.ForeColor = System.Drawing.Color.White;
            this.radioButton_BD.Location = new System.Drawing.Point(8, 26);
            this.radioButton_BD.Name = "radioButton_BD";
            this.radioButton_BD.Size = new System.Drawing.Size(126, 17);
            this.radioButton_BD.TabIndex = 0;
            this.radioButton_BD.TabStop = true;
            this.radioButton_BD.Text = "Bearing and Distance";
            this.radioButton_BD.UseVisualStyleBackColor = true;
            this.radioButton_BD.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.checkBox_thousand_sep);
            this.panel1.Controls.Add(this.checkBox_no_space_in_bearing);
            this.panel1.Controls.Add(this.checkBox_tolerance);
            this.panel1.Controls.Add(this.checkBox_background_mask);
            this.panel1.Controls.Add(this.comboBox_Label_Position);
            this.panel1.Controls.Add(this.comboBox_coord_system);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.button_Label);
            this.panel1.Controls.Add(this.radioButton_BD);
            this.panel1.Controls.Add(this.textBox_CT_Index);
            this.panel1.Controls.Add(this.radioButton_add_0_ang_dim);
            this.panel1.Controls.Add(this.radioButton_draw_arcL);
            this.panel1.Controls.Add(this.radioButton_PI);
            this.panel1.Controls.Add(this.radioButton_NE);
            this.panel1.Controls.Add(this.textBox_LT_Index);
            this.panel1.Controls.Add(this.label_Index);
            this.panel1.Controls.Add(this.radioButton_CT);
            this.panel1.Controls.Add(this.radioButton_B);
            this.panel1.Controls.Add(this.radioButton_tie_acholade);
            this.panel1.Controls.Add(this.radioButton_D);
            this.panel1.Controls.Add(this.radioButton_LT);
            this.panel1.Location = new System.Drawing.Point(6, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(177, 480);
            this.panel1.TabIndex = 2;
            // 
            // checkBox_thousand_sep
            // 
            this.checkBox_thousand_sep.AutoSize = true;
            this.checkBox_thousand_sep.Location = new System.Drawing.Point(7, 332);
            this.checkBox_thousand_sep.Name = "checkBox_thousand_sep";
            this.checkBox_thousand_sep.Size = new System.Drawing.Size(119, 17);
            this.checkBox_thousand_sep.TabIndex = 19;
            this.checkBox_thousand_sep.Text = "Add 1000 separator";
            this.checkBox_thousand_sep.UseVisualStyleBackColor = true;
            // 
            // checkBox_no_space_in_bearing
            // 
            this.checkBox_no_space_in_bearing.AutoSize = true;
            this.checkBox_no_space_in_bearing.Checked = true;
            this.checkBox_no_space_in_bearing.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_no_space_in_bearing.Location = new System.Drawing.Point(50, 73);
            this.checkBox_no_space_in_bearing.Name = "checkBox_no_space_in_bearing";
            this.checkBox_no_space_in_bearing.Size = new System.Drawing.Size(122, 17);
            this.checkBox_no_space_in_bearing.TabIndex = 19;
            this.checkBox_no_space_in_bearing.Text = "No space in Bearing";
            this.checkBox_no_space_in_bearing.UseVisualStyleBackColor = true;
            // 
            // checkBox_tolerance
            // 
            this.checkBox_tolerance.AutoSize = true;
            this.checkBox_tolerance.Location = new System.Drawing.Point(7, 308);
            this.checkBox_tolerance.Name = "checkBox_tolerance";
            this.checkBox_tolerance.Size = new System.Drawing.Size(62, 17);
            this.checkBox_tolerance.TabIndex = 19;
            this.checkBox_tolerance.Text = "Add +/-";
            this.checkBox_tolerance.UseVisualStyleBackColor = true;
            // 
            // checkBox_background_mask
            // 
            this.checkBox_background_mask.AutoSize = true;
            this.checkBox_background_mask.Checked = true;
            this.checkBox_background_mask.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_background_mask.Location = new System.Drawing.Point(7, 284);
            this.checkBox_background_mask.Name = "checkBox_background_mask";
            this.checkBox_background_mask.Size = new System.Drawing.Size(133, 17);
            this.checkBox_background_mask.TabIndex = 19;
            this.checkBox_background_mask.Text = "Use background mask";
            this.checkBox_background_mask.UseVisualStyleBackColor = true;
            // 
            // comboBox_Label_Position
            // 
            this.comboBox_Label_Position.BackColor = System.Drawing.Color.DimGray;
            this.comboBox_Label_Position.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Label_Position.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_Label_Position.ForeColor = System.Drawing.Color.White;
            this.comboBox_Label_Position.FormattingEnabled = true;
            this.comboBox_Label_Position.Items.AddRange(new object[] {
            "Top",
            "Bottom",
            "Curved leader"});
            this.comboBox_Label_Position.Location = new System.Drawing.Point(64, 450);
            this.comboBox_Label_Position.Name = "comboBox_Label_Position";
            this.comboBox_Label_Position.Size = new System.Drawing.Size(81, 21);
            this.comboBox_Label_Position.TabIndex = 18;
            this.comboBox_Label_Position.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectedIndexChanged);
            // 
            // comboBox_coord_system
            // 
            this.comboBox_coord_system.BackColor = System.Drawing.Color.DimGray;
            this.comboBox_coord_system.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_coord_system.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_coord_system.ForeColor = System.Drawing.Color.White;
            this.comboBox_coord_system.FormattingEnabled = true;
            this.comboBox_coord_system.Items.AddRange(new object[] {
            "",
            "NJ83F"});
            this.comboBox_coord_system.Location = new System.Drawing.Point(8, 3);
            this.comboBox_coord_system.Name = "comboBox_coord_system";
            this.comboBox_coord_system.Size = new System.Drawing.Size(127, 21);
            this.comboBox_coord_system.TabIndex = 4;
            this.comboBox_coord_system.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectedIndexChanged);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label_precision);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.comboBox_Precision);
            this.panel2.Controls.Add(this.comboBox_Scale);
            this.panel2.Controls.Add(this.label_textheight);
            this.panel2.Controls.Add(this.textBox_textheight);
            this.panel2.Location = new System.Drawing.Point(0, 355);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(147, 89);
            this.panel2.TabIndex = 17;
            // 
            // label_precision
            // 
            this.label_precision.AutoSize = true;
            this.label_precision.ForeColor = System.Drawing.Color.White;
            this.label_precision.Location = new System.Drawing.Point(3, 10);
            this.label_precision.Name = "label_precision";
            this.label_precision.Size = new System.Drawing.Size(50, 13);
            this.label_precision.TabIndex = 5;
            this.label_precision.Text = "Precision";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(3, 37);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "Scale";
            // 
            // comboBox_Precision
            // 
            this.comboBox_Precision.BackColor = System.Drawing.Color.DimGray;
            this.comboBox_Precision.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Precision.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_Precision.ForeColor = System.Drawing.Color.White;
            this.comboBox_Precision.FormattingEnabled = true;
            this.comboBox_Precision.Items.AddRange(new object[] {
            "0",
            "0.0",
            "0.00",
            "0.000",
            "0.0000"});
            this.comboBox_Precision.Location = new System.Drawing.Point(64, 7);
            this.comboBox_Precision.Name = "comboBox_Precision";
            this.comboBox_Precision.Size = new System.Drawing.Size(81, 21);
            this.comboBox_Precision.TabIndex = 4;
            this.comboBox_Precision.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectedIndexChanged);
            // 
            // comboBox_Scale
            // 
            this.comboBox_Scale.BackColor = System.Drawing.Color.DimGray;
            this.comboBox_Scale.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Scale.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_Scale.ForeColor = System.Drawing.Color.White;
            this.comboBox_Scale.FormattingEnabled = true;
            this.comboBox_Scale.Items.AddRange(new object[] {
            "1:1",
            "1 = 10",
            "1 = 20",
            "1 = 30",
            "1 = 40",
            "1 = 50",
            "1 = 60",
            "1 = 100",
            "1 = 200",
            "1 = 300",
            "1 = 400",
            "1 = 500",
            "1 = 600",
            "1 = 1000",
            "1 = 2000",
            "PSpace"});
            this.comboBox_Scale.Location = new System.Drawing.Point(64, 34);
            this.comboBox_Scale.Name = "comboBox_Scale";
            this.comboBox_Scale.Size = new System.Drawing.Size(81, 21);
            this.comboBox_Scale.TabIndex = 15;
            this.comboBox_Scale.SelectedValueChanged += new System.EventHandler(this.comboBox_SelectedIndexChanged);
            // 
            // label_textheight
            // 
            this.label_textheight.AutoSize = true;
            this.label_textheight.ForeColor = System.Drawing.Color.White;
            this.label_textheight.Location = new System.Drawing.Point(3, 64);
            this.label_textheight.Name = "label_textheight";
            this.label_textheight.Size = new System.Drawing.Size(62, 13);
            this.label_textheight.TabIndex = 6;
            this.label_textheight.Text = "Text Height";
            // 
            // textBox_textheight
            // 
            this.textBox_textheight.BackColor = System.Drawing.Color.DimGray;
            this.textBox_textheight.ForeColor = System.Drawing.Color.White;
            this.textBox_textheight.Location = new System.Drawing.Point(64, 61);
            this.textBox_textheight.Name = "textBox_textheight";
            this.textBox_textheight.Size = new System.Drawing.Size(81, 20);
            this.textBox_textheight.TabIndex = 7;
            this.textBox_textheight.Text = ".08";
            // 
            // button_Label
            // 
            this.button_Label.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Label.ForeColor = System.Drawing.Color.White;
            this.button_Label.Location = new System.Drawing.Point(0, 450);
            this.button_Label.Name = "button_Label";
            this.button_Label.Size = new System.Drawing.Size(61, 21);
            this.button_Label.TabIndex = 10;
            this.button_Label.Text = "Label";
            this.button_Label.UseVisualStyleBackColor = true;
            this.button_Label.Click += new System.EventHandler(this.button_Label_Click);
            // 
            // textBox_CT_Index
            // 
            this.textBox_CT_Index.BackColor = System.Drawing.Color.DimGray;
            this.textBox_CT_Index.ForeColor = System.Drawing.Color.White;
            this.textBox_CT_Index.Location = new System.Drawing.Point(89, 161);
            this.textBox_CT_Index.Name = "textBox_CT_Index";
            this.textBox_CT_Index.Size = new System.Drawing.Size(53, 20);
            this.textBox_CT_Index.TabIndex = 14;
            this.textBox_CT_Index.Text = "C1";
            // 
            // radioButton_add_0_ang_dim
            // 
            this.radioButton_add_0_ang_dim.AutoSize = true;
            this.radioButton_add_0_ang_dim.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_add_0_ang_dim.ForeColor = System.Drawing.Color.White;
            this.radioButton_add_0_ang_dim.Location = new System.Drawing.Point(7, 257);
            this.radioButton_add_0_ang_dim.Name = "radioButton_add_0_ang_dim";
            this.radioButton_add_0_ang_dim.Size = new System.Drawing.Size(147, 17);
            this.radioButton_add_0_ang_dim.TabIndex = 11;
            this.radioButton_add_0_ang_dim.Text = "Format Angular Dimension";
            this.radioButton_add_0_ang_dim.UseVisualStyleBackColor = true;
            this.radioButton_add_0_ang_dim.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // radioButton_draw_arcL
            // 
            this.radioButton_draw_arcL.AutoSize = true;
            this.radioButton_draw_arcL.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_draw_arcL.ForeColor = System.Drawing.Color.White;
            this.radioButton_draw_arcL.Location = new System.Drawing.Point(7, 234);
            this.radioButton_draw_arcL.Name = "radioButton_draw_arcL";
            this.radioButton_draw_arcL.Size = new System.Drawing.Size(106, 17);
            this.radioButton_draw_arcL.TabIndex = 11;
            this.radioButton_draw_arcL.Text = "Draw Curve leder";
            this.radioButton_draw_arcL.UseVisualStyleBackColor = true;
            this.radioButton_draw_arcL.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // radioButton_PI
            // 
            this.radioButton_PI.AutoSize = true;
            this.radioButton_PI.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_PI.ForeColor = System.Drawing.Color.White;
            this.radioButton_PI.Location = new System.Drawing.Point(7, 210);
            this.radioButton_PI.Name = "radioButton_PI";
            this.radioButton_PI.Size = new System.Drawing.Size(63, 17);
            this.radioButton_PI.TabIndex = 11;
            this.radioButton_PI.Text = "Insert PI";
            this.radioButton_PI.UseVisualStyleBackColor = true;
            this.radioButton_PI.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // radioButton_NE
            // 
            this.radioButton_NE.AutoSize = true;
            this.radioButton_NE.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_NE.ForeColor = System.Drawing.Color.White;
            this.radioButton_NE.Location = new System.Drawing.Point(7, 186);
            this.radioButton_NE.Name = "radioButton_NE";
            this.radioButton_NE.Size = new System.Drawing.Size(123, 17);
            this.radioButton_NE.TabIndex = 11;
            this.radioButton_NE.Text = "Northing and Easting";
            this.radioButton_NE.UseVisualStyleBackColor = true;
            this.radioButton_NE.CheckedChanged += new System.EventHandler(this.radioButton_ne);
            // 
            // textBox_LT_Index
            // 
            this.textBox_LT_Index.BackColor = System.Drawing.Color.DimGray;
            this.textBox_LT_Index.ForeColor = System.Drawing.Color.White;
            this.textBox_LT_Index.Location = new System.Drawing.Point(89, 134);
            this.textBox_LT_Index.Name = "textBox_LT_Index";
            this.textBox_LT_Index.Size = new System.Drawing.Size(53, 20);
            this.textBox_LT_Index.TabIndex = 13;
            this.textBox_LT_Index.Text = "L1";
            // 
            // label_Index
            // 
            this.label_Index.AutoSize = true;
            this.label_Index.ForeColor = System.Drawing.Color.White;
            this.label_Index.Location = new System.Drawing.Point(97, 118);
            this.label_Index.Name = "label_Index";
            this.label_Index.Size = new System.Drawing.Size(33, 13);
            this.label_Index.TabIndex = 12;
            this.label_Index.Text = "Index";
            // 
            // radioButton_CT
            // 
            this.radioButton_CT.AutoSize = true;
            this.radioButton_CT.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_CT.ForeColor = System.Drawing.Color.White;
            this.radioButton_CT.Location = new System.Drawing.Point(7, 162);
            this.radioButton_CT.Name = "radioButton_CT";
            this.radioButton_CT.Size = new System.Drawing.Size(82, 17);
            this.radioButton_CT.TabIndex = 9;
            this.radioButton_CT.Text = "Curve Table";
            this.radioButton_CT.UseVisualStyleBackColor = true;
            this.radioButton_CT.CheckedChanged += new System.EventHandler(this.radioButton_CT_CheckedChanged);
            // 
            // radioButton_B
            // 
            this.radioButton_B.AutoSize = true;
            this.radioButton_B.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_B.ForeColor = System.Drawing.Color.White;
            this.radioButton_B.Location = new System.Drawing.Point(8, 50);
            this.radioButton_B.Name = "radioButton_B";
            this.radioButton_B.Size = new System.Drawing.Size(60, 17);
            this.radioButton_B.TabIndex = 2;
            this.radioButton_B.Text = "Bearing";
            this.radioButton_B.UseVisualStyleBackColor = true;
            this.radioButton_B.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // radioButton_tie_acholade
            // 
            this.radioButton_tie_acholade.AutoSize = true;
            this.radioButton_tie_acholade.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_tie_acholade.ForeColor = System.Drawing.Color.White;
            this.radioButton_tie_acholade.Location = new System.Drawing.Point(7, 114);
            this.radioButton_tie_acholade.Name = "radioButton_tie_acholade";
            this.radioButton_tie_acholade.Size = new System.Drawing.Size(84, 17);
            this.radioButton_tie_acholade.TabIndex = 3;
            this.radioButton_tie_acholade.Text = "Tie Distance";
            this.radioButton_tie_acholade.UseVisualStyleBackColor = true;
            this.radioButton_tie_acholade.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // radioButton_D
            // 
            this.radioButton_D.AutoSize = true;
            this.radioButton_D.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_D.ForeColor = System.Drawing.Color.White;
            this.radioButton_D.Location = new System.Drawing.Point(7, 90);
            this.radioButton_D.Name = "radioButton_D";
            this.radioButton_D.Size = new System.Drawing.Size(66, 17);
            this.radioButton_D.TabIndex = 3;
            this.radioButton_D.Text = "Distance";
            this.radioButton_D.UseVisualStyleBackColor = true;
            this.radioButton_D.CheckedChanged += new System.EventHandler(this.radioButton_others);
            // 
            // radioButton_LT
            // 
            this.radioButton_LT.AutoSize = true;
            this.radioButton_LT.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_LT.ForeColor = System.Drawing.Color.White;
            this.radioButton_LT.Location = new System.Drawing.Point(7, 138);
            this.radioButton_LT.Name = "radioButton_LT";
            this.radioButton_LT.Size = new System.Drawing.Size(74, 17);
            this.radioButton_LT.TabIndex = 8;
            this.radioButton_LT.Text = "Line Table";
            this.radioButton_LT.UseVisualStyleBackColor = true;
            this.radioButton_LT.CheckedChanged += new System.EventHandler(this.radioButton_LT_CheckedChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(1, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(201, 518);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.DimGray;
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.ForeColor = System.Drawing.Color.White;
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(193, 492);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "A";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.DimGray;
            this.tabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage2.Controls.Add(this.panel3);
            this.tabPage2.ForeColor = System.Drawing.Color.White;
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(193, 492);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "B";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.DimGray;
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel3.Controls.Add(this.button_label_on_poly);
            this.panel3.Controls.Add(this.comboBox_blocks);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.label7);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.label6);
            this.panel3.Controls.Add(this.textBox_block_height);
            this.panel3.Controls.Add(this.textBox_current_number);
            this.panel3.Location = new System.Drawing.Point(0, 1);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(183, 481);
            this.panel3.TabIndex = 4;
            this.panel3.Click += new System.EventHandler(this.panel3_Click);
            // 
            // button_label_on_poly
            // 
            this.button_label_on_poly.BackColor = System.Drawing.Color.White;
            this.button_label_on_poly.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_label_on_poly.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_label_on_poly.ForeColor = System.Drawing.Color.Black;
            this.button_label_on_poly.Location = new System.Drawing.Point(4, 163);
            this.button_label_on_poly.Name = "button_label_on_poly";
            this.button_label_on_poly.Size = new System.Drawing.Size(172, 21);
            this.button_label_on_poly.TabIndex = 10;
            this.button_label_on_poly.Text = "Label";
            this.button_label_on_poly.UseVisualStyleBackColor = false;
            this.button_label_on_poly.Click += new System.EventHandler(this.button_label_on_poly_Click);
            // 
            // comboBox_blocks
            // 
            this.comboBox_blocks.BackColor = System.Drawing.Color.White;
            this.comboBox_blocks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_blocks.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_blocks.ForeColor = System.Drawing.Color.Black;
            this.comboBox_blocks.FormattingEnabled = true;
            this.comboBox_blocks.Items.AddRange(new object[] {
            ""});
            this.comboBox_blocks.Location = new System.Drawing.Point(4, 110);
            this.comboBox_blocks.Name = "comboBox_blocks";
            this.comboBox_blocks.Size = new System.Drawing.Size(172, 21);
            this.comboBox_blocks.TabIndex = 4;
            this.comboBox_blocks.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(8, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 41);
            this.label1.TabIndex = 5;
            this.label1.Text = "Bearing and \r\nDistance Table along \r\nSelected Polyline";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.DimGray;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(4, 92);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 5;
            this.label7.Text = "Block:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.DimGray;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(3, 63);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(85, 13);
            this.label6.TabIndex = 5;
            this.label6.Text = "Point number:";
            // 
            // textBox_current_number
            // 
            this.textBox_current_number.BackColor = System.Drawing.Color.White;
            this.textBox_current_number.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_current_number.ForeColor = System.Drawing.Color.Black;
            this.textBox_current_number.Location = new System.Drawing.Point(96, 60);
            this.textBox_current_number.Name = "textBox_current_number";
            this.textBox_current_number.Size = new System.Drawing.Size(51, 20);
            this.textBox_current_number.TabIndex = 4;
            this.textBox_current_number.Text = "A01";
            this.textBox_current_number.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_block_height
            // 
            this.textBox_block_height.BackColor = System.Drawing.Color.White;
            this.textBox_block_height.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_block_height.ForeColor = System.Drawing.Color.Black;
            this.textBox_block_height.Location = new System.Drawing.Point(100, 137);
            this.textBox_block_height.Name = "textBox_block_height";
            this.textBox_block_height.Size = new System.Drawing.Size(51, 20);
            this.textBox_block_height.TabIndex = 4;
            this.textBox_block_height.Text = "0.1914";
            this.textBox_block_height.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.DimGray;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(7, 140);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Block Height:";
            // 
            // Bearing_and_dist_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DimGray;
            this.ClientSize = new System.Drawing.Size(204, 524);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "Bearing_and_dist_form";
            this.Text = "BRD";
            this.Load += new System.EventHandler(this.Bearing_and_dist_form_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButton_BD;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comboBox_Precision;
        private System.Windows.Forms.RadioButton radioButton_D;
        private System.Windows.Forms.RadioButton radioButton_B;
        private System.Windows.Forms.TextBox textBox_textheight;
        private System.Windows.Forms.Label label_textheight;
        private System.Windows.Forms.Label label_precision;
        private System.Windows.Forms.RadioButton radioButton_NE;
        private System.Windows.Forms.Button button_Label;
        private System.Windows.Forms.RadioButton radioButton_CT;
        private System.Windows.Forms.RadioButton radioButton_LT;
        private System.Windows.Forms.TextBox textBox_CT_Index;
        private System.Windows.Forms.TextBox textBox_LT_Index;
        private System.Windows.Forms.Label label_Index;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox_Scale;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ComboBox comboBox_Label_Position;
        private System.Windows.Forms.RadioButton radioButton_PI;
        private System.Windows.Forms.ComboBox comboBox_coord_system;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button_label_on_poly;
        private System.Windows.Forms.ComboBox comboBox_blocks;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_current_number;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox_background_mask;
        private System.Windows.Forms.CheckBox checkBox_tolerance;
        private System.Windows.Forms.CheckBox checkBox_thousand_sep;
        private System.Windows.Forms.RadioButton radioButton_draw_arcL;
        private System.Windows.Forms.RadioButton radioButton_tie_acholade;
        private System.Windows.Forms.RadioButton radioButton_add_0_ang_dim;
        private System.Windows.Forms.CheckBox checkBox_no_space_in_bearing;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_block_height;
    }
}
