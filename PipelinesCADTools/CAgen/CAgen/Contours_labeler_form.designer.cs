namespace Alignment_mdi
{
    partial class contours_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(contours_form));
            this.button_draw = new System.Windows.Forms.Button();
            this.label20 = new System.Windows.Forms.Label();
            this.comboBox_scales = new System.Windows.Forms.ComboBox();
            this.panel17 = new System.Windows.Forms.Panel();
            this.radioButton_use_elevation = new System.Windows.Forms.RadioButton();
            this.radioButton_use_OD = new System.Windows.Forms.RadioButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.textBox_replace = new System.Windows.Forms.TextBox();
            this.textBox_suffix = new System.Windows.Forms.TextBox();
            this.textBox_find = new System.Windows.Forms.TextBox();
            this.comboBox_precision = new System.Windows.Forms.ComboBox();
            this.label_kpmp_precision = new System.Windows.Forms.Label();
            this.checkBox_rotate_180 = new System.Windows.Forms.CheckBox();
            this.comboBox_field = new System.Windows.Forms.ComboBox();
            this.label25 = new System.Windows.Forms.Label();
            this.textBox_text_height = new System.Windows.Forms.TextBox();
            this.label_field = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.comboBox_text_styles = new System.Windows.Forms.ComboBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel_header = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.button_minimize = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.label_mm = new System.Windows.Forms.Label();
            this.checkBox_label_only_10 = new System.Windows.Forms.CheckBox();
            this.panel17.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel_header.SuspendLayout();
            this.SuspendLayout();
            // 
            // button_draw
            // 
            this.button_draw.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_draw.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_draw.BackgroundImage")));
            this.button_draw.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_draw.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_draw.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_draw.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_draw.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_draw.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_draw.ForeColor = System.Drawing.Color.White;
            this.button_draw.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_draw.Location = new System.Drawing.Point(168, 291);
            this.button_draw.Name = "button_draw";
            this.button_draw.Size = new System.Drawing.Size(138, 28);
            this.button_draw.TabIndex = 1;
            this.button_draw.Text = "Label Contours";
            this.button_draw.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_draw.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_draw.UseVisualStyleBackColor = false;
            this.button_draw.Click += new System.EventHandler(this.button_draw_Click);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label20.ForeColor = System.Drawing.Color.White;
            this.label20.Location = new System.Drawing.Point(3, 89);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(90, 14);
            this.label20.TabIndex = 43;
            this.label20.Text = "Viewport Scale";
            // 
            // comboBox_scales
            // 
            this.comboBox_scales.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_scales.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_scales.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_scales.ForeColor = System.Drawing.Color.White;
            this.comboBox_scales.FormattingEnabled = true;
            this.comboBox_scales.Items.AddRange(new object[] {
            "1:1",
            "1:10",
            "1:20",
            "1:30",
            "1:40",
            "1:50",
            "1:60",
            "1:100",
            "1:200",
            "1:300",
            "1:400",
            "1:500",
            "1:600"});
            this.comboBox_scales.Location = new System.Drawing.Point(111, 86);
            this.comboBox_scales.Name = "comboBox_scales";
            this.comboBox_scales.Size = new System.Drawing.Size(186, 21);
            this.comboBox_scales.TabIndex = 2090;
            // 
            // panel17
            // 
            this.panel17.Controls.Add(this.radioButton_use_elevation);
            this.panel17.Controls.Add(this.radioButton_use_OD);
            this.panel17.Location = new System.Drawing.Point(3, 3);
            this.panel17.Name = "panel17";
            this.panel17.Size = new System.Drawing.Size(165, 50);
            this.panel17.TabIndex = 2091;
            // 
            // radioButton_use_elevation
            // 
            this.radioButton_use_elevation.AutoSize = true;
            this.radioButton_use_elevation.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.radioButton_use_elevation.ForeColor = System.Drawing.Color.White;
            this.radioButton_use_elevation.Location = new System.Drawing.Point(3, 25);
            this.radioButton_use_elevation.Name = "radioButton_use_elevation";
            this.radioButton_use_elevation.Size = new System.Drawing.Size(120, 18);
            this.radioButton_use_elevation.TabIndex = 115;
            this.radioButton_use_elevation.Text = "Polyline Elevation";
            this.radioButton_use_elevation.UseVisualStyleBackColor = true;
            // 
            // radioButton_use_OD
            // 
            this.radioButton_use_OD.AutoSize = true;
            this.radioButton_use_OD.Checked = true;
            this.radioButton_use_OD.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.radioButton_use_OD.ForeColor = System.Drawing.Color.White;
            this.radioButton_use_OD.Location = new System.Drawing.Point(3, 3);
            this.radioButton_use_OD.Name = "radioButton_use_OD";
            this.radioButton_use_OD.Size = new System.Drawing.Size(110, 18);
            this.radioButton_use_OD.TabIndex = 114;
            this.radioButton_use_OD.TabStop = true;
            this.radioButton_use_OD.Text = "Use Object Data";
            this.radioButton_use_OD.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.textBox_replace);
            this.panel1.Controls.Add(this.textBox_suffix);
            this.panel1.Controls.Add(this.textBox_find);
            this.panel1.Controls.Add(this.comboBox_precision);
            this.panel1.Controls.Add(this.label_kpmp_precision);
            this.panel1.Controls.Add(this.checkBox_label_only_10);
            this.panel1.Controls.Add(this.checkBox_rotate_180);
            this.panel1.Controls.Add(this.label20);
            this.panel1.Controls.Add(this.comboBox_scales);
            this.panel1.Controls.Add(this.comboBox_field);
            this.panel1.Controls.Add(this.label25);
            this.panel1.Controls.Add(this.button_draw);
            this.panel1.Controls.Add(this.textBox_text_height);
            this.panel1.Controls.Add(this.panel17);
            this.panel1.Controls.Add(this.label_field);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.comboBox_text_styles);
            this.panel1.ForeColor = System.Drawing.Color.White;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(311, 324);
            this.panel1.TabIndex = 2;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(190, 208);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 14);
            this.label2.TabIndex = 2141;
            this.label2.Text = "with";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(3, 233);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(38, 14);
            this.label3.TabIndex = 2141;
            this.label3.Text = "Suffix";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.Color.White;
            this.label14.Location = new System.Drawing.Point(3, 207);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(86, 14);
            this.label14.TabIndex = 2141;
            this.label14.Text = "Replace string";
            // 
            // textBox_replace
            // 
            this.textBox_replace.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_replace.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_replace.ForeColor = System.Drawing.Color.White;
            this.textBox_replace.Location = new System.Drawing.Point(225, 205);
            this.textBox_replace.Name = "textBox_replace";
            this.textBox_replace.Size = new System.Drawing.Size(72, 20);
            this.textBox_replace.TabIndex = 2140;
            // 
            // textBox_suffix
            // 
            this.textBox_suffix.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_suffix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_suffix.ForeColor = System.Drawing.Color.White;
            this.textBox_suffix.Location = new System.Drawing.Point(110, 231);
            this.textBox_suffix.Name = "textBox_suffix";
            this.textBox_suffix.Size = new System.Drawing.Size(187, 20);
            this.textBox_suffix.TabIndex = 2140;
            // 
            // textBox_find
            // 
            this.textBox_find.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_find.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_find.ForeColor = System.Drawing.Color.White;
            this.textBox_find.Location = new System.Drawing.Point(110, 205);
            this.textBox_find.Name = "textBox_find";
            this.textBox_find.Size = new System.Drawing.Size(79, 20);
            this.textBox_find.TabIndex = 2140;
            // 
            // comboBox_precision
            // 
            this.comboBox_precision.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_precision.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_precision.ForeColor = System.Drawing.Color.White;
            this.comboBox_precision.FormattingEnabled = true;
            this.comboBox_precision.Items.AddRange(new object[] {
            "0",
            "0.0",
            "0.00",
            "0.000"});
            this.comboBox_precision.Location = new System.Drawing.Point(111, 168);
            this.comboBox_precision.Name = "comboBox_precision";
            this.comboBox_precision.Size = new System.Drawing.Size(78, 21);
            this.comboBox_precision.TabIndex = 2139;
            this.comboBox_precision.Text = "0";
            // 
            // label_kpmp_precision
            // 
            this.label_kpmp_precision.AutoSize = true;
            this.label_kpmp_precision.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_kpmp_precision.ForeColor = System.Drawing.Color.White;
            this.label_kpmp_precision.Location = new System.Drawing.Point(3, 171);
            this.label_kpmp_precision.Name = "label_kpmp_precision";
            this.label_kpmp_precision.Size = new System.Drawing.Size(59, 14);
            this.label_kpmp_precision.TabIndex = 2138;
            this.label_kpmp_precision.Text = "Precision";
            // 
            // checkBox_rotate_180
            // 
            this.checkBox_rotate_180.AutoSize = true;
            this.checkBox_rotate_180.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.checkBox_rotate_180.ForeColor = System.Drawing.Color.White;
            this.checkBox_rotate_180.Location = new System.Drawing.Point(215, 142);
            this.checkBox_rotate_180.Name = "checkBox_rotate_180";
            this.checkBox_rotate_180.Size = new System.Drawing.Size(82, 18);
            this.checkBox_rotate_180.TabIndex = 2120;
            this.checkBox_rotate_180.Text = "Rotate 180";
            this.checkBox_rotate_180.UseVisualStyleBackColor = true;
            // 
            // comboBox_field
            // 
            this.comboBox_field.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_field.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_field.ForeColor = System.Drawing.Color.White;
            this.comboBox_field.FormattingEnabled = true;
            this.comboBox_field.Location = new System.Drawing.Point(111, 59);
            this.comboBox_field.Name = "comboBox_field";
            this.comboBox_field.Size = new System.Drawing.Size(186, 21);
            this.comboBox_field.TabIndex = 2115;
            this.comboBox_field.DropDown += new System.EventHandler(this.button_load_od_field_to_combobox_dropdown);
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label25.ForeColor = System.Drawing.Color.White;
            this.label25.Location = new System.Drawing.Point(3, 142);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(68, 14);
            this.label25.TabIndex = 115;
            this.label25.Text = "Text Height";
            // 
            // textBox_text_height
            // 
            this.textBox_text_height.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_text_height.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_text_height.ForeColor = System.Drawing.Color.White;
            this.textBox_text_height.Location = new System.Drawing.Point(111, 140);
            this.textBox_text_height.Name = "textBox_text_height";
            this.textBox_text_height.Size = new System.Drawing.Size(45, 20);
            this.textBox_text_height.TabIndex = 116;
            this.textBox_text_height.Text = "0.08";
            this.textBox_text_height.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label_field
            // 
            this.label_field.AutoSize = true;
            this.label_field.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_field.ForeColor = System.Drawing.Color.White;
            this.label_field.Location = new System.Drawing.Point(3, 62);
            this.label_field.Name = "label_field";
            this.label_field.Size = new System.Drawing.Size(97, 14);
            this.label_field.TabIndex = 2118;
            this.label_field.Text = "Object Data Field";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label12.ForeColor = System.Drawing.Color.White;
            this.label12.Location = new System.Drawing.Point(3, 116);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(56, 14);
            this.label12.TabIndex = 2118;
            this.label12.Text = "Text Syle";
            // 
            // comboBox_text_styles
            // 
            this.comboBox_text_styles.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_text_styles.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_text_styles.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_text_styles.ForeColor = System.Drawing.Color.White;
            this.comboBox_text_styles.FormattingEnabled = true;
            this.comboBox_text_styles.Location = new System.Drawing.Point(111, 113);
            this.comboBox_text_styles.Name = "comboBox_text_styles";
            this.comboBox_text_styles.Size = new System.Drawing.Size(186, 21);
            this.comboBox_text_styles.TabIndex = 2119;
            this.comboBox_text_styles.DropDown += new System.EventHandler(this.button_refresh_text_styles_Click);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.panel1);
            this.panel3.Location = new System.Drawing.Point(7, 48);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(321, 332);
            this.panel3.TabIndex = 0;
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel_header.Controls.Add(this.label1);
            this.panel_header.Controls.Add(this.button_minimize);
            this.panel_header.Controls.Add(this.button_Exit);
            this.panel_header.Controls.Add(this.label_mm);
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(338, 39);
            this.panel_header.TabIndex = 42;
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
            this.label1.Size = new System.Drawing.Size(112, 16);
            this.label1.TabIndex = 164;
            this.label1.Text = "Label Contours";
            this.label1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.label1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.label1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
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
            this.button_minimize.Location = new System.Drawing.Point(265, 3);
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
            this.button_Exit.Location = new System.Drawing.Point(301, 4);
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
            this.label_mm.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.label_mm.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.label_mm.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // checkBox_label_only_10
            // 
            this.checkBox_label_only_10.AutoSize = true;
            this.checkBox_label_only_10.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.checkBox_label_only_10.ForeColor = System.Drawing.Color.White;
            this.checkBox_label_only_10.Location = new System.Drawing.Point(6, 257);
            this.checkBox_label_only_10.Name = "checkBox_label_only_10";
            this.checkBox_label_only_10.Size = new System.Drawing.Size(159, 18);
            this.checkBox_label_only_10.TabIndex = 2120;
            this.checkBox_label_only_10.Text = "Label only 10 ft intervals";
            this.checkBox_label_only_10.UseVisualStyleBackColor = true;
            // 
            // contours_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(334, 392);
            this.Controls.Add(this.panel_header);
            this.Controls.Add(this.panel3);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "contours_form";
            this.Text = "AGENProjectSetup";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel17.ResumeLayout(false);
            this.panel17.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button_draw;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ComboBox comboBox_text_styles;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox comboBox_field;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.TextBox textBox_text_height;
        private System.Windows.Forms.ComboBox comboBox_scales;
        private System.Windows.Forms.Panel panel17;
        private System.Windows.Forms.RadioButton radioButton_use_elevation;
        private System.Windows.Forms.RadioButton radioButton_use_OD;
        private System.Windows.Forms.Label label_field;
        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.CheckBox checkBox_rotate_180;
        private System.Windows.Forms.ComboBox comboBox_precision;
        private System.Windows.Forms.Label label_kpmp_precision;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox textBox_replace;
        private System.Windows.Forms.TextBox textBox_find;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_suffix;
        private System.Windows.Forms.CheckBox checkBox_label_only_10;
    }
}
