namespace Alignment_mdi
{
    partial class slicer_form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(slicer_form));
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button_export_data_to_excel = new System.Windows.Forms.Button();
            this.button_place_stations_along_cl = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.textBox_z_exclude = new System.Windows.Forms.TextBox();
            this.textBox_start_station = new System.Windows.Forms.TextBox();
            this.checkBox_append_slices = new System.Windows.Forms.CheckBox();
            this.button_slices_loaded = new System.Windows.Forms.Button();
            this.button_cl_loaded = new System.Windows.Forms.Button();
            this.button_generate_aligned_polylines = new System.Windows.Forms.Button();
            this.textBox_scanning_precision = new System.Windows.Forms.TextBox();
            this.button_loft = new System.Windows.Forms.Button();
            this.button_select_slices = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.button_select_centerline = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.panel_header = new System.Windows.Forms.Panel();
            this.button_minimize = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.label_mm = new System.Windows.Forms.Label();
            this.panel5.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel_header.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel5.Controls.Add(this.panel7);
            this.panel5.Controls.Add(this.panel1);
            this.panel5.Location = new System.Drawing.Point(4, 43);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(348, 340);
            this.panel5.TabIndex = 0;
            this.panel5.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel5.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel5.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.label4);
            this.panel7.Location = new System.Drawing.Point(3, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(340, 25);
            this.panel7.TabIndex = 2136;
            this.panel7.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel7.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel7.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label4.Location = new System.Drawing.Point(3, 3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 18);
            this.label4.TabIndex = 2033;
            this.label4.Text = "Slice Align";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.button_export_data_to_excel);
            this.panel1.Controls.Add(this.button_place_stations_along_cl);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.textBox_z_exclude);
            this.panel1.Controls.Add(this.textBox_start_station);
            this.panel1.Controls.Add(this.checkBox_append_slices);
            this.panel1.Controls.Add(this.button_slices_loaded);
            this.panel1.Controls.Add(this.button_cl_loaded);
            this.panel1.Controls.Add(this.button_generate_aligned_polylines);
            this.panel1.Controls.Add(this.textBox_scanning_precision);
            this.panel1.Controls.Add(this.button_loft);
            this.panel1.Controls.Add(this.button_select_slices);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.button_select_centerline);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Location = new System.Drawing.Point(3, 27);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(340, 306);
            this.panel1.TabIndex = 0;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // button_export_data_to_excel
            // 
            this.button_export_data_to_excel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_export_data_to_excel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_export_data_to_excel.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_export_data_to_excel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_export_data_to_excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_export_data_to_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_export_data_to_excel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_export_data_to_excel.ForeColor = System.Drawing.Color.White;
            this.button_export_data_to_excel.Image = ((System.Drawing.Image)(resources.GetObject("button_export_data_to_excel.Image")));
            this.button_export_data_to_excel.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.button_export_data_to_excel.Location = new System.Drawing.Point(129, 258);
            this.button_export_data_to_excel.Name = "button_export_data_to_excel";
            this.button_export_data_to_excel.Size = new System.Drawing.Size(203, 39);
            this.button_export_data_to_excel.TabIndex = 2140;
            this.button_export_data_to_excel.Text = "Export 3dpoly data to excel";
            this.button_export_data_to_excel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_export_data_to_excel.UseVisualStyleBackColor = false;
            this.button_export_data_to_excel.Click += new System.EventHandler(this.button_export_data_to_excel_Click);
            // 
            // button_place_stations_along_cl
            // 
            this.button_place_stations_along_cl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_place_stations_along_cl.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_place_stations_along_cl.BackgroundImage")));
            this.button_place_stations_along_cl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_place_stations_along_cl.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_place_stations_along_cl.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_place_stations_along_cl.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_place_stations_along_cl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_place_stations_along_cl.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_place_stations_along_cl.ForeColor = System.Drawing.Color.White;
            this.button_place_stations_along_cl.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_place_stations_along_cl.Location = new System.Drawing.Point(99, 189);
            this.button_place_stations_along_cl.Name = "button_place_stations_along_cl";
            this.button_place_stations_along_cl.Size = new System.Drawing.Size(236, 29);
            this.button_place_stations_along_cl.TabIndex = 2139;
            this.button_place_stations_along_cl.Text = "select slices and create station text";
            this.button_place_stations_along_cl.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_place_stations_along_cl.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_place_stations_along_cl.UseVisualStyleBackColor = false;
            this.button_place_stations_along_cl.Click += new System.EventHandler(this.button_place_stations_along_cl_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(216, 131);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 14);
            this.label2.TabIndex = 2138;
            this.label2.Text = "of previous slice";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(3, 131);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(153, 14);
            this.label1.TabIndex = 2138;
            this.label1.Text = "don\"t include slices within";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.Color.White;
            this.label14.Location = new System.Drawing.Point(6, 3);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(92, 14);
            this.label14.TabIndex = 2138;
            this.label14.Text = "CL Start Station";
            // 
            // textBox_z_exclude
            // 
            this.textBox_z_exclude.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_z_exclude.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_z_exclude.ForeColor = System.Drawing.Color.White;
            this.textBox_z_exclude.Location = new System.Drawing.Point(162, 129);
            this.textBox_z_exclude.Name = "textBox_z_exclude";
            this.textBox_z_exclude.Size = new System.Drawing.Size(48, 20);
            this.textBox_z_exclude.TabIndex = 2137;
            this.textBox_z_exclude.Text = "0.01";
            // 
            // textBox_start_station
            // 
            this.textBox_start_station.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_start_station.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_start_station.ForeColor = System.Drawing.Color.White;
            this.textBox_start_station.Location = new System.Drawing.Point(124, 1);
            this.textBox_start_station.Name = "textBox_start_station";
            this.textBox_start_station.Size = new System.Drawing.Size(67, 20);
            this.textBox_start_station.TabIndex = 2137;
            this.textBox_start_station.Text = "7169+76.5";
            // 
            // checkBox_append_slices
            // 
            this.checkBox_append_slices.AutoSize = true;
            this.checkBox_append_slices.Checked = true;
            this.checkBox_append_slices.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_append_slices.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.checkBox_append_slices.ForeColor = System.Drawing.Color.White;
            this.checkBox_append_slices.Location = new System.Drawing.Point(211, 95);
            this.checkBox_append_slices.Name = "checkBox_append_slices";
            this.checkBox_append_slices.Size = new System.Drawing.Size(69, 18);
            this.checkBox_append_slices.TabIndex = 43;
            this.checkBox_append_slices.Text = "Append";
            this.checkBox_append_slices.UseVisualStyleBackColor = true;
            // 
            // button_slices_loaded
            // 
            this.button_slices_loaded.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_slices_loaded.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_slices_loaded.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gold;
            this.button_slices_loaded.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.button_slices_loaded.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_slices_loaded.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_slices_loaded.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_slices_loaded.Location = new System.Drawing.Point(175, 92);
            this.button_slices_loaded.Name = "button_slices_loaded";
            this.button_slices_loaded.Size = new System.Drawing.Size(21, 21);
            this.button_slices_loaded.TabIndex = 2136;
            this.button_slices_loaded.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_slices_loaded.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_slices_loaded.UseVisualStyleBackColor = false;
            this.button_slices_loaded.Visible = false;
            // 
            // button_cl_loaded
            // 
            this.button_cl_loaded.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_cl_loaded.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_cl_loaded.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gold;
            this.button_cl_loaded.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.button_cl_loaded.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_cl_loaded.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_cl_loaded.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_cl_loaded.Location = new System.Drawing.Point(175, 56);
            this.button_cl_loaded.Name = "button_cl_loaded";
            this.button_cl_loaded.Size = new System.Drawing.Size(21, 21);
            this.button_cl_loaded.TabIndex = 2136;
            this.button_cl_loaded.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_cl_loaded.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_cl_loaded.UseVisualStyleBackColor = false;
            this.button_cl_loaded.Visible = false;
            // 
            // button_generate_aligned_polylines
            // 
            this.button_generate_aligned_polylines.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_generate_aligned_polylines.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_generate_aligned_polylines.BackgroundImage")));
            this.button_generate_aligned_polylines.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_generate_aligned_polylines.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_generate_aligned_polylines.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_generate_aligned_polylines.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_generate_aligned_polylines.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_generate_aligned_polylines.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_generate_aligned_polylines.ForeColor = System.Drawing.Color.White;
            this.button_generate_aligned_polylines.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_generate_aligned_polylines.Location = new System.Drawing.Point(175, 155);
            this.button_generate_aligned_polylines.Name = "button_generate_aligned_polylines";
            this.button_generate_aligned_polylines.Size = new System.Drawing.Size(157, 28);
            this.button_generate_aligned_polylines.TabIndex = 2135;
            this.button_generate_aligned_polylines.Text = "Generate 3D polylines";
            this.button_generate_aligned_polylines.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_generate_aligned_polylines.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_generate_aligned_polylines.UseVisualStyleBackColor = false;
            this.button_generate_aligned_polylines.Click += new System.EventHandler(this.button_generate_aligned_polylines_Click);
            // 
            // textBox_scanning_precision
            // 
            this.textBox_scanning_precision.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_scanning_precision.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_scanning_precision.ForeColor = System.Drawing.Color.White;
            this.textBox_scanning_precision.Location = new System.Drawing.Point(124, 27);
            this.textBox_scanning_precision.Name = "textBox_scanning_precision";
            this.textBox_scanning_precision.Size = new System.Drawing.Size(40, 20);
            this.textBox_scanning_precision.TabIndex = 6;
            this.textBox_scanning_precision.Text = "1";
            this.textBox_scanning_precision.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button_loft
            // 
            this.button_loft.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_loft.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_loft.BackgroundImage")));
            this.button_loft.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_loft.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_loft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_loft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_loft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_loft.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_loft.ForeColor = System.Drawing.Color.White;
            this.button_loft.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.button_loft.Location = new System.Drawing.Point(177, 224);
            this.button_loft.Name = "button_loft";
            this.button_loft.Size = new System.Drawing.Size(155, 28);
            this.button_loft.TabIndex = 2135;
            this.button_loft.Text = "Create loft surfaces";
            this.button_loft.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_loft.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_loft.UseVisualStyleBackColor = false;
            this.button_loft.Click += new System.EventHandler(this.button_loft_Click);
            // 
            // button_select_slices
            // 
            this.button_select_slices.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_select_slices.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_select_slices.BackgroundImage")));
            this.button_select_slices.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_select_slices.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_select_slices.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_select_slices.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_select_slices.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_select_slices.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_select_slices.ForeColor = System.Drawing.Color.White;
            this.button_select_slices.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.button_select_slices.Location = new System.Drawing.Point(9, 89);
            this.button_select_slices.Name = "button_select_slices";
            this.button_select_slices.Size = new System.Drawing.Size(155, 28);
            this.button_select_slices.TabIndex = 2135;
            this.button_select_slices.Text = "Select slices";
            this.button_select_slices.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_select_slices.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_select_slices.UseVisualStyleBackColor = false;
            this.button_select_slices.Click += new System.EventHandler(this.button_select_slices_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.label5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(6, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(112, 14);
            this.label5.TabIndex = 2046;
            this.label5.Text = "Scanning Precision";
            // 
            // button_select_centerline
            // 
            this.button_select_centerline.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_select_centerline.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_select_centerline.BackgroundImage")));
            this.button_select_centerline.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_select_centerline.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_select_centerline.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_select_centerline.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_select_centerline.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_select_centerline.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_select_centerline.ForeColor = System.Drawing.Color.White;
            this.button_select_centerline.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_select_centerline.Location = new System.Drawing.Point(9, 53);
            this.button_select_centerline.Name = "button_select_centerline";
            this.button_select_centerline.Size = new System.Drawing.Size(155, 28);
            this.button_select_centerline.TabIndex = 2135;
            this.button_select_centerline.Text = "Select centerline";
            this.button_select_centerline.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_select_centerline.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_select_centerline.UseVisualStyleBackColor = false;
            this.button_select_centerline.Click += new System.EventHandler(this.button_select_centerline_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.label7.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(170, 29);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(100, 14);
            this.label7.TabIndex = 2046;
            this.label7.Text = "Decimal Degrees";
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.panel_header.Controls.Add(this.button_minimize);
            this.panel_header.Controls.Add(this.button_Exit);
            this.panel_header.Controls.Add(this.label_mm);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(356, 39);
            this.panel_header.TabIndex = 42;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel_header.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel_header.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
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
            this.button_minimize.Location = new System.Drawing.Point(281, 5);
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
            this.button_Exit.Location = new System.Drawing.Point(317, 6);
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
            this.label_mm.Location = new System.Drawing.Point(3, 10);
            this.label_mm.Name = "label_mm";
            this.label_mm.Size = new System.Drawing.Size(137, 20);
            this.label_mm.TabIndex = 3;
            this.label_mm.Text = "Mott Macdonald";
            // 
            // slicer_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(356, 388);
            this.Controls.Add(this.panel_header);
            this.Controls.Add(this.panel5);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "slicer_form";
            this.Text = "SLICE ALIGN";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel5.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox_scanning_precision;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button_select_centerline;
        private System.Windows.Forms.Button button_generate_aligned_polylines;
        private System.Windows.Forms.Button button_select_slices;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.Button button_slices_loaded;
        private System.Windows.Forms.Button button_cl_loaded;
        private System.Windows.Forms.CheckBox checkBox_append_slices;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox textBox_start_station;
        private System.Windows.Forms.Button button_place_stations_along_cl;
        private System.Windows.Forms.Button button_loft;
        private System.Windows.Forms.Button button_export_data_to_excel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_z_exclude;
    }
}
