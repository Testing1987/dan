using FlatTabControl;

namespace Alignment_mdi
{
    partial class Centerline_form
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Centerline_form));
            this.panel_Matl_Tool = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dataGridView_cl = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label_cl = new System.Windows.Forms.Label();
            this.panel_bottom = new System.Windows.Forms.Panel();
            this.button_save_cl = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_clear_elbows = new System.Windows.Forms.Button();
            this.button_load_dwg_cl = new System.Windows.Forms.Button();
            this.button_load_xl_centerline = new System.Windows.Forms.Button();
            this.comboBox_elbow = new System.Windows.Forms.ComboBox();
            this.label_default_material = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_min_angle = new System.Windows.Forms.TextBox();
            this.comboBox_description = new System.Windows.Forms.ComboBox();
            this.button_filter = new System.Windows.Forms.Button();
            this.button_assign_elbows = new System.Windows.Forms.Button();
            this.button_transfer_to_mat = new System.Windows.Forms.Button();
            this.panel8 = new System.Windows.Forms.Panel();
            this.label_header = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.button_add_stationing = new System.Windows.Forms.Button();
            this.panel_Matl_Tool.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_cl)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel_bottom.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel8.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_Matl_Tool
            // 
            this.panel_Matl_Tool.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_Matl_Tool.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_Matl_Tool.Controls.Add(this.panel1);
            this.panel_Matl_Tool.Controls.Add(this.panel2);
            this.panel_Matl_Tool.Controls.Add(this.panel_bottom);
            this.panel_Matl_Tool.Controls.Add(this.tableLayoutPanel1);
            this.panel_Matl_Tool.Controls.Add(this.panel8);
            this.panel_Matl_Tool.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel_Matl_Tool.ForeColor = System.Drawing.Color.White;
            this.panel_Matl_Tool.Location = new System.Drawing.Point(0, 0);
            this.panel_Matl_Tool.Margin = new System.Windows.Forms.Padding(0);
            this.panel_Matl_Tool.Name = "panel_Matl_Tool";
            this.panel_Matl_Tool.Size = new System.Drawing.Size(997, 502);
            this.panel_Matl_Tool.TabIndex = 2110;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dataGridView_cl);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(185, 57);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(810, 416);
            this.panel1.TabIndex = 2218;
            // 
            // dataGridView_cl
            // 
            this.dataGridView_cl.AllowUserToAddRows = false;
            this.dataGridView_cl.AllowUserToDeleteRows = false;
            this.dataGridView_cl.AllowUserToOrderColumns = true;
            this.dataGridView_cl.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.dataGridView_cl.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView_cl.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_cl.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_cl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_cl.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView_cl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_cl.Location = new System.Drawing.Point(0, 0);
            this.dataGridView_cl.Name = "dataGridView_cl";
            this.dataGridView_cl.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dataGridView_cl.RowHeadersVisible = false;
            this.dataGridView_cl.RowHeadersWidth = 20;
            this.dataGridView_cl.Size = new System.Drawing.Size(810, 416);
            this.dataGridView_cl.TabIndex = 2217;
            this.dataGridView_cl.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_cl_CellClick);
            this.dataGridView_cl.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_cl_CellEndEdit);
            this.dataGridView_cl.CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_cl_CellMouseUp);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label_cl);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(185, 27);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(810, 30);
            this.panel2.TabIndex = 2220;
            // 
            // label_cl
            // 
            this.label_cl.AutoSize = true;
            this.label_cl.BackColor = System.Drawing.Color.Transparent;
            this.label_cl.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_cl.ForeColor = System.Drawing.Color.Red;
            this.label_cl.Location = new System.Drawing.Point(0, 6);
            this.label_cl.Margin = new System.Windows.Forms.Padding(0);
            this.label_cl.Name = "label_cl";
            this.label_cl.Size = new System.Drawing.Size(164, 18);
            this.label_cl.TabIndex = 2055;
            this.label_cl.Text = "Centerline not loaded";
            this.label_cl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel_bottom
            // 
            this.panel_bottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_bottom.Controls.Add(this.button_save_cl);
            this.panel_bottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel_bottom.Location = new System.Drawing.Point(185, 473);
            this.panel_bottom.Name = "panel_bottom";
            this.panel_bottom.Size = new System.Drawing.Size(810, 27);
            this.panel_bottom.TabIndex = 2219;
            // 
            // button_save_cl
            // 
            this.button_save_cl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_save_cl.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_save_cl.BackgroundImage")));
            this.button_save_cl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_save_cl.Dock = System.Windows.Forms.DockStyle.Right;
            this.button_save_cl.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_save_cl.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_save_cl.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_save_cl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_save_cl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_save_cl.ForeColor = System.Drawing.Color.Black;
            this.button_save_cl.Location = new System.Drawing.Point(781, 0);
            this.button_save_cl.Margin = new System.Windows.Forms.Padding(2);
            this.button_save_cl.Name = "button_save_cl";
            this.button_save_cl.Size = new System.Drawing.Size(27, 25);
            this.button_save_cl.TabIndex = 2412;
            this.button_save_cl.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_save_cl.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_save_cl.UseVisualStyleBackColor = false;
            this.button_save_cl.Click += new System.EventHandler(this.button_save_cl_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.button_load_dwg_cl, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_load_xl_centerline, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_transfer_to_mat, 0, 15);
            this.tableLayoutPanel1.Controls.Add(this.label_default_material, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.comboBox_elbow, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.comboBox_description, 0, 6);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 7);
            this.tableLayoutPanel1.Controls.Add(this.textBox_min_angle, 0, 8);
            this.tableLayoutPanel1.Controls.Add(this.button_filter, 0, 9);
            this.tableLayoutPanel1.Controls.Add(this.button_assign_elbows, 0, 10);
            this.tableLayoutPanel1.Controls.Add(this.button_clear_elbows, 0, 14);
            this.tableLayoutPanel1.Controls.Add(this.button_add_stationing, 0, 12);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 27);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 16;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 16F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 9F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 18F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 18F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 42F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 41F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 12F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(185, 473);
            this.tableLayoutPanel1.TabIndex = 2216;
            // 
            // button_clear_elbows
            // 
            this.button_clear_elbows.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_clear_elbows.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_clear_elbows.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_clear_elbows.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_clear_elbows.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_clear_elbows.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_clear_elbows.ForeColor = System.Drawing.Color.White;
            this.button_clear_elbows.Image = ((System.Drawing.Image)(resources.GetObject("button_clear_elbows.Image")));
            this.button_clear_elbows.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_clear_elbows.Location = new System.Drawing.Point(3, 404);
            this.button_clear_elbows.Name = "button_clear_elbows";
            this.button_clear_elbows.Size = new System.Drawing.Size(179, 29);
            this.button_clear_elbows.TabIndex = 2482;
            this.button_clear_elbows.Text = "Remove Elbows";
            this.button_clear_elbows.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_clear_elbows.UseVisualStyleBackColor = false;
            this.button_clear_elbows.Click += new System.EventHandler(this.button_clear_elbows_Click);
            // 
            // button_load_dwg_cl
            // 
            this.button_load_dwg_cl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_load_dwg_cl.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_load_dwg_cl.BackgroundImage")));
            this.button_load_dwg_cl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_load_dwg_cl.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_load_dwg_cl.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_load_dwg_cl.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_dwg_cl.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_load_dwg_cl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_dwg_cl.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_load_dwg_cl.ForeColor = System.Drawing.Color.White;
            this.button_load_dwg_cl.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_load_dwg_cl.Location = new System.Drawing.Point(3, 3);
            this.button_load_dwg_cl.Name = "button_load_dwg_cl";
            this.button_load_dwg_cl.Size = new System.Drawing.Size(179, 27);
            this.button_load_dwg_cl.TabIndex = 2150;
            this.button_load_dwg_cl.Text = "Load DWG Centerline";
            this.button_load_dwg_cl.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_load_dwg_cl.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_load_dwg_cl.UseVisualStyleBackColor = false;
            this.button_load_dwg_cl.Click += new System.EventHandler(this.button_load_dwg_cl_Click);
            // 
            // button_load_xl_centerline
            // 
            this.button_load_xl_centerline.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_load_xl_centerline.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_load_xl_centerline.BackgroundImage")));
            this.button_load_xl_centerline.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_load_xl_centerline.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_load_xl_centerline.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_load_xl_centerline.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_xl_centerline.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_load_xl_centerline.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_xl_centerline.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_load_xl_centerline.ForeColor = System.Drawing.Color.White;
            this.button_load_xl_centerline.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_load_xl_centerline.Location = new System.Drawing.Point(3, 39);
            this.button_load_xl_centerline.Name = "button_load_xl_centerline";
            this.button_load_xl_centerline.Size = new System.Drawing.Size(179, 27);
            this.button_load_xl_centerline.TabIndex = 2149;
            this.button_load_xl_centerline.Text = "Load XL Centerline File";
            this.button_load_xl_centerline.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_load_xl_centerline.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_load_xl_centerline.UseVisualStyleBackColor = false;
            this.button_load_xl_centerline.Click += new System.EventHandler(this.button_load_xl_centerline_Click);
            // 
            // comboBox_elbow
            // 
            this.comboBox_elbow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_elbow.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_elbow.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_elbow.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.comboBox_elbow.ForeColor = System.Drawing.Color.White;
            this.comboBox_elbow.FormattingEnabled = true;
            this.comboBox_elbow.Location = new System.Drawing.Point(3, 95);
            this.comboBox_elbow.Margin = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.comboBox_elbow.Name = "comboBox_elbow";
            this.comboBox_elbow.Size = new System.Drawing.Size(179, 22);
            this.comboBox_elbow.TabIndex = 2406;
            this.comboBox_elbow.SelectedIndexChanged += new System.EventHandler(this.comboBox_elbow_SelectedIndexChanged);
            // 
            // label_default_material
            // 
            this.label_default_material.AutoSize = true;
            this.label_default_material.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_default_material.ForeColor = System.Drawing.Color.White;
            this.label_default_material.Location = new System.Drawing.Point(3, 70);
            this.label_default_material.Name = "label_default_material";
            this.label_default_material.Size = new System.Drawing.Size(87, 14);
            this.label_default_material.TabIndex = 2405;
            this.label_default_material.Text = "Elbow Material";
            this.label_default_material.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(3, 123);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 14);
            this.label1.TabIndex = 2405;
            this.label1.Text = "Description";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(3, 167);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 14);
            this.label2.TabIndex = 2405;
            this.label2.Text = "Min Angle [DD]";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // textBox_min_angle
            // 
            this.textBox_min_angle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_min_angle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_min_angle.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.textBox_min_angle.ForeColor = System.Drawing.Color.White;
            this.textBox_min_angle.Location = new System.Drawing.Point(3, 188);
            this.textBox_min_angle.Name = "textBox_min_angle";
            this.textBox_min_angle.Size = new System.Drawing.Size(179, 20);
            this.textBox_min_angle.TabIndex = 2407;
            // 
            // comboBox_description
            // 
            this.comboBox_description.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_description.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_description.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_description.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.comboBox_description.ForeColor = System.Drawing.Color.White;
            this.comboBox_description.FormattingEnabled = true;
            this.comboBox_description.Location = new System.Drawing.Point(3, 141);
            this.comboBox_description.Margin = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.comboBox_description.Name = "comboBox_description";
            this.comboBox_description.Size = new System.Drawing.Size(179, 22);
            this.comboBox_description.TabIndex = 2406;
            this.comboBox_description.SelectedIndexChanged += new System.EventHandler(this.comboBox_description_SelectedIndexChanged);
            // 
            // button_filter
            // 
            this.button_filter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_filter.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_filter.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_filter.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_filter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_filter.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_filter.ForeColor = System.Drawing.Color.White;
            this.button_filter.Image = ((System.Drawing.Image)(resources.GetObject("button_filter.Image")));
            this.button_filter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_filter.Location = new System.Drawing.Point(3, 215);
            this.button_filter.Name = "button_filter";
            this.button_filter.Size = new System.Drawing.Size(179, 28);
            this.button_filter.TabIndex = 2478;
            this.button_filter.Text = "Filter";
            this.button_filter.UseVisualStyleBackColor = false;
            this.button_filter.Click += new System.EventHandler(this.button_filter_Click);
            // 
            // button_assign_elbows
            // 
            this.button_assign_elbows.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_assign_elbows.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_assign_elbows.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_assign_elbows.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_assign_elbows.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_assign_elbows.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_assign_elbows.ForeColor = System.Drawing.Color.White;
            this.button_assign_elbows.Image = ((System.Drawing.Image)(resources.GetObject("button_assign_elbows.Image")));
            this.button_assign_elbows.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_assign_elbows.Location = new System.Drawing.Point(3, 253);
            this.button_assign_elbows.Name = "button_assign_elbows";
            this.button_assign_elbows.Size = new System.Drawing.Size(179, 32);
            this.button_assign_elbows.TabIndex = 2478;
            this.button_assign_elbows.Text = "Apply Elbow Materials";
            this.button_assign_elbows.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_assign_elbows.UseVisualStyleBackColor = false;
            this.button_assign_elbows.Click += new System.EventHandler(this.button_assign_elbows_Click);
            // 
            // button_transfer_to_mat
            // 
            this.button_transfer_to_mat.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_transfer_to_mat.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_transfer_to_mat.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_transfer_to_mat.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_transfer_to_mat.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_transfer_to_mat.Font = new System.Drawing.Font("Arial", 7F, System.Drawing.FontStyle.Bold);
            this.button_transfer_to_mat.ForeColor = System.Drawing.Color.White;
            this.button_transfer_to_mat.Image = global::Alignment_mdi.Properties.Resources.Target;
            this.button_transfer_to_mat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_transfer_to_mat.Location = new System.Drawing.Point(3, 439);
            this.button_transfer_to_mat.Name = "button_transfer_to_mat";
            this.button_transfer_to_mat.Size = new System.Drawing.Size(179, 30);
            this.button_transfer_to_mat.TabIndex = 2481;
            this.button_transfer_to_mat.Text = "Elbows to Material Design";
            this.button_transfer_to_mat.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_transfer_to_mat.UseVisualStyleBackColor = false;
            this.button_transfer_to_mat.Click += new System.EventHandler(this.button_transfer_to_mat_Click);
            // 
            // panel8
            // 
            this.panel8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel8.Controls.Add(this.label_header);
            this.panel8.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel8.Location = new System.Drawing.Point(0, 0);
            this.panel8.Margin = new System.Windows.Forms.Padding(0);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(995, 27);
            this.panel8.TabIndex = 2144;
            // 
            // label_header
            // 
            this.label_header.AutoSize = true;
            this.label_header.BackColor = System.Drawing.Color.Transparent;
            this.label_header.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_header.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label_header.Location = new System.Drawing.Point(1, 3);
            this.label_header.Margin = new System.Windows.Forms.Padding(0);
            this.label_header.Name = "label_header";
            this.label_header.Size = new System.Drawing.Size(119, 18);
            this.label_header.TabIndex = 2055;
            this.label_header.Text = "Route Designer";
            this.label_header.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.tableLayoutPanel2.ColumnCount = 5;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 174F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(800, 27);
            this.tableLayoutPanel2.TabIndex = 2216;
            // 
            // button_add_stationing
            // 
            this.button_add_stationing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_add_stationing.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_add_stationing.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_add_stationing.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_add_stationing.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_add_stationing.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_add_stationing.ForeColor = System.Drawing.Color.White;
            this.button_add_stationing.Image = ((System.Drawing.Image)(resources.GetObject("button_add_stationing.Image")));
            this.button_add_stationing.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_add_stationing.Location = new System.Drawing.Point(3, 328);
            this.button_add_stationing.Name = "button_add_stationing";
            this.button_add_stationing.Size = new System.Drawing.Size(179, 29);
            this.button_add_stationing.TabIndex = 2482;
            this.button_add_stationing.Text = "Draw Stationing";
            this.button_add_stationing.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_add_stationing.UseVisualStyleBackColor = false;
            this.button_add_stationing.Click += new System.EventHandler(this.button_add_stationing_Click);
            // 
            // Centerline_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(997, 502);
            this.Controls.Add(this.panel_Matl_Tool);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Centerline_form";
            this.Text = "Form1";
            this.panel_Matl_Tool.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_cl)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel_bottom.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel8.ResumeLayout(false);
            this.panel8.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_Matl_Tool;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Button button_load_xl_centerline;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Button button_load_dwg_cl;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView dataGridView_cl;
        private System.Windows.Forms.Panel panel_bottom;
        private System.Windows.Forms.Button button_save_cl;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label_cl;
        private System.Windows.Forms.Label label_default_material;
        private System.Windows.Forms.ComboBox comboBox_elbow;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_assign_elbows;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_min_angle;
        private System.Windows.Forms.ComboBox comboBox_description;
        private System.Windows.Forms.Button button_filter;
        private System.Windows.Forms.Button button_transfer_to_mat;
        private System.Windows.Forms.Label label_header;
        private System.Windows.Forms.Button button_clear_elbows;
        private System.Windows.Forms.Button button_add_stationing;
    }
}




