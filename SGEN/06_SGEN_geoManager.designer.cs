namespace Alignment_mdi
{
    partial class Geo_tools_form
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
            this.panel6 = new System.Windows.Forms.Panel();
            this.label7 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.radioButton_BLOCKS = new System.Windows.Forms.RadioButton();
            this.radioButton_OD = new System.Windows.Forms.RadioButton();
            this.label_drawing_name = new System.Windows.Forms.Label();
            this.label_correct_od_table = new System.Windows.Forms.Label();
            this.button_refresh_grid = new System.Windows.Forms.Button();
            this.button_refresh_layer_tables = new System.Windows.Forms.Button();
            this.label_current_layer_block = new System.Windows.Forms.Label();
            this.comboBox_layers_blocks_geomanager = new System.Windows.Forms.ComboBox();
            this.comboBox_od_existing_tables = new System.Windows.Forms.ComboBox();
            this.button_export_to_excel = new System.Windows.Forms.Button();
            this.button_import_from_excel = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.Button_Update_object_data = new System.Windows.Forms.Button();
            this.panel_navigation = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.button_multiselect = new System.Windows.Forms.Button();
            this.button_zoom_row_object_data = new System.Windows.Forms.Button();
            this.button_zoom = new System.Windows.Forms.Button();
            this.panel_stats = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button_Filter = new System.Windows.Forms.Button();
            this.button_add_OD_table = new System.Windows.Forms.Button();
            this.textBox_Features = new System.Windows.Forms.TextBox();
            this.textBox_MultipleOD = new System.Windows.Forms.TextBox();
            this.textBox_no_wrong_od = new System.Windows.Forms.TextBox();
            this.textBox_no_rows = new System.Windows.Forms.TextBox();
            this.textBox_no_tables = new System.Windows.Forms.TextBox();
            this.textBox_no_od_zero = new System.Windows.Forms.TextBox();
            this.textBox_no_od_2 = new System.Windows.Forms.TextBox();
            this.textBox_INCORRECT_od = new System.Windows.Forms.TextBox();
            this.textBox_missing_OD = new System.Windows.Forms.TextBox();
            this.textBox_OD_TABLES = new System.Windows.Forms.TextBox();
            this.label_processing1 = new System.Windows.Forms.Label();
            this.DataGridView_data = new System.Windows.Forms.DataGridView();
            this.radioButton_mtext = new System.Windows.Forms.RadioButton();
            this.panel6.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel_navigation.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel_stats.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView_data)).BeginInit();
            this.SuspendLayout();
            // 
            // panel6
            // 
            this.panel6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel6.Controls.Add(this.label7);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel6.Location = new System.Drawing.Point(0, 0);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(970, 25);
            this.panel6.TabIndex = 2152;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label7.Location = new System.Drawing.Point(430, 2);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(109, 18);
            this.label7.TabIndex = 2054;
            this.label7.Text = "Data Manager";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.radioButton_mtext);
            this.panel1.Controls.Add(this.radioButton_BLOCKS);
            this.panel1.Controls.Add(this.radioButton_OD);
            this.panel1.Controls.Add(this.label_drawing_name);
            this.panel1.Controls.Add(this.label_correct_od_table);
            this.panel1.Controls.Add(this.button_refresh_grid);
            this.panel1.Controls.Add(this.button_refresh_layer_tables);
            this.panel1.Controls.Add(this.label_current_layer_block);
            this.panel1.Controls.Add(this.comboBox_layers_blocks_geomanager);
            this.panel1.Controls.Add(this.comboBox_od_existing_tables);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(970, 94);
            this.panel1.TabIndex = 2153;
            // 
            // radioButton_BLOCKS
            // 
            this.radioButton_BLOCKS.AutoSize = true;
            this.radioButton_BLOCKS.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_BLOCKS.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.radioButton_BLOCKS.ForeColor = System.Drawing.Color.White;
            this.radioButton_BLOCKS.Location = new System.Drawing.Point(3, 34);
            this.radioButton_BLOCKS.Name = "radioButton_BLOCKS";
            this.radioButton_BLOCKS.Size = new System.Drawing.Size(63, 19);
            this.radioButton_BLOCKS.TabIndex = 1;
            this.radioButton_BLOCKS.TabStop = true;
            this.radioButton_BLOCKS.Text = "Blocks";
            this.radioButton_BLOCKS.UseVisualStyleBackColor = true;
            this.radioButton_BLOCKS.CheckedChanged += new System.EventHandler(this.radioButton_OD_blocks_CheckedChanged);
            // 
            // radioButton_OD
            // 
            this.radioButton_OD.AutoSize = true;
            this.radioButton_OD.Checked = true;
            this.radioButton_OD.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_OD.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.radioButton_OD.ForeColor = System.Drawing.Color.White;
            this.radioButton_OD.Location = new System.Drawing.Point(3, 9);
            this.radioButton_OD.Name = "radioButton_OD";
            this.radioButton_OD.Size = new System.Drawing.Size(81, 19);
            this.radioButton_OD.TabIndex = 0;
            this.radioButton_OD.TabStop = true;
            this.radioButton_OD.Text = "OD Tables";
            this.radioButton_OD.UseVisualStyleBackColor = true;
            this.radioButton_OD.CheckedChanged += new System.EventHandler(this.radioButton_OD_blocks_CheckedChanged);
            // 
            // label_drawing_name
            // 
            this.label_drawing_name.AutoSize = true;
            this.label_drawing_name.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_drawing_name.ForeColor = System.Drawing.Color.White;
            this.label_drawing_name.Location = new System.Drawing.Point(433, 2);
            this.label_drawing_name.Name = "label_drawing_name";
            this.label_drawing_name.Size = new System.Drawing.Size(76, 16);
            this.label_drawing_name.TabIndex = 42;
            this.label_drawing_name.Text = "Drawing Name";
            this.label_drawing_name.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label_correct_od_table
            // 
            this.label_correct_od_table.AutoSize = true;
            this.label_correct_od_table.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.label_correct_od_table.ForeColor = System.Drawing.Color.White;
            this.label_correct_od_table.Location = new System.Drawing.Point(392, 42);
            this.label_correct_od_table.Name = "label_correct_od_table";
            this.label_correct_od_table.Size = new System.Drawing.Size(103, 15);
            this.label_correct_od_table.TabIndex = 43;
            this.label_correct_od_table.Text = "Correct OD Table";
            // 
            // button_refresh_grid
            // 
            this.button_refresh_grid.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_refresh_grid.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_refresh_grid.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_refresh_grid.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_refresh_grid.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_refresh_grid.ForeColor = System.Drawing.Color.White;
            this.button_refresh_grid.Location = new System.Drawing.Point(679, 59);
            this.button_refresh_grid.Name = "button_refresh_grid";
            this.button_refresh_grid.Size = new System.Drawing.Size(150, 24);
            this.button_refresh_grid.TabIndex = 41;
            this.button_refresh_grid.Text = "Build Table";
            this.button_refresh_grid.UseVisualStyleBackColor = false;
            this.button_refresh_grid.Click += new System.EventHandler(this.button_LOAD_DATA_Click);
            // 
            // button_refresh_layer_tables
            // 
            this.button_refresh_layer_tables.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_refresh_layer_tables.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_refresh_layer_tables.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_refresh_layer_tables.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_refresh_layer_tables.ForeColor = System.Drawing.Color.White;
            this.button_refresh_layer_tables.Location = new System.Drawing.Point(3, 59);
            this.button_refresh_layer_tables.Name = "button_refresh_layer_tables";
            this.button_refresh_layer_tables.Size = new System.Drawing.Size(100, 24);
            this.button_refresh_layer_tables.TabIndex = 39;
            this.button_refresh_layer_tables.Text = "Load Data";
            this.button_refresh_layer_tables.UseVisualStyleBackColor = false;
            this.button_refresh_layer_tables.Click += new System.EventHandler(this.button_load_layers_and_data_tables_Click);
            // 
            // label_current_layer_block
            // 
            this.label_current_layer_block.AutoSize = true;
            this.label_current_layer_block.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.label_current_layer_block.ForeColor = System.Drawing.Color.White;
            this.label_current_layer_block.Location = new System.Drawing.Point(112, 44);
            this.label_current_layer_block.Name = "label_current_layer_block";
            this.label_current_layer_block.Size = new System.Drawing.Size(85, 15);
            this.label_current_layer_block.TabIndex = 40;
            this.label_current_layer_block.Text = "Current Layer";
            // 
            // comboBox_layers_blocks_geomanager
            // 
            this.comboBox_layers_blocks_geomanager.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_layers_blocks_geomanager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_layers_blocks_geomanager.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_layers_blocks_geomanager.ForeColor = System.Drawing.Color.White;
            this.comboBox_layers_blocks_geomanager.FormattingEnabled = true;
            this.comboBox_layers_blocks_geomanager.Location = new System.Drawing.Point(112, 62);
            this.comboBox_layers_blocks_geomanager.Name = "comboBox_layers_blocks_geomanager";
            this.comboBox_layers_blocks_geomanager.Size = new System.Drawing.Size(272, 21);
            this.comboBox_layers_blocks_geomanager.TabIndex = 37;
            // 
            // comboBox_od_existing_tables
            // 
            this.comboBox_od_existing_tables.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_od_existing_tables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_od_existing_tables.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox_od_existing_tables.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.comboBox_od_existing_tables.ForeColor = System.Drawing.Color.White;
            this.comboBox_od_existing_tables.FormattingEnabled = true;
            this.comboBox_od_existing_tables.Location = new System.Drawing.Point(392, 60);
            this.comboBox_od_existing_tables.Name = "comboBox_od_existing_tables";
            this.comboBox_od_existing_tables.Size = new System.Drawing.Size(272, 24);
            this.comboBox_od_existing_tables.TabIndex = 38;
            // 
            // button_export_to_excel
            // 
            this.button_export_to_excel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_export_to_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_export_to_excel.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_export_to_excel.ForeColor = System.Drawing.Color.White;
            this.button_export_to_excel.Location = new System.Drawing.Point(682, 3);
            this.button_export_to_excel.Name = "button_export_to_excel";
            this.button_export_to_excel.Size = new System.Drawing.Size(148, 31);
            this.button_export_to_excel.TabIndex = 13;
            this.button_export_to_excel.Text = "Export to Excel";
            this.button_export_to_excel.UseVisualStyleBackColor = true;
            this.button_export_to_excel.Click += new System.EventHandler(this.button_export_to_excel_Click);
            // 
            // button_import_from_excel
            // 
            this.button_import_from_excel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_import_from_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_import_from_excel.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_import_from_excel.ForeColor = System.Drawing.Color.White;
            this.button_import_from_excel.Location = new System.Drawing.Point(682, 39);
            this.button_import_from_excel.Name = "button_import_from_excel";
            this.button_import_from_excel.Size = new System.Drawing.Size(148, 31);
            this.button_import_from_excel.TabIndex = 14;
            this.button_import_from_excel.Text = "Import from Excel";
            this.button_import_from_excel.UseVisualStyleBackColor = true;
            this.button_import_from_excel.Click += new System.EventHandler(this.button_import_from_excel_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel2.Controls.Add(this.Button_Update_object_data);
            this.panel2.Controls.Add(this.panel_navigation);
            this.panel2.Controls.Add(this.button_import_from_excel);
            this.panel2.Controls.Add(this.button_export_to_excel);
            this.panel2.Controls.Add(this.panel_stats);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 312);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(970, 190);
            this.panel2.TabIndex = 2154;
            // 
            // Button_Update_object_data
            // 
            this.Button_Update_object_data.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.Button_Update_object_data.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.Button_Update_object_data.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_Update_object_data.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.Button_Update_object_data.ForeColor = System.Drawing.Color.White;
            this.Button_Update_object_data.Location = new System.Drawing.Point(680, 147);
            this.Button_Update_object_data.Name = "Button_Update_object_data";
            this.Button_Update_object_data.Size = new System.Drawing.Size(148, 31);
            this.Button_Update_object_data.TabIndex = 24;
            this.Button_Update_object_data.Text = "Update Drawing";
            this.Button_Update_object_data.UseVisualStyleBackColor = false;
            this.Button_Update_object_data.Click += new System.EventHandler(this.Button_Update_data_Click);
            // 
            // panel_navigation
            // 
            this.panel_navigation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_navigation.Controls.Add(this.panel4);
            this.panel_navigation.Controls.Add(this.button_multiselect);
            this.panel_navigation.Controls.Add(this.button_zoom_row_object_data);
            this.panel_navigation.Controls.Add(this.button_zoom);
            this.panel_navigation.Location = new System.Drawing.Point(351, 0);
            this.panel_navigation.Name = "panel_navigation";
            this.panel_navigation.Size = new System.Drawing.Size(171, 190);
            this.panel_navigation.TabIndex = 23;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.label3);
            this.panel4.Controls.Add(this.label4);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(169, 25);
            this.panel4.TabIndex = 2154;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label3.Location = new System.Drawing.Point(19, 2);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(128, 18);
            this.label3.TabIndex = 2055;
            this.label3.Text = "Navigation Tools";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label4.Location = new System.Drawing.Point(430, 2);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(109, 18);
            this.label4.TabIndex = 2054;
            this.label4.Text = "Data Manager";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button_multiselect
            // 
            this.button_multiselect.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_multiselect.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_multiselect.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_multiselect.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_multiselect.ForeColor = System.Drawing.Color.White;
            this.button_multiselect.Location = new System.Drawing.Point(10, 116);
            this.button_multiselect.Name = "button_multiselect";
            this.button_multiselect.Size = new System.Drawing.Size(148, 31);
            this.button_multiselect.TabIndex = 5;
            this.button_multiselect.Text = "Multiselect";
            this.button_multiselect.UseVisualStyleBackColor = false;
            this.button_multiselect.Click += new System.EventHandler(this.button_multiselect_Click);
            // 
            // button_zoom_row_object_data
            // 
            this.button_zoom_row_object_data.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_zoom_row_object_data.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_zoom_row_object_data.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_zoom_row_object_data.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_zoom_row_object_data.ForeColor = System.Drawing.Color.White;
            this.button_zoom_row_object_data.Location = new System.Drawing.Point(10, 42);
            this.button_zoom_row_object_data.Name = "button_zoom_row_object_data";
            this.button_zoom_row_object_data.Size = new System.Drawing.Size(148, 31);
            this.button_zoom_row_object_data.TabIndex = 5;
            this.button_zoom_row_object_data.Text = "Select Feature";
            this.button_zoom_row_object_data.UseVisualStyleBackColor = false;
            this.button_zoom_row_object_data.Click += new System.EventHandler(this.button_go_to_table_row_Click);
            // 
            // button_zoom
            // 
            this.button_zoom.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_zoom.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_zoom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_zoom.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_zoom.ForeColor = System.Drawing.Color.White;
            this.button_zoom.Location = new System.Drawing.Point(10, 79);
            this.button_zoom.Name = "button_zoom";
            this.button_zoom.Size = new System.Drawing.Size(148, 31);
            this.button_zoom.TabIndex = 6;
            this.button_zoom.Text = "Zoom To Feature";
            this.button_zoom.UseVisualStyleBackColor = false;
            this.button_zoom.Click += new System.EventHandler(this.button_zoom_Click);
            // 
            // panel_stats
            // 
            this.panel_stats.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel_stats.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_stats.Controls.Add(this.panel3);
            this.panel_stats.Controls.Add(this.button_Filter);
            this.panel_stats.Controls.Add(this.button_add_OD_table);
            this.panel_stats.Controls.Add(this.textBox_Features);
            this.panel_stats.Controls.Add(this.textBox_MultipleOD);
            this.panel_stats.Controls.Add(this.textBox_no_wrong_od);
            this.panel_stats.Controls.Add(this.textBox_no_rows);
            this.panel_stats.Controls.Add(this.textBox_no_tables);
            this.panel_stats.Controls.Add(this.textBox_no_od_zero);
            this.panel_stats.Controls.Add(this.textBox_no_od_2);
            this.panel_stats.Controls.Add(this.textBox_INCORRECT_od);
            this.panel_stats.Controls.Add(this.textBox_missing_OD);
            this.panel_stats.Controls.Add(this.textBox_OD_TABLES);
            this.panel_stats.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel_stats.Location = new System.Drawing.Point(0, 0);
            this.panel_stats.Name = "panel_stats";
            this.panel_stats.Size = new System.Drawing.Size(345, 190);
            this.panel_stats.TabIndex = 20;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(343, 25);
            this.panel3.TabIndex = 2153;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label2.Location = new System.Drawing.Point(132, 2);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 18);
            this.label2.TabIndex = 2055;
            this.label2.Text = "Statistics";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label1.Location = new System.Drawing.Point(430, 2);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 18);
            this.label1.TabIndex = 2054;
            this.label1.Text = "Data Manager";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button_Filter
            // 
            this.button_Filter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_Filter.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_Filter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Filter.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.button_Filter.ForeColor = System.Drawing.Color.White;
            this.button_Filter.Location = new System.Drawing.Point(158, 147);
            this.button_Filter.Name = "button_Filter";
            this.button_Filter.Size = new System.Drawing.Size(138, 30);
            this.button_Filter.TabIndex = 22;
            this.button_Filter.Text = "Filter to Issues";
            this.button_Filter.UseVisualStyleBackColor = false;
            this.button_Filter.Click += new System.EventHandler(this.button_add_OD_table_and_remove_wrong_OD_Click);
            // 
            // button_add_OD_table
            // 
            this.button_add_OD_table.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.button_add_OD_table.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_add_OD_table.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_add_OD_table.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.button_add_OD_table.ForeColor = System.Drawing.Color.White;
            this.button_add_OD_table.Location = new System.Drawing.Point(9, 147);
            this.button_add_OD_table.Name = "button_add_OD_table";
            this.button_add_OD_table.Size = new System.Drawing.Size(138, 30);
            this.button_add_OD_table.TabIndex = 21;
            this.button_add_OD_table.Text = "Fix Issues";
            this.button_add_OD_table.UseVisualStyleBackColor = false;
            this.button_add_OD_table.Click += new System.EventHandler(this.button_Filter_Click);
            // 
            // textBox_Features
            // 
            this.textBox_Features.BackColor = System.Drawing.Color.White;
            this.textBox_Features.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_Features.ForeColor = System.Drawing.Color.Black;
            this.textBox_Features.Location = new System.Drawing.Point(9, 38);
            this.textBox_Features.Name = "textBox_Features";
            this.textBox_Features.ReadOnly = true;
            this.textBox_Features.Size = new System.Drawing.Size(287, 21);
            this.textBox_Features.TabIndex = 17;
            this.textBox_Features.Text = "Features: ";
            // 
            // textBox_MultipleOD
            // 
            this.textBox_MultipleOD.BackColor = System.Drawing.Color.Red;
            this.textBox_MultipleOD.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_MultipleOD.ForeColor = System.Drawing.Color.Black;
            this.textBox_MultipleOD.Location = new System.Drawing.Point(9, 80);
            this.textBox_MultipleOD.Name = "textBox_MultipleOD";
            this.textBox_MultipleOD.ReadOnly = true;
            this.textBox_MultipleOD.Size = new System.Drawing.Size(287, 21);
            this.textBox_MultipleOD.TabIndex = 17;
            this.textBox_MultipleOD.Text = "Features with Multiple OD Tables:";
            // 
            // textBox_no_wrong_od
            // 
            this.textBox_no_wrong_od.BackColor = System.Drawing.Color.SkyBlue;
            this.textBox_no_wrong_od.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_wrong_od.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_wrong_od.Location = new System.Drawing.Point(302, 101);
            this.textBox_no_wrong_od.Name = "textBox_no_wrong_od";
            this.textBox_no_wrong_od.ReadOnly = true;
            this.textBox_no_wrong_od.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_wrong_od.TabIndex = 17;
            this.textBox_no_wrong_od.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_rows
            // 
            this.textBox_no_rows.BackColor = System.Drawing.Color.White;
            this.textBox_no_rows.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_rows.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_rows.Location = new System.Drawing.Point(302, 38);
            this.textBox_no_rows.Name = "textBox_no_rows";
            this.textBox_no_rows.ReadOnly = true;
            this.textBox_no_rows.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_rows.TabIndex = 17;
            this.textBox_no_rows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_tables
            // 
            this.textBox_no_tables.BackColor = System.Drawing.Color.White;
            this.textBox_no_tables.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_tables.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_tables.Location = new System.Drawing.Point(302, 122);
            this.textBox_no_tables.Name = "textBox_no_tables";
            this.textBox_no_tables.ReadOnly = true;
            this.textBox_no_tables.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_tables.TabIndex = 17;
            this.textBox_no_tables.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_od_zero
            // 
            this.textBox_no_od_zero.BackColor = System.Drawing.Color.Yellow;
            this.textBox_no_od_zero.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_od_zero.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_od_zero.Location = new System.Drawing.Point(302, 59);
            this.textBox_no_od_zero.Name = "textBox_no_od_zero";
            this.textBox_no_od_zero.ReadOnly = true;
            this.textBox_no_od_zero.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_od_zero.TabIndex = 17;
            this.textBox_no_od_zero.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_od_2
            // 
            this.textBox_no_od_2.BackColor = System.Drawing.Color.Red;
            this.textBox_no_od_2.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_od_2.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_od_2.Location = new System.Drawing.Point(302, 80);
            this.textBox_no_od_2.Name = "textBox_no_od_2";
            this.textBox_no_od_2.ReadOnly = true;
            this.textBox_no_od_2.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_od_2.TabIndex = 17;
            this.textBox_no_od_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_INCORRECT_od
            // 
            this.textBox_INCORRECT_od.BackColor = System.Drawing.Color.SkyBlue;
            this.textBox_INCORRECT_od.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_INCORRECT_od.ForeColor = System.Drawing.Color.Black;
            this.textBox_INCORRECT_od.Location = new System.Drawing.Point(9, 101);
            this.textBox_INCORRECT_od.Name = "textBox_INCORRECT_od";
            this.textBox_INCORRECT_od.ReadOnly = true;
            this.textBox_INCORRECT_od.Size = new System.Drawing.Size(287, 21);
            this.textBox_INCORRECT_od.TabIndex = 17;
            this.textBox_INCORRECT_od.Text = "Features with Incorrect OD Tables:";
            // 
            // textBox_missing_OD
            // 
            this.textBox_missing_OD.BackColor = System.Drawing.Color.Yellow;
            this.textBox_missing_OD.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_missing_OD.ForeColor = System.Drawing.Color.Black;
            this.textBox_missing_OD.Location = new System.Drawing.Point(9, 59);
            this.textBox_missing_OD.Name = "textBox_missing_OD";
            this.textBox_missing_OD.ReadOnly = true;
            this.textBox_missing_OD.Size = new System.Drawing.Size(287, 21);
            this.textBox_missing_OD.TabIndex = 17;
            this.textBox_missing_OD.Text = "Features with Missing OD Tables:";
            // 
            // textBox_OD_TABLES
            // 
            this.textBox_OD_TABLES.BackColor = System.Drawing.Color.White;
            this.textBox_OD_TABLES.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_OD_TABLES.ForeColor = System.Drawing.Color.Black;
            this.textBox_OD_TABLES.Location = new System.Drawing.Point(9, 122);
            this.textBox_OD_TABLES.Name = "textBox_OD_TABLES";
            this.textBox_OD_TABLES.ReadOnly = true;
            this.textBox_OD_TABLES.Size = new System.Drawing.Size(287, 21);
            this.textBox_OD_TABLES.TabIndex = 17;
            this.textBox_OD_TABLES.Text = "Total Number of OD Tables on Layer:";
            // 
            // label_processing1
            // 
            this.label_processing1.AutoSize = true;
            this.label_processing1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label_processing1.Font = new System.Drawing.Font("Arial", 15F, System.Drawing.FontStyle.Bold);
            this.label_processing1.ForeColor = System.Drawing.Color.White;
            this.label_processing1.Location = new System.Drawing.Point(379, 196);
            this.label_processing1.Name = "label_processing1";
            this.label_processing1.Size = new System.Drawing.Size(142, 26);
            this.label_processing1.TabIndex = 2156;
            this.label_processing1.Text = "Processing....";
            // 
            // DataGridView_data
            // 
            this.DataGridView_data.AllowUserToAddRows = false;
            this.DataGridView_data.AllowUserToDeleteRows = false;
            this.DataGridView_data.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.DataGridView_data.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.DataGridView_data.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.DataGridView_data.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.DataGridView_data.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.DataGridView_data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DataGridView_data.DefaultCellStyle = dataGridViewCellStyle1;
            this.DataGridView_data.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DataGridView_data.GridColor = System.Drawing.Color.LightGray;
            this.DataGridView_data.Location = new System.Drawing.Point(0, 119);
            this.DataGridView_data.Name = "DataGridView_data";
            this.DataGridView_data.RowHeadersVisible = false;
            this.DataGridView_data.Size = new System.Drawing.Size(970, 193);
            this.DataGridView_data.TabIndex = 2155;
            this.DataGridView_data.VirtualMode = true;
            this.DataGridView_data.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGrid_od_data_CellClick);
            this.DataGridView_data.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView_data_CellValueChanged);
            this.DataGridView_data.Sorted += new System.EventHandler(this.DataGridView_OD_data_Sorted);
            // 
            // radioButton_mtext
            // 
            this.radioButton_mtext.AutoSize = true;
            this.radioButton_mtext.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.radioButton_mtext.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.radioButton_mtext.ForeColor = System.Drawing.Color.White;
            this.radioButton_mtext.Location = new System.Drawing.Point(112, 9);
            this.radioButton_mtext.Name = "radioButton_mtext";
            this.radioButton_mtext.Size = new System.Drawing.Size(56, 19);
            this.radioButton_mtext.TabIndex = 1;
            this.radioButton_mtext.TabStop = true;
            this.radioButton_mtext.Text = "Mtext";
            this.radioButton_mtext.UseVisualStyleBackColor = true;
            this.radioButton_mtext.CheckedChanged += new System.EventHandler(this.radioButton_OD_blocks_CheckedChanged);
            // 
            // Geo_tools_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(970, 502);
            this.Controls.Add(this.label_processing1);
            this.Controls.Add(this.DataGridView_data);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel6);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Geo_tools_form";
            this.Text = "Geo Manager";
            this.Load += new System.EventHandler(this.OD_TABLE_form_Load);
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel_navigation.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel_stats.ResumeLayout(false);
            this.panel_stats.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView_data)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton radioButton_BLOCKS;
        private System.Windows.Forms.RadioButton radioButton_OD;
        private System.Windows.Forms.Label label_drawing_name;
        private System.Windows.Forms.Label label_correct_od_table;
        private System.Windows.Forms.Button button_refresh_grid;
        private System.Windows.Forms.Button button_refresh_layer_tables;
        private System.Windows.Forms.Label label_current_layer_block;
        private System.Windows.Forms.ComboBox comboBox_layers_blocks_geomanager;
        private System.Windows.Forms.ComboBox comboBox_od_existing_tables;
        private System.Windows.Forms.Button button_import_from_excel;
        private System.Windows.Forms.Button button_export_to_excel;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel_stats;
        private System.Windows.Forms.TextBox textBox_Features;
        private System.Windows.Forms.TextBox textBox_MultipleOD;
        private System.Windows.Forms.TextBox textBox_no_wrong_od;
        private System.Windows.Forms.TextBox textBox_no_rows;
        private System.Windows.Forms.TextBox textBox_no_tables;
        private System.Windows.Forms.TextBox textBox_no_od_zero;
        private System.Windows.Forms.TextBox textBox_no_od_2;
        private System.Windows.Forms.TextBox textBox_INCORRECT_od;
        private System.Windows.Forms.TextBox textBox_missing_OD;
        private System.Windows.Forms.TextBox textBox_OD_TABLES;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_Filter;
        private System.Windows.Forms.Button button_add_OD_table;
        private System.Windows.Forms.Panel panel_navigation;
        private System.Windows.Forms.Button button_multiselect;
        private System.Windows.Forms.Button button_zoom_row_object_data;
        private System.Windows.Forms.Button button_zoom;
        private System.Windows.Forms.Button Button_Update_object_data;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label_processing1;
        private System.Windows.Forms.DataGridView DataGridView_data;
        private System.Windows.Forms.RadioButton radioButton_mtext;
    }
}