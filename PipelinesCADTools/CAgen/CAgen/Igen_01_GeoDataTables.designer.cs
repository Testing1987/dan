namespace Alignment_mdi
{
    partial class Igen_geomanager
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Igen_geomanager));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.Button_Update_object_data = new System.Windows.Forms.Button();
            this.label_dt_issues = new System.Windows.Forms.Label();
            this.panel_stats = new System.Windows.Forms.Panel();
            this.textBox_Features = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_MultipleOD = new System.Windows.Forms.TextBox();
            this.textBox_no_wrong_od = new System.Windows.Forms.TextBox();
            this.textBox_no_rows = new System.Windows.Forms.TextBox();
            this.textBox_no_tables = new System.Windows.Forms.TextBox();
            this.textBox_no_od_zero = new System.Windows.Forms.TextBox();
            this.textBox_no_od_2 = new System.Windows.Forms.TextBox();
            this.textBox_INCORRECT_od = new System.Windows.Forms.TextBox();
            this.textBox_missing_OD = new System.Windows.Forms.TextBox();
            this.textBox_OD_TABLES = new System.Windows.Forms.TextBox();
            this.button_Filter = new System.Windows.Forms.Button();
            this.button_add_OD_table = new System.Windows.Forms.Button();
            this.panel_blocks_and_OD = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.radioButton_OD = new System.Windows.Forms.RadioButton();
            this.radioButton_BLOCKS = new System.Windows.Forms.RadioButton();
            this.button_refresh_layer_tables = new System.Windows.Forms.Button();
            this.panel_logo = new System.Windows.Forms.Panel();
            this.label_Apply_Changes = new System.Windows.Forms.Label();
            this.panel_excel = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.button_export_to_excel = new System.Windows.Forms.Button();
            this.button_import_from_excel = new System.Windows.Forms.Button();
            this.label_drawing_name = new System.Windows.Forms.Label();
            this.label_correct_od_table = new System.Windows.Forms.Label();
            this.button_refresh_grid = new System.Windows.Forms.Button();
            this.label_od_block_table = new System.Windows.Forms.Label();
            this.label_current_layer_block = new System.Windows.Forms.Label();
            this.comboBox_layers_blocks_geomanager = new System.Windows.Forms.ComboBox();
            this.comboBox_od_existing_tables = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label_layer_rules = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel_navigation = new System.Windows.Forms.Panel();
            this.label10 = new System.Windows.Forms.Label();
            this.button_zoom_row_object_data = new System.Windows.Forms.Button();
            this.button_zoom = new System.Windows.Forms.Button();
            this.panel_grid = new System.Windows.Forms.Panel();
            this.label_processing1 = new System.Windows.Forms.Label();
            this.DataGridView_data = new System.Windows.Forms.DataGridView();
            this.panel_stats.SuspendLayout();
            this.panel_blocks_and_OD.SuspendLayout();
            this.panel_excel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel_navigation.SuspendLayout();
            this.panel_grid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView_data)).BeginInit();
            this.SuspendLayout();
            // 
            // Button_Update_object_data
            // 
            this.Button_Update_object_data.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.Button_Update_object_data.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.Button_Update_object_data.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Button_Update_object_data.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.Button_Update_object_data.ForeColor = System.Drawing.Color.White;
            this.Button_Update_object_data.Location = new System.Drawing.Point(656, 585);
            this.Button_Update_object_data.Name = "Button_Update_object_data";
            this.Button_Update_object_data.Size = new System.Drawing.Size(129, 25);
            this.Button_Update_object_data.TabIndex = 2;
            this.Button_Update_object_data.Text = "Apply Changes";
            this.Button_Update_object_data.UseVisualStyleBackColor = false;
            this.Button_Update_object_data.Click += new System.EventHandler(this.Button_Update_data_Click);
            // 
            // label_dt_issues
            // 
            this.label_dt_issues.AutoSize = true;
            this.label_dt_issues.BackColor = System.Drawing.Color.Transparent;
            this.label_dt_issues.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_dt_issues.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label_dt_issues.Location = new System.Drawing.Point(7, 396);
            this.label_dt_issues.Name = "label_dt_issues";
            this.label_dt_issues.Size = new System.Drawing.Size(136, 18);
            this.label_dt_issues.TabIndex = 2063;
            this.label_dt_issues.Text = "Data Table Issues";
            // 
            // panel_stats
            // 
            this.panel_stats.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(35)))));
            this.panel_stats.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_stats.Controls.Add(this.textBox_Features);
            this.panel_stats.Controls.Add(this.label6);
            this.panel_stats.Controls.Add(this.textBox_MultipleOD);
            this.panel_stats.Controls.Add(this.textBox_no_wrong_od);
            this.panel_stats.Controls.Add(this.textBox_no_rows);
            this.panel_stats.Controls.Add(this.textBox_no_tables);
            this.panel_stats.Controls.Add(this.textBox_no_od_zero);
            this.panel_stats.Controls.Add(this.textBox_no_od_2);
            this.panel_stats.Controls.Add(this.textBox_INCORRECT_od);
            this.panel_stats.Controls.Add(this.textBox_missing_OD);
            this.panel_stats.Controls.Add(this.textBox_OD_TABLES);
            this.panel_stats.Location = new System.Drawing.Point(3, 472);
            this.panel_stats.Name = "panel_stats";
            this.panel_stats.Size = new System.Drawing.Size(355, 134);
            this.panel_stats.TabIndex = 19;
            // 
            // textBox_Features
            // 
            this.textBox_Features.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_Features.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_Features.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.textBox_Features.ForeColor = System.Drawing.Color.White;
            this.textBox_Features.Location = new System.Drawing.Point(3, 21);
            this.textBox_Features.Name = "textBox_Features";
            this.textBox_Features.ReadOnly = true;
            this.textBox_Features.Size = new System.Drawing.Size(287, 20);
            this.textBox_Features.TabIndex = 17;
            this.textBox_Features.Text = "Features: ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(5, 3);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 14);
            this.label6.TabIndex = 11;
            this.label6.Text = "Statistics";
            // 
            // textBox_MultipleOD
            // 
            this.textBox_MultipleOD.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_MultipleOD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_MultipleOD.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.textBox_MultipleOD.ForeColor = System.Drawing.Color.Red;
            this.textBox_MultipleOD.Location = new System.Drawing.Point(3, 63);
            this.textBox_MultipleOD.Name = "textBox_MultipleOD";
            this.textBox_MultipleOD.ReadOnly = true;
            this.textBox_MultipleOD.Size = new System.Drawing.Size(287, 20);
            this.textBox_MultipleOD.TabIndex = 17;
            this.textBox_MultipleOD.Text = "Features with Multiple OD Tables:";
            // 
            // textBox_no_wrong_od
            // 
            this.textBox_no_wrong_od.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_no_wrong_od.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_no_wrong_od.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_wrong_od.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_wrong_od.Location = new System.Drawing.Point(311, 84);
            this.textBox_no_wrong_od.Name = "textBox_no_wrong_od";
            this.textBox_no_wrong_od.ReadOnly = true;
            this.textBox_no_wrong_od.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_wrong_od.TabIndex = 17;
            this.textBox_no_wrong_od.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_rows
            // 
            this.textBox_no_rows.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_no_rows.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_no_rows.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_rows.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_rows.Location = new System.Drawing.Point(311, 21);
            this.textBox_no_rows.Name = "textBox_no_rows";
            this.textBox_no_rows.ReadOnly = true;
            this.textBox_no_rows.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_rows.TabIndex = 17;
            this.textBox_no_rows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_tables
            // 
            this.textBox_no_tables.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_no_tables.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_no_tables.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_tables.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_tables.Location = new System.Drawing.Point(311, 105);
            this.textBox_no_tables.Name = "textBox_no_tables";
            this.textBox_no_tables.ReadOnly = true;
            this.textBox_no_tables.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_tables.TabIndex = 17;
            this.textBox_no_tables.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_od_zero
            // 
            this.textBox_no_od_zero.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_no_od_zero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_no_od_zero.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_od_zero.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_od_zero.Location = new System.Drawing.Point(311, 42);
            this.textBox_no_od_zero.Name = "textBox_no_od_zero";
            this.textBox_no_od_zero.ReadOnly = true;
            this.textBox_no_od_zero.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_od_zero.TabIndex = 17;
            this.textBox_no_od_zero.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_no_od_2
            // 
            this.textBox_no_od_2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_no_od_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_no_od_2.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.textBox_no_od_2.ForeColor = System.Drawing.Color.Black;
            this.textBox_no_od_2.Location = new System.Drawing.Point(311, 63);
            this.textBox_no_od_2.Name = "textBox_no_od_2";
            this.textBox_no_od_2.ReadOnly = true;
            this.textBox_no_od_2.Size = new System.Drawing.Size(37, 21);
            this.textBox_no_od_2.TabIndex = 17;
            this.textBox_no_od_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_INCORRECT_od
            // 
            this.textBox_INCORRECT_od.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_INCORRECT_od.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_INCORRECT_od.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.textBox_INCORRECT_od.ForeColor = System.Drawing.Color.Cyan;
            this.textBox_INCORRECT_od.Location = new System.Drawing.Point(3, 84);
            this.textBox_INCORRECT_od.Name = "textBox_INCORRECT_od";
            this.textBox_INCORRECT_od.ReadOnly = true;
            this.textBox_INCORRECT_od.Size = new System.Drawing.Size(287, 20);
            this.textBox_INCORRECT_od.TabIndex = 17;
            this.textBox_INCORRECT_od.Text = "Features with Incorrect OD Tables:";
            // 
            // textBox_missing_OD
            // 
            this.textBox_missing_OD.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_missing_OD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_missing_OD.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.textBox_missing_OD.ForeColor = System.Drawing.Color.Yellow;
            this.textBox_missing_OD.Location = new System.Drawing.Point(3, 42);
            this.textBox_missing_OD.Name = "textBox_missing_OD";
            this.textBox_missing_OD.ReadOnly = true;
            this.textBox_missing_OD.Size = new System.Drawing.Size(287, 20);
            this.textBox_missing_OD.TabIndex = 17;
            this.textBox_missing_OD.Text = "Features with Missing OD Tables:";
            // 
            // textBox_OD_TABLES
            // 
            this.textBox_OD_TABLES.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_OD_TABLES.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_OD_TABLES.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.textBox_OD_TABLES.ForeColor = System.Drawing.Color.White;
            this.textBox_OD_TABLES.Location = new System.Drawing.Point(3, 105);
            this.textBox_OD_TABLES.Name = "textBox_OD_TABLES";
            this.textBox_OD_TABLES.ReadOnly = true;
            this.textBox_OD_TABLES.Size = new System.Drawing.Size(287, 20);
            this.textBox_OD_TABLES.TabIndex = 17;
            this.textBox_OD_TABLES.Text = "Total Number of OD Tables on Layer:";
            // 
            // button_Filter
            // 
            this.button_Filter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_Filter.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_Filter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Filter.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_Filter.ForeColor = System.Drawing.Color.White;
            this.button_Filter.Location = new System.Drawing.Point(3, 428);
            this.button_Filter.Name = "button_Filter";
            this.button_Filter.Size = new System.Drawing.Size(129, 25);
            this.button_Filter.TabIndex = 18;
            this.button_Filter.Text = "Filter to Issues";
            this.button_Filter.UseVisualStyleBackColor = false;
            this.button_Filter.Click += new System.EventHandler(this.button_Filter_Click);
            // 
            // button_add_OD_table
            // 
            this.button_add_OD_table.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_add_OD_table.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_add_OD_table.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_add_OD_table.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_add_OD_table.ForeColor = System.Drawing.Color.White;
            this.button_add_OD_table.Location = new System.Drawing.Point(145, 428);
            this.button_add_OD_table.Name = "button_add_OD_table";
            this.button_add_OD_table.Size = new System.Drawing.Size(129, 25);
            this.button_add_OD_table.TabIndex = 3;
            this.button_add_OD_table.Text = "Fix Issues";
            this.button_add_OD_table.UseVisualStyleBackColor = false;
            this.button_add_OD_table.Visible = false;
            this.button_add_OD_table.Click += new System.EventHandler(this.button_add_OD_table_and_remove_wrong_OD_Click);
            // 
            // panel_blocks_and_OD
            // 
            this.panel_blocks_and_OD.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(35)))));
            this.panel_blocks_and_OD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_blocks_and_OD.Controls.Add(this.label3);
            this.panel_blocks_and_OD.Controls.Add(this.radioButton_OD);
            this.panel_blocks_and_OD.Controls.Add(this.radioButton_BLOCKS);
            this.panel_blocks_and_OD.Controls.Add(this.button_refresh_layer_tables);
            this.panel_blocks_and_OD.Location = new System.Drawing.Point(5, 33);
            this.panel_blocks_and_OD.Name = "panel_blocks_and_OD";
            this.panel_blocks_and_OD.Size = new System.Drawing.Size(181, 115);
            this.panel_blocks_and_OD.TabIndex = 66;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label3.Location = new System.Drawing.Point(3, 4);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(131, 18);
            this.label3.TabIndex = 2064;
            this.label3.Text = "Data Type Check";
            // 
            // radioButton_OD
            // 
            this.radioButton_OD.AutoSize = true;
            this.radioButton_OD.Checked = true;
            this.radioButton_OD.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.radioButton_OD.ForeColor = System.Drawing.Color.White;
            this.radioButton_OD.Location = new System.Drawing.Point(6, 30);
            this.radioButton_OD.Name = "radioButton_OD";
            this.radioButton_OD.Size = new System.Drawing.Size(86, 18);
            this.radioButton_OD.TabIndex = 0;
            this.radioButton_OD.TabStop = true;
            this.radioButton_OD.Text = "Object Data";
            this.radioButton_OD.UseVisualStyleBackColor = true;
            // 
            // radioButton_BLOCKS
            // 
            this.radioButton_BLOCKS.AutoSize = true;
            this.radioButton_BLOCKS.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.radioButton_BLOCKS.ForeColor = System.Drawing.Color.White;
            this.radioButton_BLOCKS.Location = new System.Drawing.Point(6, 54);
            this.radioButton_BLOCKS.Name = "radioButton_BLOCKS";
            this.radioButton_BLOCKS.Size = new System.Drawing.Size(114, 18);
            this.radioButton_BLOCKS.TabIndex = 1;
            this.radioButton_BLOCKS.TabStop = true;
            this.radioButton_BLOCKS.Text = "Block Attributes";
            this.radioButton_BLOCKS.UseVisualStyleBackColor = true;
            this.radioButton_BLOCKS.Visible = false;
            this.radioButton_BLOCKS.CheckedChanged += new System.EventHandler(this.radioButton_OD_blocks_CheckedChanged);
            // 
            // button_refresh_layer_tables
            // 
            this.button_refresh_layer_tables.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_refresh_layer_tables.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_refresh_layer_tables.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_refresh_layer_tables.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_refresh_layer_tables.ForeColor = System.Drawing.Color.White;
            this.button_refresh_layer_tables.Location = new System.Drawing.Point(6, 85);
            this.button_refresh_layer_tables.Name = "button_refresh_layer_tables";
            this.button_refresh_layer_tables.Size = new System.Drawing.Size(157, 25);
            this.button_refresh_layer_tables.TabIndex = 41;
            this.button_refresh_layer_tables.Text = "Load Data";
            this.button_refresh_layer_tables.UseVisualStyleBackColor = false;
            this.button_refresh_layer_tables.Click += new System.EventHandler(this.button_load_layers_and_data_tables_Click);
            // 
            // panel_logo
            // 
            this.panel_logo.BackColor = System.Drawing.Color.White;
            this.panel_logo.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("panel_logo.BackgroundImage")));
            this.panel_logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel_logo.Location = new System.Drawing.Point(36, 497);
            this.panel_logo.Name = "panel_logo";
            this.panel_logo.Size = new System.Drawing.Size(121, 100);
            this.panel_logo.TabIndex = 65;
            this.panel_logo.Click += new System.EventHandler(this.panel_logo_DoubleClick);
            // 
            // label_Apply_Changes
            // 
            this.label_Apply_Changes.AutoSize = true;
            this.label_Apply_Changes.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_Apply_Changes.ForeColor = System.Drawing.Color.Yellow;
            this.label_Apply_Changes.Location = new System.Drawing.Point(345, 392);
            this.label_Apply_Changes.MaximumSize = new System.Drawing.Size(600, 0);
            this.label_Apply_Changes.Name = "label_Apply_Changes";
            this.label_Apply_Changes.Size = new System.Drawing.Size(402, 14);
            this.label_Apply_Changes.TabIndex = 64;
            this.label_Apply_Changes.Text = "Edits will only be applied to features when \"Apply Changes\" is pressed.";
            this.label_Apply_Changes.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // panel_excel
            // 
            this.panel_excel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(35)))));
            this.panel_excel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_excel.Controls.Add(this.label5);
            this.panel_excel.Controls.Add(this.button_export_to_excel);
            this.panel_excel.Controls.Add(this.button_import_from_excel);
            this.panel_excel.Location = new System.Drawing.Point(5, 314);
            this.panel_excel.Name = "panel_excel";
            this.panel_excel.Size = new System.Drawing.Size(181, 120);
            this.panel_excel.TabIndex = 62;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label5.Location = new System.Drawing.Point(4, 5);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 18);
            this.label5.TabIndex = 18;
            this.label5.Text = "Excel";
            // 
            // button_export_to_excel
            // 
            this.button_export_to_excel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_export_to_excel.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_export_to_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_export_to_excel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_export_to_excel.ForeColor = System.Drawing.Color.White;
            this.button_export_to_excel.Image = ((System.Drawing.Image)(resources.GetObject("button_export_to_excel.Image")));
            this.button_export_to_excel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_export_to_excel.Location = new System.Drawing.Point(4, 42);
            this.button_export_to_excel.Name = "button_export_to_excel";
            this.button_export_to_excel.Size = new System.Drawing.Size(159, 25);
            this.button_export_to_excel.TabIndex = 13;
            this.button_export_to_excel.Text = "Export to Excel";
            this.button_export_to_excel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_export_to_excel.UseVisualStyleBackColor = false;
            this.button_export_to_excel.Click += new System.EventHandler(this.button_export_to_excel_Click);
            // 
            // button_import_from_excel
            // 
            this.button_import_from_excel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_import_from_excel.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_import_from_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_import_from_excel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_import_from_excel.ForeColor = System.Drawing.Color.White;
            this.button_import_from_excel.Image = ((System.Drawing.Image)(resources.GetObject("button_import_from_excel.Image")));
            this.button_import_from_excel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_import_from_excel.Location = new System.Drawing.Point(4, 81);
            this.button_import_from_excel.Name = "button_import_from_excel";
            this.button_import_from_excel.Size = new System.Drawing.Size(159, 25);
            this.button_import_from_excel.TabIndex = 14;
            this.button_import_from_excel.Text = "Import from Excel";
            this.button_import_from_excel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_import_from_excel.UseVisualStyleBackColor = false;
            this.button_import_from_excel.Click += new System.EventHandler(this.button_import_from_excel_Click);
            // 
            // label_drawing_name
            // 
            this.label_drawing_name.AutoSize = true;
            this.label_drawing_name.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_drawing_name.ForeColor = System.Drawing.Color.White;
            this.label_drawing_name.Location = new System.Drawing.Point(12, 9);
            this.label_drawing_name.Name = "label_drawing_name";
            this.label_drawing_name.Size = new System.Drawing.Size(121, 19);
            this.label_drawing_name.TabIndex = 60;
            this.label_drawing_name.Text = "Drawing Name";
            // 
            // label_correct_od_table
            // 
            this.label_correct_od_table.AutoSize = true;
            this.label_correct_od_table.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_correct_od_table.ForeColor = System.Drawing.Color.White;
            this.label_correct_od_table.Location = new System.Drawing.Point(0, 70);
            this.label_correct_od_table.Name = "label_correct_od_table";
            this.label_correct_od_table.Size = new System.Drawing.Size(99, 14);
            this.label_correct_od_table.TabIndex = 61;
            this.label_correct_od_table.Text = "Correct OD Table";
            // 
            // button_refresh_grid
            // 
            this.button_refresh_grid.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_refresh_grid.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_refresh_grid.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_refresh_grid.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_refresh_grid.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.button_refresh_grid.ForeColor = System.Drawing.Color.White;
            this.button_refresh_grid.Location = new System.Drawing.Point(3, 124);
            this.button_refresh_grid.Name = "button_refresh_grid";
            this.button_refresh_grid.Size = new System.Drawing.Size(160, 25);
            this.button_refresh_grid.TabIndex = 59;
            this.button_refresh_grid.Text = "Build Table";
            this.button_refresh_grid.UseVisualStyleBackColor = false;
            this.button_refresh_grid.Click += new System.EventHandler(this.button_LOAD_DATA_Click);
            // 
            // label_od_block_table
            // 
            this.label_od_block_table.AutoSize = true;
            this.label_od_block_table.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_od_block_table.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label_od_block_table.Location = new System.Drawing.Point(3, 2);
            this.label_od_block_table.Name = "label_od_block_table";
            this.label_od_block_table.Size = new System.Drawing.Size(137, 18);
            this.label_od_block_table.TabIndex = 57;
            this.label_od_block_table.Text = "Object Data Table";
            // 
            // label_current_layer_block
            // 
            this.label_current_layer_block.AutoSize = true;
            this.label_current_layer_block.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label_current_layer_block.ForeColor = System.Drawing.Color.White;
            this.label_current_layer_block.Location = new System.Drawing.Point(0, 26);
            this.label_current_layer_block.Name = "label_current_layer_block";
            this.label_current_layer_block.Size = new System.Drawing.Size(76, 14);
            this.label_current_layer_block.TabIndex = 58;
            this.label_current_layer_block.Text = "Target Layer";
            // 
            // comboBox_layers_blocks_geomanager
            // 
            this.comboBox_layers_blocks_geomanager.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_layers_blocks_geomanager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_layers_blocks_geomanager.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_layers_blocks_geomanager.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.comboBox_layers_blocks_geomanager.ForeColor = System.Drawing.Color.White;
            this.comboBox_layers_blocks_geomanager.FormattingEnabled = true;
            this.comboBox_layers_blocks_geomanager.Location = new System.Drawing.Point(3, 44);
            this.comboBox_layers_blocks_geomanager.Name = "comboBox_layers_blocks_geomanager";
            this.comboBox_layers_blocks_geomanager.Size = new System.Drawing.Size(173, 22);
            this.comboBox_layers_blocks_geomanager.TabIndex = 55;
            // 
            // comboBox_od_existing_tables
            // 
            this.comboBox_od_existing_tables.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_od_existing_tables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_od_existing_tables.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_od_existing_tables.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.comboBox_od_existing_tables.ForeColor = System.Drawing.Color.White;
            this.comboBox_od_existing_tables.FormattingEnabled = true;
            this.comboBox_od_existing_tables.Location = new System.Drawing.Point(3, 88);
            this.comboBox_od_existing_tables.Name = "comboBox_od_existing_tables";
            this.comboBox_od_existing_tables.Size = new System.Drawing.Size(173, 22);
            this.comboBox_od_existing_tables.TabIndex = 56;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(35)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label_layer_rules);
            this.panel1.Controls.Add(this.label_current_layer_block);
            this.panel1.Controls.Add(this.comboBox_layers_blocks_geomanager);
            this.panel1.Controls.Add(this.label_correct_od_table);
            this.panel1.Controls.Add(this.comboBox_od_existing_tables);
            this.panel1.Controls.Add(this.button_refresh_grid);
            this.panel1.Location = new System.Drawing.Point(5, 154);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(181, 154);
            this.panel1.TabIndex = 69;
            // 
            // label_layer_rules
            // 
            this.label_layer_rules.AutoSize = true;
            this.label_layer_rules.BackColor = System.Drawing.Color.Transparent;
            this.label_layer_rules.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label_layer_rules.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label_layer_rules.Location = new System.Drawing.Point(4, 4);
            this.label_layer_rules.Name = "label_layer_rules";
            this.label_layer_rules.Size = new System.Drawing.Size(93, 18);
            this.label_layer_rules.TabIndex = 2065;
            this.label_layer_rules.Text = "Layer Rules";
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(35)))));
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.panel_navigation);
            this.panel4.Controls.Add(this.label_dt_issues);
            this.panel4.Controls.Add(this.Button_Update_object_data);
            this.panel4.Controls.Add(this.panel_stats);
            this.panel4.Controls.Add(this.panel_grid);
            this.panel4.Controls.Add(this.button_Filter);
            this.panel4.Controls.Add(this.label_od_block_table);
            this.panel4.Controls.Add(this.label_Apply_Changes);
            this.panel4.Controls.Add(this.button_add_OD_table);
            this.panel4.Location = new System.Drawing.Point(188, 33);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(792, 615);
            this.panel4.TabIndex = 70;
            // 
            // panel_navigation
            // 
            this.panel_navigation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(35)))));
            this.panel_navigation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_navigation.Controls.Add(this.label10);
            this.panel_navigation.Controls.Add(this.button_zoom_row_object_data);
            this.panel_navigation.Controls.Add(this.button_zoom);
            this.panel_navigation.Location = new System.Drawing.Point(364, 472);
            this.panel_navigation.Name = "panel_navigation";
            this.panel_navigation.Size = new System.Drawing.Size(155, 134);
            this.panel_navigation.TabIndex = 2064;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label10.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label10.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label10.Location = new System.Drawing.Point(5, 4);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(128, 18);
            this.label10.TabIndex = 2063;
            this.label10.Text = "Navigation Tools";
            // 
            // button_zoom_row_object_data
            // 
            this.button_zoom_row_object_data.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_zoom_row_object_data.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_zoom_row_object_data.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_zoom_row_object_data.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_zoom_row_object_data.ForeColor = System.Drawing.Color.White;
            this.button_zoom_row_object_data.Image = ((System.Drawing.Image)(resources.GetObject("button_zoom_row_object_data.Image")));
            this.button_zoom_row_object_data.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_zoom_row_object_data.Location = new System.Drawing.Point(6, 28);
            this.button_zoom_row_object_data.Name = "button_zoom_row_object_data";
            this.button_zoom_row_object_data.Size = new System.Drawing.Size(144, 25);
            this.button_zoom_row_object_data.TabIndex = 5;
            this.button_zoom_row_object_data.Text = "Select Feature";
            this.button_zoom_row_object_data.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_zoom_row_object_data.UseVisualStyleBackColor = false;
            this.button_zoom_row_object_data.Click += new System.EventHandler(this.button_go_to_table_row_Click);
            // 
            // button_zoom
            // 
            this.button_zoom.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_zoom.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_zoom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_zoom.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_zoom.ForeColor = System.Drawing.Color.White;
            this.button_zoom.Image = ((System.Drawing.Image)(resources.GetObject("button_zoom.Image")));
            this.button_zoom.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_zoom.Location = new System.Drawing.Point(6, 62);
            this.button_zoom.Name = "button_zoom";
            this.button_zoom.Size = new System.Drawing.Size(144, 25);
            this.button_zoom.TabIndex = 6;
            this.button_zoom.Text = "Zoom To Feature";
            this.button_zoom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_zoom.UseVisualStyleBackColor = false;
            this.button_zoom.Click += new System.EventHandler(this.button_zoom_Click);
            // 
            // panel_grid
            // 
            this.panel_grid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel_grid.Controls.Add(this.label_processing1);
            this.panel_grid.Controls.Add(this.DataGridView_data);
            this.panel_grid.Location = new System.Drawing.Point(6, 23);
            this.panel_grid.Name = "panel_grid";
            this.panel_grid.Size = new System.Drawing.Size(779, 366);
            this.panel_grid.TabIndex = 55;
            // 
            // label_processing1
            // 
            this.label_processing1.AutoSize = true;
            this.label_processing1.BackColor = System.Drawing.Color.Transparent;
            this.label_processing1.Font = new System.Drawing.Font("Arial", 15F, System.Drawing.FontStyle.Bold);
            this.label_processing1.ForeColor = System.Drawing.Color.White;
            this.label_processing1.Location = new System.Drawing.Point(233, 169);
            this.label_processing1.Name = "label_processing1";
            this.label_processing1.Size = new System.Drawing.Size(255, 24);
            this.label_processing1.TabIndex = 35;
            this.label_processing1.Text = "Processing... Please Wait.";
            // 
            // DataGridView_data
            // 
            this.DataGridView_data.AllowUserToAddRows = false;
            this.DataGridView_data.AllowUserToDeleteRows = false;
            this.DataGridView_data.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.DataGridView_data.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.DataGridView_data.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.DataGridView_data.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.DataGridView_data.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.DataGridView_data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DataGridView_data.DefaultCellStyle = dataGridViewCellStyle1;
            this.DataGridView_data.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DataGridView_data.GridColor = System.Drawing.Color.LightGray;
            this.DataGridView_data.Location = new System.Drawing.Point(0, 0);
            this.DataGridView_data.Name = "DataGridView_data";
            this.DataGridView_data.RowHeadersVisible = false;
            this.DataGridView_data.Size = new System.Drawing.Size(775, 362);
            this.DataGridView_data.TabIndex = 0;
            this.DataGridView_data.VirtualMode = true;
            this.DataGridView_data.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGrid_od_data_CellClick);
            this.DataGridView_data.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView_data_CellValueChanged);
            this.DataGridView_data.Sorted += new System.EventHandler(this.DataGridView_OD_data_Sorted);
            // 
            // Igen_geomanager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(992, 661);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel_blocks_and_OD);
            this.Controls.Add(this.panel_logo);
            this.Controls.Add(this.panel_excel);
            this.Controls.Add(this.label_drawing_name);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Igen_geomanager";
            this.Text = "AGEN_0X_GeoDataTables";
            this.panel_stats.ResumeLayout(false);
            this.panel_stats.PerformLayout();
            this.panel_blocks_and_OD.ResumeLayout(false);
            this.panel_blocks_and_OD.PerformLayout();
            this.panel_excel.ResumeLayout(false);
            this.panel_excel.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel_navigation.ResumeLayout(false);
            this.panel_navigation.PerformLayout();
            this.panel_grid.ResumeLayout(false);
            this.panel_grid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView_data)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button Button_Update_object_data;
        private System.Windows.Forms.Label label_dt_issues;
        private System.Windows.Forms.Panel panel_stats;
        private System.Windows.Forms.TextBox textBox_Features;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_MultipleOD;
        private System.Windows.Forms.TextBox textBox_no_wrong_od;
        private System.Windows.Forms.TextBox textBox_no_rows;
        private System.Windows.Forms.TextBox textBox_no_tables;
        private System.Windows.Forms.TextBox textBox_no_od_zero;
        private System.Windows.Forms.TextBox textBox_no_od_2;
        private System.Windows.Forms.TextBox textBox_INCORRECT_od;
        private System.Windows.Forms.TextBox textBox_missing_OD;
        private System.Windows.Forms.TextBox textBox_OD_TABLES;
        private System.Windows.Forms.Button button_Filter;
        private System.Windows.Forms.Button button_add_OD_table;
        private System.Windows.Forms.Panel panel_blocks_and_OD;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RadioButton radioButton_OD;
        private System.Windows.Forms.Button button_refresh_layer_tables;
        private System.Windows.Forms.Panel panel_logo;
        private System.Windows.Forms.Label label_Apply_Changes;
        private System.Windows.Forms.Panel panel_excel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button_export_to_excel;
        private System.Windows.Forms.Button button_import_from_excel;
        private System.Windows.Forms.Label label_drawing_name;
        private System.Windows.Forms.Label label_correct_od_table;
        private System.Windows.Forms.Button button_refresh_grid;
        private System.Windows.Forms.Label label_od_block_table;
        private System.Windows.Forms.Label label_current_layer_block;
        private System.Windows.Forms.ComboBox comboBox_layers_blocks_geomanager;
        private System.Windows.Forms.ComboBox comboBox_od_existing_tables;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel_navigation;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button button_zoom_row_object_data;
        private System.Windows.Forms.Button button_zoom;
        private System.Windows.Forms.Panel panel_grid;
        private System.Windows.Forms.Label label_processing1;
        private System.Windows.Forms.DataGridView DataGridView_data;
        private System.Windows.Forms.Label label_layer_rules;
        private System.Windows.Forms.RadioButton radioButton_BLOCKS;
    }
}