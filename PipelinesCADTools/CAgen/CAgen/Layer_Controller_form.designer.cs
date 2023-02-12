namespace Alignment_mdi
{
    partial class Layer_controller_form
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Layer_controller_form));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel23 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.button_dwg_to_excel = new System.Windows.Forms.Button();
            this.comboBox_visretain = new System.Windows.Forms.ComboBox();
            this.comboBox_config_tabs = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label_dwg = new System.Windows.Forms.Label();
            this.label_excel_file = new System.Windows.Forms.Label();
            this.button_select_excel_file = new System.Windows.Forms.Button();
            this.panel19 = new System.Windows.Forms.Panel();
            this.label56 = new System.Windows.Forms.Label();
            this.dataGridView_drawings = new System.Windows.Forms.DataGridView();
            this.label65 = new System.Windows.Forms.Label();
            this.button_load_block_attributes_to_excel = new System.Windows.Forms.Button();
            this.button_select_drawings = new System.Windows.Forms.Button();
            this.button_export_layers_from_selection = new System.Windows.Forms.Button();
            this.button_open_excel_tblk_attrib = new System.Windows.Forms.Button();
            this.button_excel_to_dwg = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panel23.SuspendLayout();
            this.panel19.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_drawings)).BeginInit();
            this.SuspendLayout();
            // 
            // panel23
            // 
            this.panel23.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel23.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel23.Controls.Add(this.label3);
            this.panel23.Controls.Add(this.button_dwg_to_excel);
            this.panel23.Controls.Add(this.comboBox_visretain);
            this.panel23.Controls.Add(this.comboBox_config_tabs);
            this.panel23.Controls.Add(this.label1);
            this.panel23.Controls.Add(this.label2);
            this.panel23.Controls.Add(this.label_dwg);
            this.panel23.Controls.Add(this.label_excel_file);
            this.panel23.Controls.Add(this.button_select_excel_file);
            this.panel23.Controls.Add(this.panel19);
            this.panel23.Controls.Add(this.button_load_block_attributes_to_excel);
            this.panel23.Controls.Add(this.button_select_drawings);
            this.panel23.Controls.Add(this.button_export_layers_from_selection);
            this.panel23.Controls.Add(this.button_open_excel_tblk_attrib);
            this.panel23.Controls.Add(this.button_excel_to_dwg);
            this.panel23.Location = new System.Drawing.Point(12, 12);
            this.panel23.Name = "panel23";
            this.panel23.Size = new System.Drawing.Size(1126, 613);
            this.panel23.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Arial Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Yellow;
            this.label3.Location = new System.Drawing.Point(192, 586);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 18);
            this.label3.TabIndex = 2094;
            this.label3.Text = "V1.2";
            // 
            // button_dwg_to_excel
            // 
            this.button_dwg_to_excel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_dwg_to_excel.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_dwg_to_excel.BackgroundImage")));
            this.button_dwg_to_excel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_dwg_to_excel.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_dwg_to_excel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_dwg_to_excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_dwg_to_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_dwg_to_excel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_dwg_to_excel.ForeColor = System.Drawing.Color.White;
            this.button_dwg_to_excel.Location = new System.Drawing.Point(9, 220);
            this.button_dwg_to_excel.Name = "button_dwg_to_excel";
            this.button_dwg_to_excel.Size = new System.Drawing.Size(223, 28);
            this.button_dwg_to_excel.TabIndex = 2093;
            this.button_dwg_to_excel.Text = "                DWG                      XL";
            this.button_dwg_to_excel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_dwg_to_excel.UseVisualStyleBackColor = false;
            this.button_dwg_to_excel.Click += new System.EventHandler(this.Button_dwg_to_excel_Click);
            // 
            // comboBox_visretain
            // 
            this.comboBox_visretain.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_visretain.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_visretain.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_visretain.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
            this.comboBox_visretain.ForeColor = System.Drawing.Color.White;
            this.comboBox_visretain.FormattingEnabled = true;
            this.comboBox_visretain.Items.AddRange(new object[] {
            "",
            "0",
            "1"});
            this.comboBox_visretain.Location = new System.Drawing.Point(149, 284);
            this.comboBox_visretain.Name = "comboBox_visretain";
            this.comboBox_visretain.Size = new System.Drawing.Size(81, 22);
            this.comboBox_visretain.TabIndex = 2092;
            this.comboBox_visretain.SelectedIndexChanged += new System.EventHandler(this.ComboBox_config_tabs_SelectedIndexChanged);
            // 
            // comboBox_config_tabs
            // 
            this.comboBox_config_tabs.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_config_tabs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_config_tabs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_config_tabs.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.comboBox_config_tabs.ForeColor = System.Drawing.Color.White;
            this.comboBox_config_tabs.FormattingEnabled = true;
            this.comboBox_config_tabs.Location = new System.Drawing.Point(336, 25);
            this.comboBox_config_tabs.Name = "comboBox_config_tabs";
            this.comboBox_config_tabs.Size = new System.Drawing.Size(201, 24);
            this.comboBox_config_tabs.TabIndex = 2092;
            this.comboBox_config_tabs.SelectedIndexChanged += new System.EventHandler(this.ComboBox_config_tabs_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Arial Black", 8F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(65, 286);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 15);
            this.label1.TabIndex = 2052;
            this.label1.Text = "VISRETAIN:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Arial Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(238, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 18);
            this.label2.TabIndex = 2052;
            this.label2.Text = "STND TABS:";
            // 
            // label_dwg
            // 
            this.label_dwg.AutoSize = true;
            this.label_dwg.BackColor = System.Drawing.Color.Transparent;
            this.label_dwg.Font = new System.Drawing.Font("Arial Black", 9F, System.Drawing.FontStyle.Bold);
            this.label_dwg.ForeColor = System.Drawing.Color.White;
            this.label_dwg.Location = new System.Drawing.Point(543, 27);
            this.label_dwg.Name = "label_dwg";
            this.label_dwg.Size = new System.Drawing.Size(43, 17);
            this.label_dwg.TabIndex = 2052;
            this.label_dwg.Text = "DWG:";
            // 
            // label_excel_file
            // 
            this.label_excel_file.AutoSize = true;
            this.label_excel_file.BackColor = System.Drawing.Color.Transparent;
            this.label_excel_file.Font = new System.Drawing.Font("Arial Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_excel_file.ForeColor = System.Drawing.Color.Red;
            this.label_excel_file.Location = new System.Drawing.Point(5, 2);
            this.label_excel_file.Name = "label_excel_file";
            this.label_excel_file.Size = new System.Drawing.Size(242, 18);
            this.label_excel_file.TabIndex = 2052;
            this.label_excel_file.Text = "Standard excel file not specified";
            // 
            // button_select_excel_file
            // 
            this.button_select_excel_file.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_select_excel_file.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_select_excel_file.BackgroundImage")));
            this.button_select_excel_file.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_select_excel_file.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_select_excel_file.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_select_excel_file.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_select_excel_file.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_select_excel_file.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_select_excel_file.ForeColor = System.Drawing.Color.White;
            this.button_select_excel_file.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_select_excel_file.Location = new System.Drawing.Point(6, 26);
            this.button_select_excel_file.Name = "button_select_excel_file";
            this.button_select_excel_file.Size = new System.Drawing.Size(226, 28);
            this.button_select_excel_file.TabIndex = 2076;
            this.button_select_excel_file.Text = "Select Stnd Definition (.xls)";
            this.button_select_excel_file.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_select_excel_file.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_select_excel_file.UseVisualStyleBackColor = false;
            this.button_select_excel_file.Click += new System.EventHandler(this.button_select_excel_file_Click);
            // 
            // panel19
            // 
            this.panel19.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.panel19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel19.Controls.Add(this.label56);
            this.panel19.Controls.Add(this.dataGridView_drawings);
            this.panel19.Controls.Add(this.label65);
            this.panel19.Location = new System.Drawing.Point(241, 57);
            this.panel19.Name = "panel19";
            this.panel19.Size = new System.Drawing.Size(880, 551);
            this.panel19.TabIndex = 2078;
            // 
            // label56
            // 
            this.label56.AutoSize = true;
            this.label56.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label56.ForeColor = System.Drawing.Color.Yellow;
            this.label56.Location = new System.Drawing.Point(3, 18);
            this.label56.Name = "label56";
            this.label56.Size = new System.Drawing.Size(165, 14);
            this.label56.TabIndex = 2040;
            this.label56.Text = "*(to have the layers updated)";
            // 
            // dataGridView_drawings
            // 
            this.dataGridView_drawings.AllowUserToAddRows = false;
            this.dataGridView_drawings.AllowUserToDeleteRows = false;
            this.dataGridView_drawings.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView_drawings.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.dataGridView_drawings.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView_drawings.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.dataGridView_drawings.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.dataGridView_drawings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_drawings.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_drawings.GridColor = System.Drawing.Color.LightGray;
            this.dataGridView_drawings.Location = new System.Drawing.Point(3, 35);
            this.dataGridView_drawings.Name = "dataGridView_drawings";
            this.dataGridView_drawings.RowHeadersVisible = false;
            this.dataGridView_drawings.Size = new System.Drawing.Size(872, 511);
            this.dataGridView_drawings.TabIndex = 19;
            this.dataGridView_drawings.TabStop = false;
            this.dataGridView_drawings.VirtualMode = true;
            this.dataGridView_drawings.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_drawings_CellMouseClick);
            this.dataGridView_drawings.Click += new System.EventHandler(this.dataGridView_drawings_Click);
            // 
            // label65
            // 
            this.label65.AutoSize = true;
            this.label65.BackColor = System.Drawing.Color.Transparent;
            this.label65.Font = new System.Drawing.Font("Arial Black", 9.75F);
            this.label65.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.label65.Location = new System.Drawing.Point(3, 0);
            this.label65.Name = "label65";
            this.label65.Size = new System.Drawing.Size(75, 18);
            this.label65.TabIndex = 2039;
            this.label65.Text = "Drawings";
            // 
            // button_load_block_attributes_to_excel
            // 
            this.button_load_block_attributes_to_excel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_load_block_attributes_to_excel.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_load_block_attributes_to_excel.BackgroundImage")));
            this.button_load_block_attributes_to_excel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_load_block_attributes_to_excel.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_load_block_attributes_to_excel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_block_attributes_to_excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_load_block_attributes_to_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_block_attributes_to_excel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_load_block_attributes_to_excel.ForeColor = System.Drawing.Color.White;
            this.button_load_block_attributes_to_excel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_load_block_attributes_to_excel.Location = new System.Drawing.Point(6, 60);
            this.button_load_block_attributes_to_excel.Name = "button_load_block_attributes_to_excel";
            this.button_load_block_attributes_to_excel.Size = new System.Drawing.Size(226, 28);
            this.button_load_block_attributes_to_excel.TabIndex = 5;
            this.button_load_block_attributes_to_excel.Text = "Select Layer Stnd (.dwg)";
            this.button_load_block_attributes_to_excel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_load_block_attributes_to_excel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_load_block_attributes_to_excel.UseVisualStyleBackColor = false;
            this.button_load_block_attributes_to_excel.Click += new System.EventHandler(this.Create_layer_controller_spreadsheet_Click);
            // 
            // button_select_drawings
            // 
            this.button_select_drawings.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_select_drawings.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_select_drawings.BackgroundImage")));
            this.button_select_drawings.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_select_drawings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_select_drawings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_select_drawings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_select_drawings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_select_drawings.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_select_drawings.ForeColor = System.Drawing.Color.White;
            this.button_select_drawings.Location = new System.Drawing.Point(9, 120);
            this.button_select_drawings.Name = "button_select_drawings";
            this.button_select_drawings.Size = new System.Drawing.Size(223, 28);
            this.button_select_drawings.TabIndex = 2077;
            this.button_select_drawings.Text = "Select Drawings";
            this.button_select_drawings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_select_drawings.UseVisualStyleBackColor = false;
            this.button_select_drawings.Click += new System.EventHandler(this.button_select_drawings_Click);
            // 
            // button_export_layers_from_selection
            // 
            this.button_export_layers_from_selection.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_export_layers_from_selection.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_export_layers_from_selection.BackgroundImage")));
            this.button_export_layers_from_selection.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_export_layers_from_selection.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_export_layers_from_selection.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_export_layers_from_selection.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_export_layers_from_selection.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_export_layers_from_selection.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_export_layers_from_selection.ForeColor = System.Drawing.Color.White;
            this.button_export_layers_from_selection.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_export_layers_from_selection.Location = new System.Drawing.Point(9, 397);
            this.button_export_layers_from_selection.Name = "button_export_layers_from_selection";
            this.button_export_layers_from_selection.Size = new System.Drawing.Size(223, 28);
            this.button_export_layers_from_selection.TabIndex = 2076;
            this.button_export_layers_from_selection.Text = "Export layers to Excel";
            this.button_export_layers_from_selection.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_export_layers_from_selection.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_export_layers_from_selection.UseVisualStyleBackColor = false;
            this.button_export_layers_from_selection.Click += new System.EventHandler(this.button_export_layers_from_selection_Click);
            // 
            // button_open_excel_tblk_attrib
            // 
            this.button_open_excel_tblk_attrib.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_open_excel_tblk_attrib.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_open_excel_tblk_attrib.BackgroundImage")));
            this.button_open_excel_tblk_attrib.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button_open_excel_tblk_attrib.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_open_excel_tblk_attrib.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_open_excel_tblk_attrib.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_open_excel_tblk_attrib.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_open_excel_tblk_attrib.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_open_excel_tblk_attrib.ForeColor = System.Drawing.Color.White;
            this.button_open_excel_tblk_attrib.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_open_excel_tblk_attrib.Location = new System.Drawing.Point(9, 544);
            this.button_open_excel_tblk_attrib.Name = "button_open_excel_tblk_attrib";
            this.button_open_excel_tblk_attrib.Size = new System.Drawing.Size(223, 28);
            this.button_open_excel_tblk_attrib.TabIndex = 2076;
            this.button_open_excel_tblk_attrib.Text = "Open Layer Standards";
            this.button_open_excel_tblk_attrib.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_open_excel_tblk_attrib.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button_open_excel_tblk_attrib.UseVisualStyleBackColor = false;
            this.button_open_excel_tblk_attrib.Click += new System.EventHandler(this.button_open_excel_tblk_attributes_Click);
            // 
            // button_excel_to_dwg
            // 
            this.button_excel_to_dwg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_excel_to_dwg.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_excel_to_dwg.BackgroundImage")));
            this.button_excel_to_dwg.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_excel_to_dwg.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_excel_to_dwg.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_excel_to_dwg.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_excel_to_dwg.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_excel_to_dwg.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_excel_to_dwg.ForeColor = System.Drawing.Color.White;
            this.button_excel_to_dwg.Location = new System.Drawing.Point(9, 326);
            this.button_excel_to_dwg.Name = "button_excel_to_dwg";
            this.button_excel_to_dwg.Size = new System.Drawing.Size(223, 28);
            this.button_excel_to_dwg.TabIndex = 1;
            this.button_excel_to_dwg.Text = "                   XL                     DWG";
            this.button_excel_to_dwg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_excel_to_dwg.UseVisualStyleBackColor = false;
            this.button_excel_to_dwg.Click += new System.EventHandler(this.button_excel_to_dwg_Click);
            // 
            // Layer_controller_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.ClientSize = new System.Drawing.Size(1150, 637);
            this.Controls.Add(this.panel23);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Layer_controller_form";
            this.Text = "AGENCrossingBandSettingsForm";
            this.panel23.ResumeLayout(false);
            this.panel23.PerformLayout();
            this.panel19.ResumeLayout(false);
            this.panel19.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_drawings)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button_excel_to_dwg;
        private System.Windows.Forms.Panel panel23;
        private System.Windows.Forms.Button button_open_excel_tblk_attrib;
        private System.Windows.Forms.Button button_select_drawings;
        private System.Windows.Forms.Panel panel19;
        private System.Windows.Forms.Label label56;
        private System.Windows.Forms.Label label65;
        private System.Windows.Forms.Button button_load_block_attributes_to_excel;
        private System.Windows.Forms.Label label_excel_file;
        private System.Windows.Forms.Button button_select_excel_file;
        private System.Windows.Forms.ComboBox comboBox_config_tabs;
        private System.Windows.Forms.Label label_dwg;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_dwg_to_excel;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.DataGridView dataGridView_drawings;
        private System.Windows.Forms.ComboBox comboBox_visretain;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_export_layers_from_selection;
        private System.Windows.Forms.Label label3;
    }
}