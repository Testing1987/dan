namespace Workspace_band_Csharp
{
    partial class Workspace_band_form
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage_pick_from_existing = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_access_database_location = new System.Windows.Forms.TextBox();
            this.textBox_COMPILED = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel_left_right = new System.Windows.Forms.Panel();
            this.radioButton_right_left = new System.Windows.Forms.RadioButton();
            this.radioButton_left_right = new System.Windows.Forms.RadioButton();
            this.button_pick_CL_EASEMENT_WS = new System.Windows.Forms.Button();
            this.button_draw_from_pick = new System.Windows.Forms.Button();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel_matchlines = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_Matchline_end = new System.Windows.Forms.TextBox();
            this.textBox_Matchline_start = new System.Windows.Forms.TextBox();
            this.button_DRAW_COMPILED = new System.Windows.Forms.Button();
            this.button_DRAW_BAND = new System.Windows.Forms.Button();
            this.Button_connect_to_wspaceDB = new System.Windows.Forms.Button();
            this.Button_connect_to_atwsDB = new System.Windows.Forms.Button();
            this.Button_connect_to_rowDB = new System.Windows.Forms.Button();
            this.tabPage_CONFIG = new System.Windows.Forms.TabPage();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_ATWS_table_name = new System.Windows.Forms.TextBox();
            this.textBox_WORKSPACE_TABLE_NAME = new System.Windows.Forms.TextBox();
            this.textBox_row_table_name = new System.Windows.Forms.TextBox();
            this.panel_UP = new System.Windows.Forms.Panel();
            this.radioButton_right_UP = new System.Windows.Forms.RadioButton();
            this.radioButton_left_UP = new System.Windows.Forms.RadioButton();
            this.tabControl1.SuspendLayout();
            this.tabPage_pick_from_existing.SuspendLayout();
            this.panel_left_right.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel_matchlines.SuspendLayout();
            this.tabPage_CONFIG.SuspendLayout();
            this.panel_UP.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabPage_pick_from_existing);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage_CONFIG);
            this.tabControl1.Location = new System.Drawing.Point(14, 14);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(520, 279);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage_pick_from_existing
            // 
            this.tabPage_pick_from_existing.BackColor = System.Drawing.Color.Gainsboro;
            this.tabPage_pick_from_existing.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage_pick_from_existing.Controls.Add(this.label6);
            this.tabPage_pick_from_existing.Controls.Add(this.textBox_access_database_location);
            this.tabPage_pick_from_existing.Controls.Add(this.textBox_COMPILED);
            this.tabPage_pick_from_existing.Controls.Add(this.label1);
            this.tabPage_pick_from_existing.Controls.Add(this.panel_UP);
            this.tabPage_pick_from_existing.Controls.Add(this.panel_left_right);
            this.tabPage_pick_from_existing.Controls.Add(this.button_pick_CL_EASEMENT_WS);
            this.tabPage_pick_from_existing.Controls.Add(this.button_draw_from_pick);
            this.tabPage_pick_from_existing.Location = new System.Drawing.Point(4, 27);
            this.tabPage_pick_from_existing.Name = "tabPage_pick_from_existing";
            this.tabPage_pick_from_existing.Size = new System.Drawing.Size(512, 248);
            this.tabPage_pick_from_existing.TabIndex = 2;
            this.tabPage_pick_from_existing.Text = "RICE ENERGY";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label6.Location = new System.Drawing.Point(5, 58);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(130, 17);
            this.label6.TabIndex = 1;
            this.label6.Text = "Compiled Table name";
            // 
            // textBox_access_database_location
            // 
            this.textBox_access_database_location.ForeColor = System.Drawing.Color.Black;
            this.textBox_access_database_location.Location = new System.Drawing.Point(177, 3);
            this.textBox_access_database_location.Multiline = true;
            this.textBox_access_database_location.Name = "textBox_access_database_location";
            this.textBox_access_database_location.Size = new System.Drawing.Size(328, 46);
            this.textBox_access_database_location.TabIndex = 0;
            this.textBox_access_database_location.Text = "C:\\Users\\pop70694\\Documents\\Work Files\\2016-01-04 RICE\\RICE_RW_CONFIG.accdb";
            this.textBox_access_database_location.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_COMPILED
            // 
            this.textBox_COMPILED.ForeColor = System.Drawing.Color.Black;
            this.textBox_COMPILED.Location = new System.Drawing.Point(177, 55);
            this.textBox_COMPILED.Name = "textBox_COMPILED";
            this.textBox_COMPILED.Size = new System.Drawing.Size(328, 21);
            this.textBox_COMPILED.TabIndex = 0;
            this.textBox_COMPILED.Text = "RICE_2016_01_20";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(5, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Access database location";
            // 
            // panel_left_right
            // 
            this.panel_left_right.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel_left_right.Controls.Add(this.radioButton_right_left);
            this.panel_left_right.Controls.Add(this.radioButton_left_right);
            this.panel_left_right.Location = new System.Drawing.Point(8, 93);
            this.panel_left_right.Name = "panel_left_right";
            this.panel_left_right.Size = new System.Drawing.Size(114, 57);
            this.panel_left_right.TabIndex = 5;
            // 
            // radioButton_right_left
            // 
            this.radioButton_right_left.AutoSize = true;
            this.radioButton_right_left.Location = new System.Drawing.Point(3, 28);
            this.radioButton_right_left.Name = "radioButton_right_left";
            this.radioButton_right_left.Size = new System.Drawing.Size(103, 19);
            this.radioButton_right_left.TabIndex = 4;
            this.radioButton_right_left.Text = "RIGHT to LEFT";
            this.radioButton_right_left.UseVisualStyleBackColor = true;
            // 
            // radioButton_left_right
            // 
            this.radioButton_left_right.AutoSize = true;
            this.radioButton_left_right.Checked = true;
            this.radioButton_left_right.Location = new System.Drawing.Point(3, 3);
            this.radioButton_left_right.Name = "radioButton_left_right";
            this.radioButton_left_right.Size = new System.Drawing.Size(103, 19);
            this.radioButton_left_right.TabIndex = 4;
            this.radioButton_left_right.TabStop = true;
            this.radioButton_left_right.Text = "LEFT to RIGHT";
            this.radioButton_left_right.UseVisualStyleBackColor = true;
            // 
            // button_pick_CL_EASEMENT_WS
            // 
            this.button_pick_CL_EASEMENT_WS.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.button_pick_CL_EASEMENT_WS.Location = new System.Drawing.Point(279, 82);
            this.button_pick_CL_EASEMENT_WS.Name = "button_pick_CL_EASEMENT_WS";
            this.button_pick_CL_EASEMENT_WS.Size = new System.Drawing.Size(226, 80);
            this.button_pick_CL_EASEMENT_WS.TabIndex = 3;
            this.button_pick_CL_EASEMENT_WS.Text = "Pick Centerline - Esement - Temporary Workspace";
            this.button_pick_CL_EASEMENT_WS.UseVisualStyleBackColor = true;
            this.button_pick_CL_EASEMENT_WS.Click += new System.EventHandler(this.button_Pick_information_for_CL_EASEMENT_TWS);
            // 
            // button_draw_from_pick
            // 
            this.button_draw_from_pick.Location = new System.Drawing.Point(-2, 186);
            this.button_draw_from_pick.Name = "button_draw_from_pick";
            this.button_draw_from_pick.Size = new System.Drawing.Size(502, 56);
            this.button_draw_from_pick.TabIndex = 3;
            this.button_draw_from_pick.Text = "Draw";
            this.button_draw_from_pick.UseVisualStyleBackColor = true;
            this.button_draw_from_pick.Click += new System.EventHandler(this.Button_draw_from_compiled_database);
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Gainsboro;
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage1.Controls.Add(this.panel_matchlines);
            this.tabPage1.Controls.Add(this.button_DRAW_COMPILED);
            this.tabPage1.Controls.Add(this.button_DRAW_BAND);
            this.tabPage1.Controls.Add(this.Button_connect_to_wspaceDB);
            this.tabPage1.Controls.Add(this.Button_connect_to_atwsDB);
            this.tabPage1.Controls.Add(this.Button_connect_to_rowDB);
            this.tabPage1.ForeColor = System.Drawing.Color.Black;
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(512, 248);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Access";
            // 
            // panel_matchlines
            // 
            this.panel_matchlines.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel_matchlines.Controls.Add(this.label2);
            this.panel_matchlines.Controls.Add(this.textBox_Matchline_end);
            this.panel_matchlines.Controls.Add(this.textBox_Matchline_start);
            this.panel_matchlines.Location = new System.Drawing.Point(310, 6);
            this.panel_matchlines.Name = "panel_matchlines";
            this.panel_matchlines.Size = new System.Drawing.Size(196, 154);
            this.panel_matchlines.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Matchlines";
            // 
            // textBox_Matchline_end
            // 
            this.textBox_Matchline_end.ForeColor = System.Drawing.Color.Black;
            this.textBox_Matchline_end.Location = new System.Drawing.Point(3, 50);
            this.textBox_Matchline_end.Name = "textBox_Matchline_end";
            this.textBox_Matchline_end.Size = new System.Drawing.Size(100, 21);
            this.textBox_Matchline_end.TabIndex = 3;
            this.textBox_Matchline_end.Text = "90000";
            // 
            // textBox_Matchline_start
            // 
            this.textBox_Matchline_start.ForeColor = System.Drawing.Color.Black;
            this.textBox_Matchline_start.Location = new System.Drawing.Point(3, 23);
            this.textBox_Matchline_start.Name = "textBox_Matchline_start";
            this.textBox_Matchline_start.Size = new System.Drawing.Size(100, 21);
            this.textBox_Matchline_start.TabIndex = 3;
            this.textBox_Matchline_start.Text = "0";
            // 
            // button_DRAW_COMPILED
            // 
            this.button_DRAW_COMPILED.BackColor = System.Drawing.Color.Yellow;
            this.button_DRAW_COMPILED.Location = new System.Drawing.Point(313, 197);
            this.button_DRAW_COMPILED.Name = "button_DRAW_COMPILED";
            this.button_DRAW_COMPILED.Size = new System.Drawing.Size(196, 45);
            this.button_DRAW_COMPILED.TabIndex = 2;
            this.button_DRAW_COMPILED.Text = "Draw Band from compiled table";
            this.button_DRAW_COMPILED.UseVisualStyleBackColor = false;
            this.button_DRAW_COMPILED.Click += new System.EventHandler(this.Button_draw_from_compiled);
            // 
            // button_DRAW_BAND
            // 
            this.button_DRAW_BAND.Location = new System.Drawing.Point(6, 115);
            this.button_DRAW_BAND.Name = "button_DRAW_BAND";
            this.button_DRAW_BAND.Size = new System.Drawing.Size(298, 45);
            this.button_DRAW_BAND.TabIndex = 2;
            this.button_DRAW_BAND.Text = "Draw Band";
            this.button_DRAW_BAND.UseVisualStyleBackColor = true;
            this.button_DRAW_BAND.Click += new System.EventHandler(this.Button_draw_click);
            // 
            // Button_connect_to_wspaceDB
            // 
            this.Button_connect_to_wspaceDB.Location = new System.Drawing.Point(158, 6);
            this.Button_connect_to_wspaceDB.Name = "Button_connect_to_wspaceDB";
            this.Button_connect_to_wspaceDB.Size = new System.Drawing.Size(146, 41);
            this.Button_connect_to_wspaceDB.TabIndex = 2;
            this.Button_connect_to_wspaceDB.Text = "Fill the WS definitions";
            this.Button_connect_to_wspaceDB.UseVisualStyleBackColor = true;
            this.Button_connect_to_wspaceDB.Click += new System.EventHandler(this.Button_connect_to_wspaceDB_Click);
            // 
            // Button_connect_to_atwsDB
            // 
            this.Button_connect_to_atwsDB.Location = new System.Drawing.Point(94, 58);
            this.Button_connect_to_atwsDB.Name = "Button_connect_to_atwsDB";
            this.Button_connect_to_atwsDB.Size = new System.Drawing.Size(137, 42);
            this.Button_connect_to_atwsDB.TabIndex = 2;
            this.Button_connect_to_atwsDB.Text = "Fill the ATWS definitions";
            this.Button_connect_to_atwsDB.UseVisualStyleBackColor = true;
            this.Button_connect_to_atwsDB.Click += new System.EventHandler(this.Button_connect_to_atws_DB_Click);
            // 
            // Button_connect_to_rowDB
            // 
            this.Button_connect_to_rowDB.Location = new System.Drawing.Point(6, 6);
            this.Button_connect_to_rowDB.Name = "Button_connect_to_rowDB";
            this.Button_connect_to_rowDB.Size = new System.Drawing.Size(146, 41);
            this.Button_connect_to_rowDB.TabIndex = 2;
            this.Button_connect_to_rowDB.Text = "Fill the ROW definitions";
            this.Button_connect_to_rowDB.UseVisualStyleBackColor = true;
            this.Button_connect_to_rowDB.Click += new System.EventHandler(this.Button_connect_to_access_DB_Click);
            // 
            // tabPage_CONFIG
            // 
            this.tabPage_CONFIG.BackColor = System.Drawing.Color.Gainsboro;
            this.tabPage_CONFIG.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage_CONFIG.Controls.Add(this.label5);
            this.tabPage_CONFIG.Controls.Add(this.label4);
            this.tabPage_CONFIG.Controls.Add(this.label3);
            this.tabPage_CONFIG.Controls.Add(this.textBox_ATWS_table_name);
            this.tabPage_CONFIG.Controls.Add(this.textBox_WORKSPACE_TABLE_NAME);
            this.tabPage_CONFIG.Controls.Add(this.textBox_row_table_name);
            this.tabPage_CONFIG.ForeColor = System.Drawing.Color.Black;
            this.tabPage_CONFIG.Location = new System.Drawing.Point(4, 27);
            this.tabPage_CONFIG.Name = "tabPage_CONFIG";
            this.tabPage_CONFIG.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_CONFIG.Size = new System.Drawing.Size(512, 248);
            this.tabPage_CONFIG.TabIndex = 1;
            this.tabPage_CONFIG.Text = "Config";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 119);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(109, 15);
            this.label5.TabIndex = 1;
            this.label5.Text = "ATWS Table name";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 92);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(141, 15);
            this.label4.TabIndex = 1;
            this.label4.Text = "Workspace Table name";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 15);
            this.label3.TabIndex = 1;
            this.label3.Text = "ROW Table name";
            // 
            // textBox_ATWS_table_name
            // 
            this.textBox_ATWS_table_name.ForeColor = System.Drawing.Color.Black;
            this.textBox_ATWS_table_name.Location = new System.Drawing.Point(166, 116);
            this.textBox_ATWS_table_name.Name = "textBox_ATWS_table_name";
            this.textBox_ATWS_table_name.Size = new System.Drawing.Size(328, 21);
            this.textBox_ATWS_table_name.TabIndex = 0;
            this.textBox_ATWS_table_name.Text = "ATWS_TABLE";
            // 
            // textBox_WORKSPACE_TABLE_NAME
            // 
            this.textBox_WORKSPACE_TABLE_NAME.ForeColor = System.Drawing.Color.Black;
            this.textBox_WORKSPACE_TABLE_NAME.Location = new System.Drawing.Point(166, 89);
            this.textBox_WORKSPACE_TABLE_NAME.Name = "textBox_WORKSPACE_TABLE_NAME";
            this.textBox_WORKSPACE_TABLE_NAME.Size = new System.Drawing.Size(328, 21);
            this.textBox_WORKSPACE_TABLE_NAME.TabIndex = 0;
            this.textBox_WORKSPACE_TABLE_NAME.Text = "WORKSPACE_SCHEMATIC_TABLE";
            // 
            // textBox_row_table_name
            // 
            this.textBox_row_table_name.ForeColor = System.Drawing.Color.Black;
            this.textBox_row_table_name.Location = new System.Drawing.Point(166, 64);
            this.textBox_row_table_name.Name = "textBox_row_table_name";
            this.textBox_row_table_name.Size = new System.Drawing.Size(328, 21);
            this.textBox_row_table_name.TabIndex = 0;
            this.textBox_row_table_name.Text = "ROW_CONFIG_TABLE";
            // 
            // panel_UP
            // 
            this.panel_UP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel_UP.Controls.Add(this.radioButton_right_UP);
            this.panel_UP.Controls.Add(this.radioButton_left_UP);
            this.panel_UP.Location = new System.Drawing.Point(128, 93);
            this.panel_UP.Name = "panel_UP";
            this.panel_UP.Size = new System.Drawing.Size(114, 57);
            this.panel_UP.TabIndex = 5;
            // 
            // radioButton_right_UP
            // 
            this.radioButton_right_UP.AutoSize = true;
            this.radioButton_right_UP.Location = new System.Drawing.Point(3, 28);
            this.radioButton_right_UP.Name = "radioButton_right_UP";
            this.radioButton_right_UP.Size = new System.Drawing.Size(78, 19);
            this.radioButton_right_UP.TabIndex = 4;
            this.radioButton_right_UP.Text = "RIGHT UP";
            this.radioButton_right_UP.UseVisualStyleBackColor = true;
            // 
            // radioButton_left_UP
            // 
            this.radioButton_left_UP.AutoSize = true;
            this.radioButton_left_UP.Checked = true;
            this.radioButton_left_UP.Location = new System.Drawing.Point(3, 3);
            this.radioButton_left_UP.Name = "radioButton_left_UP";
            this.radioButton_left_UP.Size = new System.Drawing.Size(71, 19);
            this.radioButton_left_UP.TabIndex = 4;
            this.radioButton_left_UP.TabStop = true;
            this.radioButton_left_UP.Text = "LEFT UP";
            this.radioButton_left_UP.UseVisualStyleBackColor = true;
            // 
            // Workspace_band_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(537, 297);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MinimizeBox = false;
            this.Name = "Workspace_band_form";
            this.Text = "Workspace band";
            this.tabControl1.ResumeLayout(false);
            this.tabPage_pick_from_existing.ResumeLayout(false);
            this.tabPage_pick_from_existing.PerformLayout();
            this.panel_left_right.ResumeLayout(false);
            this.panel_left_right.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.panel_matchlines.ResumeLayout(false);
            this.panel_matchlines.PerformLayout();
            this.tabPage_CONFIG.ResumeLayout(false);
            this.tabPage_CONFIG.PerformLayout();
            this.panel_UP.ResumeLayout(false);
            this.panel_UP.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage_CONFIG;
        internal System.Windows.Forms.Button Button_connect_to_rowDB;
        private System.Windows.Forms.TextBox textBox_Matchline_end;
        private System.Windows.Forms.TextBox textBox_Matchline_start;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_access_database_location;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel_matchlines;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_row_table_name;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_WORKSPACE_TABLE_NAME;
        internal System.Windows.Forms.Button Button_connect_to_wspaceDB;
        internal System.Windows.Forms.Button button_DRAW_BAND;
        private System.Windows.Forms.TabPage tabPage_pick_from_existing;
        internal System.Windows.Forms.Button Button_connect_to_atwsDB;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox_ATWS_table_name;
        internal System.Windows.Forms.Button button_draw_from_pick;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_COMPILED;
        internal System.Windows.Forms.Button button_DRAW_COMPILED;
        internal System.Windows.Forms.Button button_pick_CL_EASEMENT_WS;
        private System.Windows.Forms.Panel panel_left_right;
        private System.Windows.Forms.RadioButton radioButton_right_left;
        private System.Windows.Forms.RadioButton radioButton_left_right;
        private System.Windows.Forms.Panel panel_UP;
        private System.Windows.Forms.RadioButton radioButton_right_UP;
        private System.Windows.Forms.RadioButton radioButton_left_UP;
    }
}