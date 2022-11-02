namespace weld_sheet
{
    partial class Form_main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_main));
            this.label_mm = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_Exit = new System.Windows.Forms.Button();
            this.button_minimize = new System.Windows.Forms.Button();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.comboBox_wm = new System.Windows.Forms.ComboBox();
            this.button_load_settings = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_extra = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.button_Export_Field_Notes = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // label_mm
            // 
            this.label_mm.AutoSize = true;
            this.label_mm.Dock = System.Windows.Forms.DockStyle.Left;
            this.label_mm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_mm.ForeColor = System.Drawing.Color.White;
            this.label_mm.Location = new System.Drawing.Point(3, 0);
            this.label_mm.Name = "label_mm";
            this.label_mm.Size = new System.Drawing.Size(137, 39);
            this.label_mm.TabIndex = 3;
            this.label_mm.Text = "Mott Macdonald";
            this.label_mm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label_mm.Click += new System.EventHandler(this.label_mm_Click);
            this.label_mm.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.label_mm.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.label_mm.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 151F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 42F));
            this.tableLayoutPanel1.Controls.Add(this.label_mm, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_Exit, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_minimize, 2, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(511, 39);
            this.tableLayoutPanel1.TabIndex = 2109;
            this.tableLayoutPanel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.tableLayoutPanel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.tableLayoutPanel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // button_Exit
            // 
            this.button_Exit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_Exit.BackgroundImage")));
            this.button_Exit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_Exit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Exit.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_Exit.FlatAppearance.BorderSize = 0;
            this.button_Exit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Exit.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.ForeColor = System.Drawing.Color.White;
            this.button_Exit.Location = new System.Drawing.Point(472, 3);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(36, 33);
            this.button_Exit.TabIndex = 161;
            this.button_Exit.UseVisualStyleBackColor = false;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // button_minimize
            // 
            this.button_minimize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_minimize.BackgroundImage")));
            this.button_minimize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_minimize.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_minimize.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_minimize.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.button_minimize.FlatAppearance.BorderSize = 0;
            this.button_minimize.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_minimize.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_minimize.ForeColor = System.Drawing.Color.White;
            this.button_minimize.Location = new System.Drawing.Point(434, 3);
            this.button_minimize.Name = "button_minimize";
            this.button_minimize.Size = new System.Drawing.Size(32, 33);
            this.button_minimize.TabIndex = 162;
            this.button_minimize.UseVisualStyleBackColor = false;
            this.button_minimize.Click += new System.EventHandler(this.button_minimize_Click);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 74F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.comboBox_wm, 1, 1);
            this.tableLayoutPanel2.Controls.Add(this.button_load_settings, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.textBox_extra, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 39);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(511, 60);
            this.tableLayoutPanel2.TabIndex = 2114;
            // 
            // comboBox_wm
            // 
            this.comboBox_wm.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.comboBox_wm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox_wm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_wm.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_wm.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.comboBox_wm.ForeColor = System.Drawing.Color.White;
            this.comboBox_wm.FormattingEnabled = true;
            this.comboBox_wm.Location = new System.Drawing.Point(77, 29);
            this.comboBox_wm.Name = "comboBox_wm";
            this.comboBox_wm.Size = new System.Drawing.Size(431, 24);
            this.comboBox_wm.TabIndex = 2203;
            this.comboBox_wm.DropDown += new System.EventHandler(this.comboBox_Pipe_DropDown);
            // 
            // button_load_settings
            // 
            this.button_load_settings.BackColor = System.Drawing.Color.White;
            this.button_load_settings.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button_load_settings.BackgroundImage")));
            this.button_load_settings.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button_load_settings.Dock = System.Windows.Forms.DockStyle.Right;
            this.button_load_settings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.button_load_settings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_load_settings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_load_settings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_load_settings.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_load_settings.ForeColor = System.Drawing.Color.White;
            this.button_load_settings.Location = new System.Drawing.Point(487, 3);
            this.button_load_settings.Name = "button_load_settings";
            this.button_load_settings.Size = new System.Drawing.Size(21, 20);
            this.button_load_settings.TabIndex = 2205;
            this.button_load_settings.UseVisualStyleBackColor = false;
            this.button_load_settings.Click += new System.EventHandler(this.button_load_settings_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(3, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 38);
            this.label2.TabIndex = 2116;
            this.label2.Text = "Weld Map";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBox_extra
            // 
            this.textBox_extra.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.textBox_extra.ForeColor = System.Drawing.Color.White;
            this.textBox_extra.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.textBox_extra.Location = new System.Drawing.Point(3, 3);
            this.textBox_extra.Name = "textBox_extra";
            this.textBox_extra.Size = new System.Drawing.Size(68, 20);
            this.textBox_extra.TabIndex = 2206;
            this.textBox_extra.Text = "0";
            this.textBox_extra.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_extra.Visible = false;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 3;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel4.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.button_Export_Field_Notes, 2, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(0, 99);
            this.tableLayoutPanel4.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(511, 33);
            this.tableLayoutPanel4.TabIndex = 2115;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(3, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(315, 13);
            this.label1.TabIndex = 2205;
            this.label1.Text = "v-1.1.0";
            // 
            // button_Export_Field_Notes
            // 
            this.button_Export_Field_Notes.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            this.button_Export_Field_Notes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Export_Field_Notes.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.button_Export_Field_Notes.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button_Export_Field_Notes.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange;
            this.button_Export_Field_Notes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Export_Field_Notes.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold);
            this.button_Export_Field_Notes.ForeColor = System.Drawing.Color.White;
            this.button_Export_Field_Notes.Image = ((System.Drawing.Image)(resources.GetObject("button_Export_Field_Notes.Image")));
            this.button_Export_Field_Notes.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_Export_Field_Notes.Location = new System.Drawing.Point(362, 3);
            this.button_Export_Field_Notes.Name = "button_Export_Field_Notes";
            this.button_Export_Field_Notes.Size = new System.Drawing.Size(144, 28);
            this.button_Export_Field_Notes.TabIndex = 2204;
            this.button_Export_Field_Notes.Text = "Export to Excel";
            this.button_Export_Field_Notes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_Export_Field_Notes.UseVisualStyleBackColor = false;
            this.button_Export_Field_Notes.Click += new System.EventHandler(this.button_Export_weld_sheets_Click);
            // 
            // Form_main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.ClientSize = new System.Drawing.Size(511, 135);
            this.Controls.Add(this.tableLayoutPanel4);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MinimumSize = new System.Drawing.Size(300, 130);
            this.Name = "Form_main";
            this.Padding = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.Text = "Form1";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button_minimize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.ComboBox comboBox_wm;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.Button button_Export_Field_Notes;
        private System.Windows.Forms.Button button_load_settings;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_extra;
    }
}

