namespace Dimensioning
{
    partial class CLTool_Form
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
            this.comboBox_Client = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button_cl_from_all_viewports = new System.Windows.Forms.Button();
            this.button_GenerateCL = new System.Windows.Forms.Button();
            this.label_Client = new System.Windows.Forms.Label();
            this.label_mm = new System.Windows.Forms.Label();
            this.button_Exit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBox_Client
            // 
            this.comboBox_Client.BackColor = System.Drawing.Color.White;
            this.comboBox_Client.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Client.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.comboBox_Client.ForeColor = System.Drawing.Color.Black;
            this.comboBox_Client.FormattingEnabled = true;
            this.comboBox_Client.Location = new System.Drawing.Point(17, 37);
            this.comboBox_Client.Name = "comboBox_Client";
            this.comboBox_Client.Size = new System.Drawing.Size(261, 24);
            this.comboBox_Client.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(28)))), ((int)(((byte)(28)))), ((int)(((byte)(28)))));
            this.panel1.Controls.Add(this.button_cl_from_all_viewports);
            this.panel1.Controls.Add(this.button_GenerateCL);
            this.panel1.Controls.Add(this.label_Client);
            this.panel1.Controls.Add(this.comboBox_Client);
            this.panel1.Location = new System.Drawing.Point(0, 33);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(290, 140);
            this.panel1.TabIndex = 1;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.panel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // button_cl_from_all_viewports
            // 
            this.button_cl_from_all_viewports.BackColor = System.Drawing.Color.Black;
            this.button_cl_from_all_viewports.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.button_cl_from_all_viewports.FlatAppearance.CheckedBackColor = System.Drawing.Color.Black;
            this.button_cl_from_all_viewports.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gray;
            this.button_cl_from_all_viewports.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_cl_from_all_viewports.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_cl_from_all_viewports.ForeColor = System.Drawing.Color.White;
            this.button_cl_from_all_viewports.Location = new System.Drawing.Point(154, 96);
            this.button_cl_from_all_viewports.Name = "button_cl_from_all_viewports";
            this.button_cl_from_all_viewports.Size = new System.Drawing.Size(124, 30);
            this.button_cl_from_all_viewports.TabIndex = 53;
            this.button_cl_from_all_viewports.Text = "CL Dwg";
            this.button_cl_from_all_viewports.UseVisualStyleBackColor = false;
            this.button_cl_from_all_viewports.Click += new System.EventHandler(this.button_cl_from_all_viewports_Click);
            // 
            // button_GenerateCL
            // 
            this.button_GenerateCL.BackColor = System.Drawing.Color.Black;
            this.button_GenerateCL.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.button_GenerateCL.FlatAppearance.CheckedBackColor = System.Drawing.Color.Black;
            this.button_GenerateCL.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gray;
            this.button_GenerateCL.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_GenerateCL.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_GenerateCL.ForeColor = System.Drawing.Color.White;
            this.button_GenerateCL.Location = new System.Drawing.Point(12, 96);
            this.button_GenerateCL.Name = "button_GenerateCL";
            this.button_GenerateCL.Size = new System.Drawing.Size(124, 30);
            this.button_GenerateCL.TabIndex = 53;
            this.button_GenerateCL.Text = "CL Viewport";
            this.button_GenerateCL.UseVisualStyleBackColor = false;
            this.button_GenerateCL.Click += new System.EventHandler(this.button_GenerateCL_Click);
            // 
            // label_Client
            // 
            this.label_Client.AutoSize = true;
            this.label_Client.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Client.ForeColor = System.Drawing.Color.White;
            this.label_Client.Location = new System.Drawing.Point(17, 16);
            this.label_Client.Name = "label_Client";
            this.label_Client.Size = new System.Drawing.Size(89, 16);
            this.label_Client.TabIndex = 46;
            this.label_Client.Text = "Select Client";
            // 
            // label_mm
            // 
            this.label_mm.AutoSize = true;
            this.label_mm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_mm.ForeColor = System.Drawing.Color.White;
            this.label_mm.Location = new System.Drawing.Point(9, 7);
            this.label_mm.Name = "label_mm";
            this.label_mm.Size = new System.Drawing.Size(137, 20);
            this.label_mm.TabIndex = 4;
            this.label_mm.Text = "Mott Macdonald";
            this.label_mm.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.label_mm.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.label_mm.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            // 
            // button_Exit
            // 
            this.button_Exit.BackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(32)))), ((int)(((byte)(40)))));
            this.button_Exit.FlatAppearance.BorderSize = 0;
            this.button_Exit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.ForeColor = System.Drawing.Color.AliceBlue;
            this.button_Exit.Location = new System.Drawing.Point(229, 175);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(46, 24);
            this.button_Exit.TabIndex = 17;
            this.button_Exit.Text = "Exit";
            this.button_Exit.UseVisualStyleBackColor = false;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(7, 179);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 16);
            this.label1.TabIndex = 46;
            this.label1.Text = "V 1.0";
            // 
            // CLTool_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.ClientSize = new System.Drawing.Size(290, 200);
            this.Controls.Add(this.button_Exit);
            this.Controls.Add(this.label_mm);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.Name = "CLTool_Form";
            this.Text = "Centerline Tool";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.clickmove_MouseUp);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox_Client;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button_GenerateCL;
        private System.Windows.Forms.Label label_Client;
        private System.Windows.Forms.Label label_mm;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.Button button_cl_from_all_viewports;
        private System.Windows.Forms.Label label1;
    }
}