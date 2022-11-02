namespace CLTools.UI
{
    partial class Help_Form
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
            this.label_help_message = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel_help_title = new System.Windows.Forms.Panel();
            this.label_help_title = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel_help_title.SuspendLayout();
            this.SuspendLayout();
            // 
            // label_help_message
            // 
            this.label_help_message.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold);
            this.label_help_message.ForeColor = System.Drawing.Color.White;
            this.label_help_message.Location = new System.Drawing.Point(3, 0);
            this.label_help_message.Name = "label_help_message";
            this.label_help_message.Size = new System.Drawing.Size(743, 369);
            this.label_help_message.TabIndex = 0;
            this.label_help_message.Text = "This is the help text";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.label_help_message, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(26, 61);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(749, 369);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // panel_help_title
            // 
            this.panel_help_title.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(201)))));
            this.panel_help_title.Controls.Add(this.label_help_title);
            this.panel_help_title.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_help_title.Location = new System.Drawing.Point(0, 0);
            this.panel_help_title.Name = "panel_help_title";
            this.panel_help_title.Size = new System.Drawing.Size(800, 40);
            this.panel_help_title.TabIndex = 2;
            // 
            // label_help_title
            // 
            this.label_help_title.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_help_title.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_help_title.ForeColor = System.Drawing.Color.White;
            this.label_help_title.Location = new System.Drawing.Point(0, 0);
            this.label_help_title.Name = "label_help_title";
            this.label_help_title.Size = new System.Drawing.Size(800, 40);
            this.label_help_title.TabIndex = 0;
            this.label_help_title.Text = "Help";
            this.label_help_title.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Help_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(37)))), ((int)(((byte)(38)))));
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.panel_help_title);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "Help_Form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Help_Form";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel_help_title.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label_help_message;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel_help_title;
        private System.Windows.Forms.Label label_help_title;
    }
}