namespace Alignment_mdi
{
    partial class OD_TABLE_form
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
            this.DataGrid1 = new System.Windows.Forms.DataGridView();
            this.button_LOAD = new System.Windows.Forms.Button();
            this.Button_Update_object_data = new System.Windows.Forms.Button();
            this.button_add_OD_table_as_layer_name = new System.Windows.Forms.Button();
            this.button_add_OD_table_as_layer_name_entire_drawing = new System.Windows.Forms.Button();
            this.button_go_to_object_data = new System.Windows.Forms.Button();
            this.button_zoom = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DataGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // DataGrid1
            // 
            this.DataGrid1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.DataGrid1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.DataGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGrid1.Dock = System.Windows.Forms.DockStyle.Top;
            this.DataGrid1.GridColor = System.Drawing.Color.LightGray;
            this.DataGrid1.Location = new System.Drawing.Point(0, 0);
            this.DataGrid1.Name = "DataGrid1";
            this.DataGrid1.RowHeadersVisible = false;
            this.DataGrid1.Size = new System.Drawing.Size(797, 342);
            this.DataGrid1.TabIndex = 0;
            this.DataGrid1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGrid1_CellClick);
            this.DataGrid1.CurrentCellChanged += new System.EventHandler(this.DataGrid1_CurrentCellChanged);
            // 
            // button_LOAD
            // 
            this.button_LOAD.Location = new System.Drawing.Point(0, 348);
            this.button_LOAD.Name = "button_LOAD";
            this.button_LOAD.Size = new System.Drawing.Size(191, 56);
            this.button_LOAD.TabIndex = 1;
            this.button_LOAD.Text = "Read Existing Object Data\r\nfrom selected Layer";
            this.button_LOAD.UseVisualStyleBackColor = true;
            this.button_LOAD.Click += new System.EventHandler(this.button_LOAD_Click);
            // 
            // Button_Update_object_data
            // 
            this.Button_Update_object_data.Location = new System.Drawing.Point(0, 403);
            this.Button_Update_object_data.Name = "Button_Update_object_data";
            this.Button_Update_object_data.Size = new System.Drawing.Size(191, 67);
            this.Button_Update_object_data.TabIndex = 2;
            this.Button_Update_object_data.Text = "UPDATE Object Data Table \r\nto the selected Object\r\n(current row)";
            this.Button_Update_object_data.UseVisualStyleBackColor = true;
            this.Button_Update_object_data.Click += new System.EventHandler(this.Button_Update_object_data_Click);
            // 
            // button_add_OD_table_as_layer_name
            // 
            this.button_add_OD_table_as_layer_name.Location = new System.Drawing.Point(406, 390);
            this.button_add_OD_table_as_layer_name.Name = "button_add_OD_table_as_layer_name";
            this.button_add_OD_table_as_layer_name.Size = new System.Drawing.Size(191, 74);
            this.button_add_OD_table_as_layer_name.TabIndex = 3;
            this.button_add_OD_table_as_layer_name.Text = "Assign object data table \r\nto the selected Object\r\n(Values added)";
            this.button_add_OD_table_as_layer_name.UseVisualStyleBackColor = true;
            this.button_add_OD_table_as_layer_name.Click += new System.EventHandler(this.button_add_OD_table_as_layer_name_Click);
            // 
            // button_add_OD_table_as_layer_name_entire_drawing
            // 
            this.button_add_OD_table_as_layer_name_entire_drawing.Location = new System.Drawing.Point(603, 348);
            this.button_add_OD_table_as_layer_name_entire_drawing.Name = "button_add_OD_table_as_layer_name_entire_drawing";
            this.button_add_OD_table_as_layer_name_entire_drawing.Size = new System.Drawing.Size(188, 116);
            this.button_add_OD_table_as_layer_name_entire_drawing.TabIndex = 4;
            this.button_add_OD_table_as_layer_name_entire_drawing.Text = "ADD Object Data Table \r\nto the entire drawing \r\nbased on the layer name\r\n(no valu" +
    "es added)";
            this.button_add_OD_table_as_layer_name_entire_drawing.UseVisualStyleBackColor = true;
            this.button_add_OD_table_as_layer_name_entire_drawing.Click += new System.EventHandler(this.button_add_OD_table_as_layer_name_entire_drawing_Click);
            // 
            // button_go_to_object_data
            // 
            this.button_go_to_object_data.Location = new System.Drawing.Point(209, 348);
            this.button_go_to_object_data.Name = "button_go_to_object_data";
            this.button_go_to_object_data.Size = new System.Drawing.Size(191, 56);
            this.button_go_to_object_data.TabIndex = 5;
            this.button_go_to_object_data.Text = "Make object OD values\r\ncurrent row";
            this.button_go_to_object_data.UseVisualStyleBackColor = true;
            this.button_go_to_object_data.Click += new System.EventHandler(this.button_go_to_object_data_Click);
            // 
            // button_zoom
            // 
            this.button_zoom.Location = new System.Drawing.Point(209, 410);
            this.button_zoom.Name = "button_zoom";
            this.button_zoom.Size = new System.Drawing.Size(191, 54);
            this.button_zoom.TabIndex = 6;
            this.button_zoom.Text = "Zoom TO";
            this.button_zoom.UseVisualStyleBackColor = true;
            this.button_zoom.Click += new System.EventHandler(this.button_zoom_Click);
            // 
            // OD_TABLE_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(797, 469);
            this.Controls.Add(this.button_zoom);
            this.Controls.Add(this.button_go_to_object_data);
            this.Controls.Add(this.button_add_OD_table_as_layer_name_entire_drawing);
            this.Controls.Add(this.button_add_OD_table_as_layer_name);
            this.Controls.Add(this.Button_Update_object_data);
            this.Controls.Add(this.button_LOAD);
            this.Controls.Add(this.DataGrid1);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "OD_TABLE_form";
            this.Text = "Object Data Table";
            ((System.ComponentModel.ISupportInitialize)(this.DataGrid1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DataGrid1;
        private System.Windows.Forms.Button button_LOAD;
        private System.Windows.Forms.Button Button_Update_object_data;
        private System.Windows.Forms.Button button_add_OD_table_as_layer_name;
        private System.Windows.Forms.Button button_add_OD_table_as_layer_name_entire_drawing;
        private System.Windows.Forms.Button button_go_to_object_data;
        private System.Windows.Forms.Button button_zoom;
    }
}