using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Filter_Box : Form
    {
        //Global Variables

        string col_cat = "Category";
        string col_select = "Select";
        public static int X = 0;
        public static int Y = 0;
        System.Data.DataTable dt_cat_display = null;
        public Filter_Box()
        {
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(X, Y);

            InitializeComponent();
            dt_cat_display = build_dt_cat_display();
            DataGridViewCheckBoxColumn dgrid_col_select = Mat_Design_form.datagrid_to_datatable_checkbox(dt_cat_display, col_select);
            DataGridViewTextBoxColumn dgrid_col_category = Mat_Design_form.datagrid_to_datatable_textbox(dt_cat_display, col_cat);

            dataGridView_Filter.AutoGenerateColumns = false;
            dataGridView_Filter.Columns.AddRange(dgrid_col_select, dgrid_col_category);

            dataGridView_Filter.DataSource = dt_cat_display;
            dataGridView_Filter.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_Filter.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_Filter.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Padding newpadding = new Padding(4, 0, 0, 0);
            dataGridView_Filter.Columns[0].Width = 40;
            dataGridView_Filter.Columns[1].Width = 129;
            dataGridView_Filter.ColumnHeadersDefaultCellStyle.Padding = newpadding;
            dataGridView_Filter.RowHeadersWidth = 20;
            dataGridView_Filter.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_Filter.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
            dataGridView_Filter.EnableHeadersVisualStyles = false;
            int filter_height = (dataGridView_Filter.RowTemplate.Height * dt_cat_display.Rows.Count) + dataGridView_Filter.ColumnHeadersHeight;

            if (filter_height > 320)
            {
                filter_height = 320;
            }

            dataGridView_Filter.Height = filter_height;
            panel_filter.Height = filter_height + 53;
            this.Height = panel_filter.Height;

        }

        public System.Data.DataTable build_dt_cat_display()
        {
            List<string> lista1 = ds_main.tpage_mat_design.build_category_and_type_list_and_dt_ct();
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(col_select, typeof(bool));
            dt1.Columns.Add(col_cat, typeof(string));

            for (int i = 0; i < lista1.Count; ++i)
            {
                dt1.Rows.Add();
                dt1.Rows[i][0] = false;
                dt1.Rows[i][1] = lista1[i];
            }
            return dt1;
        }

        private void button_Filter_Click(object sender, EventArgs e)
        {
            ds_main.tpage_mat_design.ct_list = new List<string>();
            for (int i = 0; i < dt_cat_display.Rows.Count; ++i)
            {
                if ((bool)dt_cat_display.Rows[i][0] == true)
                {
                    ds_main.tpage_mat_design.ct_list.Add(Convert.ToString(dt_cat_display.Rows[i][1]));
                }
            }
            if (ds_main.tpage_mat_design.ct_list.Count == 0)
            {
                ds_main.tpage_mat_design.ct_list = ds_main.tpage_mat_design.build_category_and_type_list_and_dt_ct();
            }
            ds_main.tpage_mat_design.add_tab_pages();
            this.Close();
        }

        private void checkBox_Filter_Select_All_Click(object sender, EventArgs e)
        {
            if (dt_cat_display != null)
            {
                for (int i = 0; i < dt_cat_display.Rows.Count; ++i)
                {
                    if (checkBox_Filter_Select_All.Checked == true)
                    {
                        dt_cat_display.Rows[i][0] = true;
                    }
                    else
                    {
                        dt_cat_display.Rows[i][0] = false;
                    }
                }
            }
        }

        private void dataGridView_Filter_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            bool all_checked = true;

            for (int i = 0; i < dataGridView_Filter.Rows.Count; ++i)
            {
                if (Convert.ToBoolean(dataGridView_Filter.Rows[i].Cells[0].Value) == false)
                {
                    all_checked = false;
                    i = dataGridView_Filter.Rows.Count;
                }
            }

            if (all_checked == false)
            {
                checkBox_Filter_Select_All.Checked = false;
            }
            else
            {
                checkBox_Filter_Select_All.Checked = true;
            }
        }

    }
}
