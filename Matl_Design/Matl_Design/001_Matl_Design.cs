using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Mat_Design_form : Form
    {
        //Global Variables
        public System.Data.DataTable dt_mat_library = null;
        public System.Data.DataTable dt_filter = null;
        public System.Data.DataTable[] dt_ct = null;

        public List<string> ct_list = null;

        string col_mmid = "MMID";
        string col_item_no = "ItemNo";

        string col_category = "Category";
        string col_type = "Type";
        string col_layer = "Layer";


        string col_MSblock = "MS Block";

        int ft_mat_library = 0;
        int ft_pipe = 0;





        List<TabPage> tabPages_list = null;

        int nr_max = 150000;



        string pipe_col_mmid = "MMID";
        string pipe_col_2d1 = "2DStaBeg";
        string pipe_col_2d2 = "2DStaEnd";
        string pipe_col_3d1 = "3DStaBeg";
        string pipe_col_3d2 = "3DStaEnd";
        string pipe_col_eq1 = "EqStaBeg";
        string pipe_col_eq2 = "EqStaEnd";

        string pipe_col_len2d = "2D len";
        string pipe_col_len3d = "3D len";

        string pipe_col11 = "X_Beg";
        string pipe_col12 = "Y_Beg";
        string pipe_col13 = "X_End";
        string pipe_col14 = "Y_End";

        string pipe_col_block = "BLOCK";
        string pipe_col_mat = "MAT";
        string pipe_col17 = "ATR2";
        string pipe_col18 = "ATR3";
        string pipe_col19 = "ATR4";
        string pipe_col20 = "Visibility";

        string col_x = "X";
        string col_y = "Y";


        string col_atr1 = "ATR1";
        string col_atr2 = "ATR2";
        string col_atr3 = "ATR3";
        string col_atr4 = "ATR4";
        string col_visibility = "Visibility";
        string col_elbow = "ELBOW";
        string col_mat_elbow = "Elbow Item No";

        string pipes_layer = "Pipes";
        string pipes_od = "Pipes";
        string points_od = "MatPoints";
        string extra_od = "LinearOther";









        string pipe_us_od_item_no = "ItemNo";
        string pipe_us_od_descr = "Description";
        string pipe_us_od_cat = "Category";
        string pipe_us_od_sta1 = "BeginSta";
        string pipe_us_od_sta2 = "EndSta";

        string col_elbow_ref_id = "Reference_id";

        string col_ref_dwg_id = "Reference dwg";
        string col_od_ref_dwg_id = "Ref_dwg";
        string col_od_min_depth = "Min_cvr";
        string col_min_depth = "Minimum cover";
        string col_agen_cvr = "Agen_cvr";
        string col_xing_method = "Crossing_method";
        string col_od_xing_method = "XingMethod";





        string col_sta = "STA";
        string col_descr = "Descr";




        System.Drawing.Font font10 = null;
        System.Drawing.Font font8 = null;



        string extra_layer_to_be_deleted = "";

        string col_2dbeg = "2DStaBeg";
        string col_2dsta = "2DSta";
        string col_2dend = "2DStaEnd";
        string col_3dbeg = "3DStaBeg";
        string col_3dsta = "3DSta";
        string col_3dend = "3DStaEnd";
        string col_eqStabeg = "EqStaBeg";
        string col_eqsta = "EqSta";
        string col_eqStaend = "EqStaEnd";
        string col_2dlen = "2D Length";
        string col_3dlen = "3D Length";
        string col_altdesc = "AltDesc";
        string col_symbol = "Symbol";
        string col_xbeg = "X_Beg";
        string col_ybeg = "Y_Beg";
        string col_xend = "X_End";
        string col_yend = "Y_End";

        string col_z = "Z";
        string col_block = "BLOCK";
        string col_blockdescr = "DESCR";
        string col_note1 = "NOTE1";
        string col_qty = "QTY";

        string col_mat = "MAT";

        string col_id = "ID";
        string col_id2 = "ID2";
        string col_cvr = "CVR";



        private ContextMenuStrip ContextMenuStrip_dt_mat_lib;
        private ContextMenuStrip ContextMenuStrip_dt_pipe;
        private ContextMenuStrip ContextMenuStrip_dt_pt;
        private ContextMenuStrip ContextMenuStrip_dt_extra;
        string current_ct = "";

        public Mat_Design_form()
        {
            InitializeComponent();
            ds_main.tpage_mat_design = this;
            font10 = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            font8 = new System.Drawing.Font("Arial", 8f, FontStyle.Bold);
            tabPages_list = new List<TabPage>();
            for (int i = 0; i < flatTabControl1.TabPages.Count; ++i)
            {
                if (i > 0)
                {
                    tabPages_list.Add(flatTabControl1.TabPages[i]);
                }
            }

            flatTabControl1.TabPages.RemoveAt(3);
            flatTabControl1.TabPages.RemoveAt(2);
            flatTabControl1.TabPages.RemoveAt(1);


            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Delete Row" };
            toolStripMenuItem1.Click += delete_current_row_mat_lib_Click;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Clear All Data" };
            toolStripMenuItem2.Click += clear_all_data_mat_lib_Click;

            ContextMenuStrip_dt_mat_lib = new ContextMenuStrip();
            ContextMenuStrip_dt_mat_lib.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1 });//, toolStripMenuItem2 });

            var toolStripMenuItem3 = new ToolStripMenuItem { Text = "Delete Row" };
            toolStripMenuItem3.Click += delete_current_row_dt_pipe_Click;


            ContextMenuStrip_dt_pipe = new ContextMenuStrip();
            ContextMenuStrip_dt_pipe.Items.AddRange(new ToolStripItem[] { toolStripMenuItem3 });

            var toolStripMenuItem4 = new ToolStripMenuItem { Text = "Delete Row" };
            toolStripMenuItem4.Click += delete_current_row_dt_pt_Click;

            ContextMenuStrip_dt_pt = new ContextMenuStrip();
            ContextMenuStrip_dt_pt.Items.AddRange(new ToolStripItem[] { toolStripMenuItem4 });

            var toolStripMenuItem5 = new ToolStripMenuItem { Text = "Delete Row" };
            toolStripMenuItem5.Click += delete_current_row_dt_extra_Click;

            ContextMenuStrip_dt_extra = new ContextMenuStrip();
            ContextMenuStrip_dt_extra.Items.AddRange(new ToolStripItem[] { toolStripMenuItem5 });


        }

        private void delete_current_row_dt_pt_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Delete row", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {

                    DataGridView dgv1 = null;


                    foreach (TabPage tab1 in flatTabControl1.TabPages)
                    {
                        foreach (Panel panel1 in tab1.Controls)
                        {
                            foreach (Control ctrl1 in panel1.Controls)
                            {
                                DataGridView datagrid1 = ctrl1 as DataGridView;
                                if (datagrid1 != null && datagrid1.Name.Replace("dgv_", "").ToUpper() == current_ct)
                                {
                                    dgv1 = datagrid1;
                                }
                            }
                        }
                    }


                    if (dgv1 != null && dgv1.RowCount > 0)
                    {
                        int index_pt = dgv1.CurrentCell.RowIndex;
                        if (index_pt == -1)
                        {
                            return;
                        }


                        string mat1 = Convert.ToString(dgv1.Rows[index_pt].Cells[0].Value);
                        double sta1 = Convert.ToDouble(dgv1.Rows[index_pt].Cells[1].Value);

                        dgv1.Rows.Remove(dgv1.Rows[index_pt]);

                        for (int i = ds_main.dt_points.Rows.Count - 1; i >= 0; --i)
                        {
                            if (ds_main.dt_points.Rows[i][col_item_no] != DBNull.Value)
                            {
                                string mat2 = Convert.ToString(ds_main.dt_points.Rows[i][col_item_no]);

                                if (mat2.ToUpper() == mat1.ToUpper())
                                {
                                    double sta2 = -1;

                                    if (ds_main.dt_points.Rows[i][col_2dsta] != DBNull.Value)
                                    {
                                        sta2 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_2dsta]);
                                    }
                                    if (ds_main.dt_points.Rows[i][col_3dsta] != DBNull.Value)
                                    {
                                        sta2 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_3dsta]);
                                    }

                                    if (Math.Round(sta1, 2) == Math.Round(sta2, 2))
                                    {
                                        ds_main.dt_points.Rows[i].Delete();
                                    }

                                }
                            }
                        }

                        if (ct_list.Contains(current_ct) == true)
                        {
                            int index1 = ct_list.IndexOf(current_ct);

                            System.Data.DataTable dt1 = dt_ct[index1];

                            for (int k = dt1.Rows.Count - 1; k >= 0; --k)
                            {
                                if (dt1.Rows[k][col_item_no] != DBNull.Value)
                                {
                                    string mat2 = Convert.ToString(dt1.Rows[k][col_item_no]);

                                    if (mat2.ToUpper() == mat1.ToUpper())
                                    {
                                        double sta2 = -1;

                                        if (dt1.Rows[k][col_2dsta] != DBNull.Value)
                                        {
                                            sta2 = Convert.ToDouble(dt1.Rows[k][col_2dsta]);
                                        }
                                        if (dt1.Rows[k][col_3dsta] != DBNull.Value)
                                        {
                                            sta2 = Convert.ToDouble(dt1.Rows[k][col_3dsta]);
                                        }

                                        if (Math.Round(sta1, 2) == Math.Round(sta2, 2))
                                        {
                                            dt1.Rows[k].Delete();
                                        }

                                    }
                                }
                            }

                            if (current_ct == "ELL_POINT")
                            {


                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



            }



        }

        private void delete_current_row_mat_lib_Click(object sender, EventArgs e)
        {



            if (comboBox_default_mat.Text == "")
            {
                MessageBox.Show("No default material selected/r/nOperation aborted");
                return;
            }

            string mat0 = comboBox_default_mat.Text;

            try
            {

                if (dataGridView_mat_library.RowCount > 0)
                {
                    int index_grid_matl = dataGridView_mat_library.CurrentCell.RowIndex;
                    if (index_grid_matl == -1)
                    {
                        return;
                    }


                    string mat1 = "";
                    if (dataGridView_mat_library.Rows[index_grid_matl].Cells[col_item_no].Value != DBNull.Value)
                    {
                        mat1 = Convert.ToString(dataGridView_mat_library.Rows[index_grid_matl].Cells[col_item_no].Value);
                    }

                    if (mat0 == mat1)
                    {
                        MessageBox.Show("You can't delete the default material/r/nOperation aborted");
                        return;
                    }

                    if (mat1 == "")
                    {
                        MessageBox.Show("no material selected/r/nOperation aborted");
                        return;
                    }

                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        if (dt_mat_library.Rows[i][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[i][col_item_no]) == mat1)
                        {
                            if (dt_mat_library.Rows[i][col_layer] != DBNull.Value)
                            {
                                extra_layer_to_be_deleted = Convert.ToString(dt_mat_library.Rows[i][col_layer]);
                            }

                            dt_mat_library.Rows[i].Delete();
                        }
                    }


                    bool is3D = ds_main.is3D;


                    if (ds_main.dt_pipe != null && ds_main.dt_pipe.Rows.Count > 0)
                    {
                        for (int i = ds_main.dt_pipe.Rows.Count - 1; i >= 0; --i)
                        {

                            if (ds_main.dt_pipe.Rows[i][col_item_no] != DBNull.Value && Convert.ToString(ds_main.dt_pipe.Rows[i][col_item_no]) == mat1)
                            {
                                ds_main.dt_pipe.Rows[i][col_item_no] = mat0;


                            }
                        }
                        if (is3D == true)
                        {
                            ds_main.dt_pipe = Functions.Sort_data_table(ds_main.dt_pipe, pipe_col_3d1);
                        }
                        else
                        {
                            ds_main.dt_pipe = Functions.Sort_data_table(ds_main.dt_pipe, pipe_col_2d1);
                        }

                        string prev_mat = Convert.ToString(ds_main.dt_pipe.Rows[ds_main.dt_pipe.Rows.Count - 1][col_item_no]);
                        for (int i = ds_main.dt_pipe.Rows.Count - 2; i >= 0; --i)
                        {
                            string mat3 = Convert.ToString(ds_main.dt_pipe.Rows[i][col_item_no]);

                            if (mat3 == prev_mat)
                            {
                                if (is3D == true)
                                {
                                    ds_main.dt_pipe.Rows[i + 1][pipe_col_3d1] = ds_main.dt_pipe.Rows[i][pipe_col_3d1];
                                }
                                else
                                {
                                    ds_main.dt_pipe.Rows[i + 1][pipe_col_2d1] = ds_main.dt_pipe.Rows[i][pipe_col_2d1];
                                }
                                ds_main.dt_pipe.Rows[i].Delete();

                            }
                            prev_mat = mat3;
                        }
                        draw_pipes();
                    }
                    sync_dt_mat_with_filter();
                    add_tab_pages();
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void delete_current_row_dt_extra_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Delete row", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {

                    DataGridView dgv1 = null;


                    foreach (TabPage tab1 in flatTabControl1.TabPages)
                    {
                        foreach (Panel panel1 in tab1.Controls)
                        {
                            foreach (Control ctrl1 in panel1.Controls)
                            {
                                DataGridView datagrid1 = ctrl1 as DataGridView;
                                if (datagrid1 != null && datagrid1.Name.Replace("dgv_", "").ToUpper() == current_ct)
                                {
                                    dgv1 = datagrid1;
                                }
                            }
                        }
                    }


                    if (dgv1 != null && dgv1.RowCount > 0)
                    {
                        int index_row_lin = dgv1.CurrentCell.RowIndex;
                        if (index_row_lin == -1)
                        {
                            return;
                        }


                        string mat1 = Convert.ToString(dgv1.Rows[index_row_lin].Cells[0].Value);
                        double sta1 = Convert.ToDouble(dgv1.Rows[index_row_lin].Cells[1].Value);
                        double sta2 = Convert.ToDouble(dgv1.Rows[index_row_lin].Cells[2].Value);

                        dgv1.Rows.Remove(dgv1.Rows[index_row_lin]);

                        for (int i = ds_main.dt_extra.Rows.Count - 1; i >= 0; --i)
                        {
                            if (ds_main.dt_extra.Rows[i][col_item_no] != DBNull.Value)
                            {
                                string mat2 = Convert.ToString(ds_main.dt_extra.Rows[i][col_item_no]);

                                if (mat2.ToUpper() == mat1.ToUpper())
                                {
                                    double sta3 = -1;
                                    double sta4 = -1;

                                    if (ds_main.dt_extra.Rows[i][col_2dbeg] != DBNull.Value)
                                    {
                                        sta3 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_2dbeg]);
                                    }
                                    if (ds_main.dt_extra.Rows[i][col_3dbeg] != DBNull.Value)
                                    {
                                        sta3 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_3dbeg]);
                                    }

                                    if (ds_main.dt_extra.Rows[i][col_2dend] != DBNull.Value)
                                    {
                                        sta4 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_2dend]);
                                    }
                                    if (ds_main.dt_extra.Rows[i][col_3dend] != DBNull.Value)
                                    {
                                        sta4 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_3dend]);
                                    }

                                    if (Math.Round(sta1, 2) == Math.Round(sta3, 2) && Math.Round(sta2, 2) == Math.Round(sta4, 2))
                                    {
                                        ds_main.dt_extra.Rows[i].Delete();
                                    }

                                }
                            }
                        }

                        if (ct_list.Contains(current_ct) == true)
                        {
                            int index1 = ct_list.IndexOf(current_ct);

                            System.Data.DataTable dt1 = dt_ct[index1];

                            for (int k = dt1.Rows.Count - 1; k >= 0; --k)
                            {
                                if (dt1.Rows[k][col_item_no] != DBNull.Value)
                                {
                                    string mat2 = Convert.ToString(dt1.Rows[k][col_item_no]);

                                    if (mat2.ToUpper() == mat1.ToUpper())
                                    {
                                        double sta3 = -1;
                                        double sta4 = -1;

                                        if (dt1.Rows[k][col_2dbeg] != DBNull.Value)
                                        {
                                            sta3 = Convert.ToDouble(dt1.Rows[k][col_2dbeg]);
                                        }
                                        if (dt1.Rows[k][col_3dbeg] != DBNull.Value)
                                        {
                                            sta3 = Convert.ToDouble(dt1.Rows[k][col_3dbeg]);
                                        }

                                        if (dt1.Rows[k][col_2dend] != DBNull.Value)
                                        {
                                            sta4 = Convert.ToDouble(dt1.Rows[k][col_2dend]);
                                        }
                                        if (dt1.Rows[k][col_3dend] != DBNull.Value)
                                        {
                                            sta4 = Convert.ToDouble(dt1.Rows[k][col_3dend]);
                                        }

                                        if (Math.Round(sta1, 2) == Math.Round(sta3, 2) && Math.Round(sta2, 2) == Math.Round(sta4, 2))
                                        {
                                            dt1.Rows[k].Delete();
                                        }

                                    }
                                }
                            }

                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



            }



        }

        private void sync_dt_mat_with_filter()
        {
            if (dt_mat_library != null && dt_mat_library.Rows.Count > 0 && dt_filter != null && dt_filter.Rows.Count > 0)
            {
                for (int i = 0; i < dt_filter.Rows.Count; ++i)
                {
                    string mmid1 = Convert.ToString(dt_filter.Rows[i][col_mmid]);
                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                    {
                        string mmid2 = Convert.ToString(dt_mat_library.Rows[j][col_mmid]);
                        if (mmid1 == mmid2)
                        {
                            dt_mat_library.Rows[j].ItemArray = dt_filter.Rows[i].ItemArray;
                            j = dt_mat_library.Rows.Count;
                        }
                    }
                }
            }
        }




        public void add_tab_pages()
        {
            for (int i = flatTabControl1.TabPages.Count - 1; i > 0; --i)
            {
                flatTabControl1.TabPages.RemoveAt(i);
            }


            for (int i = 0; i < ct_list.Count; ++i)
            {
                string ct1 = ct_list[i].ToUpper().Replace(" ", "");

                System.Data.DataTable dt1 = dt_ct[ct_list.IndexOf(ct1)];



                List<string> lista_mat = new List<string>();


                if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
                {
                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                    {
                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value)
                        {
                            string mat1 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]);
                            string category2 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                            string type2 = Convert.ToString(dt_mat_library.Rows[j][col_type]);

                            string ct2 = (category2.ToUpper() + "_" + type2.ToUpper()).Replace(" ", "");

                            if (ct1 == ct2)
                            {
                                if (lista_mat.Contains(mat1) == false) lista_mat.Add(mat1);
                            }
                        }
                    }
                }

                if (ct1.ToLower() == "pipe_linear")
                {
                    flatTabControl1.TabPages.Add(tabPages_list[0]);
                }
                else if (ct1.ToUpper().Contains("_POINT") == true)
                {
                    TabPage tab_pt = tabPages_list[1];

                    TabPage tp = create_point_tab(tab_pt, lista_mat, dt1);
                    flatTabControl1.TabPages.Add(tp);
                }
                else if (ct1.ToUpper().Contains("_LINEAR") == true)
                {
                    TabPage tab_lin = tabPages_list[2];
                    TabPage tp = create_linear_tab(tab_lin, lista_mat, dt1);
                    flatTabControl1.TabPages.Add(tp);
                }

            }
            dt_filter = dt_mat_library.Copy();
            for (int i = dt_filter.Rows.Count - 1; i >= 0; --i)
            {
                if (dt_filter.Rows[i][col_category] != DBNull.Value && dt_filter.Rows[i][col_type] != DBNull.Value)
                {
                    string cat1 = Convert.ToString(dt_filter.Rows[i][col_category]).ToUpper().Replace(" ", "");
                    string type1 = Convert.ToString(dt_filter.Rows[i][col_type]).ToUpper().Replace(" ", "");


                    if (cat1.Replace(" ", "").Length > 0)
                    {
                        if (ct_list.Contains(cat1 + "_" + type1) == false)
                        {
                            dt_filter.Rows[i].Delete();
                        }
                    }
                }
            }

            add_linear_mat_to_combobox(comboBox_default_mat);
            add_linear_mat_to_combobox(comboBox_pipe_mat);

            display_mat_lib_on_dgv();



        }

        private void display_mat_lib_on_dgv()
        {
            DataGridViewTextBoxColumn DG_Col_Item_No = datagrid_to_datatable_textbox(dt_filter, col_item_no);
            DataGridViewTextBoxColumn DG_Col_Desc = datagrid_to_datatable_textbox(dt_filter, col_descr);
            DataGridViewTextBoxColumn DG_Col_Cat = datagrid_to_datatable_textbox(dt_filter, col_category);
            DataGridViewTextBoxColumn DG_Col_Type = datagrid_to_datatable_textbox(dt_filter, col_type);
            DataGridViewTextBoxColumn DG_col_layer = datagrid_to_datatable_textbox(dt_filter, col_layer);
            DataGridViewTextBoxColumn DG_Col_Block = datagrid_to_datatable_textbox(dt_filter, col_MSblock);

            if (ft_mat_library == 0)
            {
                dataGridView_mat_library.Columns.AddRange(DG_Col_Item_No, DG_Col_Desc, DG_Col_Cat, DG_Col_Type, DG_col_layer, DG_Col_Block);//, DG_Col_OD_Table, DG_Col_Field);
                dataGridView_mat_library.Columns[0].Name = col_item_no;
                dataGridView_mat_library.Columns[1].Name = col_descr;
                dataGridView_mat_library.Columns[2].Name = col_category;
                dataGridView_mat_library.Columns[3].Name = col_type;
                dataGridView_mat_library.Columns[4].Name = col_layer;
                dataGridView_mat_library.Columns[5].Name = col_MSblock;
                ft_mat_library = 1;
            }

            dataGridView_mat_library.AutoGenerateColumns = false;
            dataGridView_mat_library.DataSource = dt_filter;
            dataGridView_mat_library.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            //DG_Col_OD_Table.FlatStyle = FlatStyle.Flat;
            //DG_Col_OD_Table.Width = 100;
            dataGridView_mat_library.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_mat_library.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_mat_library.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Padding newpadding = new Padding(4, 0, 0, 0);
            dataGridView_mat_library.ColumnHeadersDefaultCellStyle.Padding = newpadding;
            dataGridView_mat_library.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_mat_library.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
            dataGridView_mat_library.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_mat_library.EnableHeadersVisualStyles = false;
            set_textBox_library_content();

        }


        public void set_textBox_library_content()
        {
            textBox_library.Text = ds_main.config_xls;
            textBox_library.ForeColor = Color.LightGreen;
            textBox_library.Font = font8;
            textBox_library.TextAlign = HorizontalAlignment.Right;
        }


        public TabPage create_point_tab(TabPage tab_pt, List<string> list_of_materials, System.Data.DataTable dt1 = null)
        {

            string category1 = "**";
            string type1 = "**";
            string descr = "**";

            for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
            {
                if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                {
                    string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper().Replace(" ", "");
                    if (list_of_materials[0] == mat2)
                    {
                        category1 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                        type1 = Convert.ToString(dt_mat_library.Rows[j][col_type]);
                        descr = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                    }
                }
            }


            TabPage tp = new TabPage();
            tp.BackColor = tab_pt.BackColor;
            tp.ForeColor = tab_pt.ForeColor;
            tp.Font = tab_pt.Font;
            tp.Text = category1;

            Panel pan_top = new Panel();
            pan_top.Bounds = panel_pt_top.Bounds;
            pan_top.ForeColor = panel_pt_top.ForeColor;
            pan_top.BackColor = panel_pt_top.BackColor;
            pan_top.BorderStyle = BorderStyle.FixedSingle;
            pan_top.Dock = DockStyle.Top;


            System.Windows.Forms.Label lab_pt = new System.Windows.Forms.Label();
            lab_pt.Text = category1 + " Material Data";
            lab_pt.Font = label_point.Font;
            lab_pt.Location = label_point.Location;
            lab_pt.BackColor = label_point.BackColor;
            lab_pt.ForeColor = label_point.ForeColor;
            lab_pt.Dock = label_point.Dock;
            lab_pt.AutoSize = true;
            pan_top.Controls.AddRange(new Control[] { lab_pt });

            Panel pan_bottom = new Panel();
            pan_bottom.Bounds = panel_pt_bottom.Bounds;
            pan_bottom.ForeColor = panel_pt_bottom.ForeColor;
            pan_bottom.BackColor = panel_pt_bottom.BackColor;
            pan_bottom.BorderStyle = BorderStyle.FixedSingle;
            pan_bottom.Dock = DockStyle.Bottom;

            System.Windows.Forms.Button btn_save_pt = new System.Windows.Forms.Button();
            btn_save_pt.FlatStyle = button_pt_save.FlatStyle;
            btn_save_pt.FlatAppearance.BorderColor = button_pt_save.FlatAppearance.BorderColor;
            btn_save_pt.FlatAppearance.MouseDownBackColor = button_pt_save.FlatAppearance.MouseDownBackColor;
            btn_save_pt.FlatAppearance.MouseOverBackColor = button_pt_save.FlatAppearance.MouseOverBackColor;
            btn_save_pt.BackgroundImage = button_pt_save.BackgroundImage;
            btn_save_pt.BackgroundImageLayout = button_pt_save.BackgroundImageLayout;
            btn_save_pt.Bounds = button_pt_save.Bounds;
            btn_save_pt.Dock = button_pt_save.Dock;
            btn_save_pt.Name = "button_save_" + category1 + "_" + type1;
            btn_save_pt.Click += new EventHandler(button_save_point_Click);
            pan_bottom.Controls.AddRange(new Control[] { btn_save_pt });


            Panel pan_left = new Panel();
            pan_left.Location = new System.Drawing.Point(0, pan_top.Bottom);
            pan_left.Size = new Size(panel_pt_left.Width, pan_bottom.Top - pan_top.Bottom);
            pan_left.ForeColor = panel_pt_left.ForeColor;
            pan_left.BackColor = panel_pt_left.BackColor;
            pan_left.BorderStyle = BorderStyle.FixedSingle;

            System.Windows.Forms.Label lab_pt_left = new System.Windows.Forms.Label();
            lab_pt_left.Text = descr;
            lab_pt_left.Font = label_pt_left.Font;
            lab_pt_left.Location = label_pt_left.Location;
            lab_pt_left.BackColor = label_pt_left.BackColor;
            lab_pt_left.ForeColor = label_pt_left.ForeColor;
            lab_pt_left.AutoSize = true;
            lab_pt_left.Name = "lab_pt_left_" + category1 + "_" + type1;

            System.Windows.Forms.Label lab_pt_current = new System.Windows.Forms.Label();
            lab_pt_current.Text = label_point_current.Text;
            lab_pt_current.Font = label_point_current.Font;
            lab_pt_current.Location = label_point_current.Location;
            lab_pt_current.BackColor = label_point_current.BackColor;
            lab_pt_current.ForeColor = label_point_current.ForeColor;
            lab_pt_current.AutoSize = true;



            ComboBox combo_pt = new ComboBox();
            combo_pt.Location = comboBox_current_pt.Location;
            combo_pt.Size = comboBox_current_pt.Size;
            combo_pt.BackColor = comboBox_current_pt.BackColor;
            combo_pt.ForeColor = comboBox_current_pt.ForeColor;
            combo_pt.FlatStyle = comboBox_current_pt.FlatStyle;
            combo_pt.Name = "combo_" + category1 + "_" + type1;
            combo_pt.SelectedIndexChanged += new EventHandler(comboBox_current_pt_SelectedIndexChanged);

            if (list_of_materials.Count > 0)
            {
                foreach (string mat1 in list_of_materials)
                {
                    combo_pt.Items.Add(mat1);
                }
                combo_pt.SelectedIndex = 0;
            }



            System.Windows.Forms.Button btn_pt_zoom = new System.Windows.Forms.Button();
            btn_pt_zoom.FlatStyle = button_pt_zoom.FlatStyle;
            btn_pt_zoom.FlatAppearance.BorderColor = button_pt_zoom.FlatAppearance.BorderColor;
            btn_pt_zoom.FlatAppearance.MouseDownBackColor = button_pt_zoom.FlatAppearance.MouseDownBackColor;
            btn_pt_zoom.FlatAppearance.MouseOverBackColor = button_pt_zoom.FlatAppearance.MouseOverBackColor;
            btn_pt_zoom.BackgroundImage = button_pt_zoom.BackgroundImage;
            btn_pt_zoom.BackgroundImageLayout = button_pt_zoom.BackgroundImageLayout;
            btn_pt_zoom.Bounds = button_pt_zoom.Bounds;
            btn_pt_zoom.ForeColor = button_pt_zoom.ForeColor;
            btn_pt_zoom.BackColor = button_pt_zoom.BackColor;
            btn_pt_zoom.Text = button_pt_zoom.Text;
            btn_pt_zoom.TextAlign = button_pt_zoom.TextAlign;
            btn_pt_zoom.Image = button_pt_zoom.Image;
            btn_pt_zoom.ImageAlign = button_pt_zoom.ImageAlign;
            btn_pt_zoom.Name = "btn_zoom_to_" + category1 + "_" + type1;
            btn_pt_zoom.Click += new EventHandler(button_pt_zoom_Click);


            System.Windows.Forms.Button btn_pt_select = new System.Windows.Forms.Button();
            btn_pt_select.FlatStyle = button_pt_select.FlatStyle;
            btn_pt_select.FlatAppearance.BorderColor = button_pt_select.FlatAppearance.BorderColor;
            btn_pt_select.FlatAppearance.MouseDownBackColor = button_pt_select.FlatAppearance.MouseDownBackColor;
            btn_pt_select.FlatAppearance.MouseOverBackColor = button_pt_select.FlatAppearance.MouseOverBackColor;
            btn_pt_select.BackgroundImage = button_pt_select.BackgroundImage;
            btn_pt_select.BackgroundImageLayout = button_pt_select.BackgroundImageLayout;
            btn_pt_select.Bounds = button_pt_select.Bounds;
            btn_pt_select.ForeColor = button_pt_select.ForeColor;
            btn_pt_select.BackColor = button_pt_select.BackColor;
            btn_pt_select.Text = button_pt_select.Text;
            btn_pt_select.TextAlign = button_pt_select.TextAlign;
            btn_pt_select.Image = button_pt_select.Image;
            btn_pt_select.ImageAlign = button_pt_select.ImageAlign;
            btn_pt_select.Name = "btn_select_" + category1 + "_" + type1;
            btn_pt_select.Click += new EventHandler(button_pt_select_Click);

            System.Windows.Forms.Button btn_pt_dwg = new System.Windows.Forms.Button();
            btn_pt_dwg.FlatStyle = button_pt_dwg.FlatStyle;
            btn_pt_dwg.FlatAppearance.BorderColor = button_pt_dwg.FlatAppearance.BorderColor;
            btn_pt_dwg.FlatAppearance.MouseDownBackColor = button_pt_dwg.FlatAppearance.MouseDownBackColor;
            btn_pt_dwg.FlatAppearance.MouseOverBackColor = button_pt_dwg.FlatAppearance.MouseOverBackColor;
            btn_pt_dwg.BackgroundImage = button_pt_dwg.BackgroundImage;
            btn_pt_dwg.BackgroundImageLayout = button_pt_dwg.BackgroundImageLayout;
            btn_pt_dwg.Bounds = button_pt_dwg.Bounds;
            btn_pt_dwg.ForeColor = button_pt_dwg.ForeColor;
            btn_pt_dwg.BackColor = button_pt_dwg.BackColor;
            btn_pt_dwg.Text = button_pt_dwg.Text;
            btn_pt_dwg.TextAlign = button_pt_dwg.TextAlign;
            btn_pt_dwg.Image = button_pt_dwg.Image;
            btn_pt_dwg.ImageAlign = button_pt_dwg.ImageAlign;
            btn_pt_dwg.Name = "button_draw_" + category1 + "_" + type1;
            btn_pt_dwg.Click += new EventHandler(button_place_point_Click);

            System.Windows.Forms.Button btn_pt_refresh = new System.Windows.Forms.Button();
            btn_pt_refresh.FlatStyle = button_pt_refresh.FlatStyle;
            btn_pt_refresh.FlatAppearance.BorderColor = button_pt_refresh.FlatAppearance.BorderColor;
            btn_pt_refresh.FlatAppearance.MouseDownBackColor = button_pt_refresh.FlatAppearance.MouseDownBackColor;
            btn_pt_refresh.FlatAppearance.MouseOverBackColor = button_pt_refresh.FlatAppearance.MouseOverBackColor;
            btn_pt_refresh.BackgroundImage = button_pt_refresh.BackgroundImage;
            btn_pt_refresh.BackgroundImageLayout = button_pt_refresh.BackgroundImageLayout;
            btn_pt_refresh.Bounds = button_pt_refresh.Bounds;
            btn_pt_refresh.ForeColor = button_pt_refresh.ForeColor;
            btn_pt_refresh.BackColor = button_pt_refresh.BackColor;
            btn_pt_refresh.Text = button_pt_refresh.Text;
            btn_pt_refresh.TextAlign = button_pt_refresh.TextAlign;
            btn_pt_refresh.Image = button_pt_refresh.Image;
            btn_pt_refresh.ImageAlign = button_pt_refresh.ImageAlign;
            btn_pt_refresh.Name = "button_refresh_" + category1 + "_" + type1;
            btn_pt_refresh.Click += new EventHandler(button_pt_refresh_Click);



            pan_left.Controls.AddRange(new Control[] { lab_pt_left, btn_pt_zoom, btn_pt_select, btn_pt_dwg, btn_pt_refresh, lab_pt_current, combo_pt });



            Panel pan_right = new Panel();
            pan_right.Location = new System.Drawing.Point(pan_left.Width, pan_top.Bottom);
            pan_right.Size = new Size(pan_top.Width - pan_left.Width, pan_bottom.Top - pan_top.Bottom);
            pan_right.ForeColor = panel_pt_left.ForeColor;
            pan_right.BackColor = panel_pt_left.BackColor;
            pan_right.BorderStyle = BorderStyle.FixedSingle;

            DataGridView dgv1 = new DataGridView();

            dgv1.BorderStyle = BorderStyle.FixedSingle;
            dgv1.Dock = DockStyle.Fill;
            dgv1.BackgroundColor = dataGridView_pt.BackgroundColor;
            dgv1.ForeColor = dataGridView_pt.ForeColor;
            dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dgv1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Padding newpadding = new Padding(4, 0, 0, 0);
            dgv1.ColumnHeadersDefaultCellStyle.Padding = newpadding;
            dgv1.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dgv1.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
            dgv1.DefaultCellStyle.ForeColor = Color.White;
            dgv1.EnableHeadersVisualStyles = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.RowHeadersVisible = false;
            dgv1.Name = "dgv_" + category1 + "_" + type1;
            dgv1.CellMouseClick += new DataGridViewCellMouseEventHandler(dataGridView_pt_CellMouseClick);

            DataGridViewTextBoxColumn dgv_item = Functions.datagrid_textbox_column(col_item_no);
            dgv_item.Name = col_item_no;
            DataGridViewTextBoxColumn dgv_2dsta = Functions.datagrid_textbox_column(col_2dsta);
            DataGridViewTextBoxColumn dgv_3dsta = Functions.datagrid_textbox_column(col_3dsta);
            DataGridViewTextBoxColumn dgv_eqsta = Functions.datagrid_textbox_column(col_eqsta);
            DataGridViewTextBoxColumn dgv_altdesc = Functions.datagrid_textbox_column(col_altdesc);
            // DataGridViewTextBoxColumn dgv_MSblock = Functions.datagrid_textbox_column(col_MSblock);
            //DataGridViewTextBoxColumn dgv_layer = Functions.datagrid_textbox_column(col_layer);
            //dgv_layer.Name = col_layer;
            //dgv_MSblock.Name = col_MSblock;

            dgv_altdesc.HeaderText = "Description";
            dgv1.Rows.Clear();

            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    if (dt1.Rows[0][col_2dsta] != DBNull.Value && dt1.Rows[0][col_3dsta] == DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_2dsta, dgv_altdesc);//, dgv_MSblock, dgv_layer);
                    }
                    if (dt1.Rows[0][col_3dsta] != DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_3dsta, dgv_altdesc);//, dgv_MSblock, dgv_layer);
                    }
                    if (dt1.Rows[0][col_eqsta] != DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_eqsta, dgv_altdesc);//, dgv_MSblock, dgv_layer);
                    }
                }
                else
                {
                    dgv1.Columns.Clear();
                    dgv1.Columns.AddRange(dgv_item, dgv_2dsta, dgv_altdesc);//, dgv_MSblock, dgv_layer);
                }
            }

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dt1;
            dgv1.CellClick += new DataGridViewCellEventHandler(dataGridView_pt_CellClick);

            pan_right.Controls.AddRange(new Control[] { dgv1 });

            tp.Controls.AddRange(new Control[] { pan_top, pan_bottom, pan_left, pan_right });


            return tp;
        }

        public TabPage create_linear_tab(TabPage tab_lin, List<string> list_of_materials, System.Data.DataTable dt1 = null)
        {
            string category1 = "**";
            string type1 = "**";
            string descr = "**";

            for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
            {
                if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                {
                    string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper().Replace(" ", "");
                    if (list_of_materials[0] == mat2)
                    {
                        category1 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                        type1 = Convert.ToString(dt_mat_library.Rows[j][col_type]);
                        descr = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                    }
                }
            }

            TabPage tp = new TabPage();
            tp.BackColor = tab_lin.BackColor;
            tp.ForeColor = tab_lin.ForeColor;
            tp.Font = tab_lin.Font;
            tp.Text = category1;

            Panel pan_top = new Panel();
            pan_top.Bounds = panel_lin_top.Bounds;
            pan_top.ForeColor = panel_lin_top.ForeColor;
            pan_top.BackColor = panel_lin_top.BackColor;
            pan_top.BorderStyle = BorderStyle.FixedSingle;
            pan_top.Dock = DockStyle.Top;


            System.Windows.Forms.Label lab_linear = new System.Windows.Forms.Label();
            lab_linear.Text = category1 + " Material Data";
            lab_linear.Font = label_linear.Font;
            lab_linear.Location = label_linear.Location;
            lab_linear.BackColor = label_linear.BackColor;
            lab_linear.ForeColor = label_linear.ForeColor;
            lab_linear.Dock = label_linear.Dock;
            lab_linear.AutoSize = true;
            pan_top.Controls.AddRange(new Control[] { lab_linear });

            Panel pan_bottom = new Panel();
            pan_bottom.Bounds = panel_lin_bottom.Bounds;
            pan_bottom.ForeColor = panel_lin_bottom.ForeColor;
            pan_bottom.BackColor = panel_lin_bottom.BackColor;
            pan_bottom.BorderStyle = BorderStyle.FixedSingle;
            pan_bottom.Dock = DockStyle.Bottom;

            System.Windows.Forms.Button btn_save_lin = new System.Windows.Forms.Button();
            btn_save_lin.FlatStyle = button_lin_save.FlatStyle;
            btn_save_lin.FlatAppearance.BorderColor = button_lin_save.FlatAppearance.BorderColor;
            btn_save_lin.FlatAppearance.MouseDownBackColor = button_lin_save.FlatAppearance.MouseDownBackColor;
            btn_save_lin.FlatAppearance.MouseOverBackColor = button_lin_save.FlatAppearance.MouseOverBackColor;
            btn_save_lin.BackgroundImage = button_lin_save.BackgroundImage;
            btn_save_lin.BackgroundImageLayout = button_lin_save.BackgroundImageLayout;
            btn_save_lin.Bounds = button_lin_save.Bounds;
            btn_save_lin.Dock = button_lin_save.Dock;
            btn_save_lin.Click += new EventHandler(button_lin_save_Click);
            pan_bottom.Controls.AddRange(new Control[] { btn_save_lin });



            Panel pan_left = new Panel();
            pan_left.Location = new System.Drawing.Point(0, pan_top.Bottom);
            pan_left.Size = new Size(panel_lin_left.Width, pan_bottom.Top - pan_top.Bottom);
            pan_left.ForeColor = panel_lin_left.ForeColor;
            pan_left.BackColor = panel_lin_left.BackColor;
            pan_left.BorderStyle = BorderStyle.FixedSingle;

            System.Windows.Forms.Label lab_current = new System.Windows.Forms.Label();
            lab_current.Text = label_linear_current.Text;
            lab_current.Font = label_linear_current.Font;
            lab_current.Location = label_linear_current.Location;
            lab_current.BackColor = label_linear_current.BackColor;
            lab_current.ForeColor = label_linear_current.ForeColor;
            lab_current.AutoSize = true;

            System.Windows.Forms.Label lab_lin_left = new System.Windows.Forms.Label();
            lab_lin_left.Text = descr;
            lab_lin_left.Font = label_lin_left.Font;
            lab_lin_left.Location = label_lin_left.Location;
            lab_lin_left.BackColor = label_lin_left.BackColor;
            lab_lin_left.ForeColor = label_lin_left.ForeColor;
            lab_lin_left.AutoSize = true;
            lab_lin_left.Name = "lab_lin_left_" + category1 + "_" + type1;

            ComboBox combo_linear = new ComboBox();
            combo_linear.Location = comboBox_linear_current.Location;
            combo_linear.Size = comboBox_linear_current.Size;
            combo_linear.BackColor = comboBox_linear_current.BackColor;
            combo_linear.ForeColor = comboBox_linear_current.ForeColor;
            combo_linear.FlatStyle = comboBox_linear_current.FlatStyle;
            combo_linear.Name = "combo_lin_" + category1 + "_" + type1;
            combo_linear.SelectedIndexChanged += new EventHandler(comboBox_linear_current_SelectedIndexChanged);

            if (list_of_materials.Count > 0)
            {
                foreach (string mat1 in list_of_materials)
                {
                    combo_linear.Items.Add(mat1);
                }
                combo_linear.SelectedIndex = 0;
            }




            System.Windows.Forms.Button btn_lin_dwg = new System.Windows.Forms.Button();
            btn_lin_dwg.FlatStyle = button_lin_dwg.FlatStyle;
            btn_lin_dwg.FlatAppearance.BorderColor = button_lin_dwg.FlatAppearance.BorderColor;
            btn_lin_dwg.FlatAppearance.MouseDownBackColor = button_lin_dwg.FlatAppearance.MouseDownBackColor;
            btn_lin_dwg.FlatAppearance.MouseOverBackColor = button_lin_dwg.FlatAppearance.MouseOverBackColor;
            btn_lin_dwg.BackgroundImage = button_lin_dwg.BackgroundImage;
            btn_lin_dwg.BackgroundImageLayout = button_lin_dwg.BackgroundImageLayout;
            btn_lin_dwg.Bounds = button_lin_dwg.Bounds;
            btn_lin_dwg.ForeColor = button_lin_dwg.ForeColor;
            btn_lin_dwg.BackColor = button_lin_dwg.BackColor;
            btn_lin_dwg.Text = "Draw " + descr;
            btn_lin_dwg.TextAlign = button_lin_dwg.TextAlign;
            btn_lin_dwg.Image = button_lin_dwg.Image;
            btn_lin_dwg.ImageAlign = button_lin_dwg.ImageAlign;
            btn_lin_dwg.Name = "button_lin_draw_" + category1 + "_" + type1;
            btn_lin_dwg.Click += new EventHandler(button_lin_dwg_Click);


            System.Windows.Forms.Button btn_lin_zoom = new System.Windows.Forms.Button();
            btn_lin_zoom.FlatStyle = button_lin_zoom.FlatStyle;
            btn_lin_zoom.FlatAppearance.BorderColor = button_lin_zoom.FlatAppearance.BorderColor;
            btn_lin_zoom.FlatAppearance.MouseDownBackColor = button_lin_zoom.FlatAppearance.MouseDownBackColor;
            btn_lin_zoom.FlatAppearance.MouseOverBackColor = button_lin_zoom.FlatAppearance.MouseOverBackColor;
            btn_lin_zoom.BackgroundImage = button_lin_zoom.BackgroundImage;
            btn_lin_zoom.BackgroundImageLayout = button_lin_zoom.BackgroundImageLayout;
            btn_lin_zoom.Bounds = button_lin_zoom.Bounds;
            btn_lin_zoom.ForeColor = button_lin_zoom.ForeColor;
            btn_lin_zoom.BackColor = button_lin_zoom.BackColor;
            btn_lin_zoom.Text = button_lin_zoom.Text;
            btn_lin_zoom.TextAlign = button_lin_zoom.TextAlign;
            btn_lin_zoom.Image = button_lin_zoom.Image;
            btn_lin_zoom.ImageAlign = button_lin_zoom.ImageAlign;
            btn_lin_zoom.Name = "btn_zoom_to_" + category1 + "_" + type1;
            btn_lin_zoom.Click += new EventHandler(button_lin_zoom_Click);


            System.Windows.Forms.Button btn_lin_select = new System.Windows.Forms.Button();
            btn_lin_select.FlatStyle = button_lin_select.FlatStyle;
            btn_lin_select.FlatAppearance.BorderColor = button_lin_select.FlatAppearance.BorderColor;
            btn_lin_select.FlatAppearance.MouseDownBackColor = button_lin_select.FlatAppearance.MouseDownBackColor;
            btn_lin_select.FlatAppearance.MouseOverBackColor = button_lin_select.FlatAppearance.MouseOverBackColor;
            btn_lin_select.BackgroundImage = button_lin_select.BackgroundImage;
            btn_lin_select.BackgroundImageLayout = button_lin_select.BackgroundImageLayout;
            btn_lin_select.Bounds = button_lin_select.Bounds;
            btn_lin_select.ForeColor = button_lin_select.ForeColor;
            btn_lin_select.BackColor = button_lin_select.BackColor;
            btn_lin_select.Text = button_lin_select.Text;
            btn_lin_select.TextAlign = button_lin_select.TextAlign;
            btn_lin_select.Image = button_lin_select.Image;
            btn_lin_select.ImageAlign = button_lin_select.ImageAlign;
            btn_lin_select.Name = "btn_select_" + category1 + "_" + type1;
            btn_lin_select.Click += new EventHandler(button_lin_select_Click);

            System.Windows.Forms.Button btn_line_refresh = new System.Windows.Forms.Button();
            btn_line_refresh.FlatStyle = button_line_refresh.FlatStyle;
            btn_line_refresh.FlatAppearance.BorderColor = button_line_refresh.FlatAppearance.BorderColor;
            btn_line_refresh.FlatAppearance.MouseDownBackColor = button_line_refresh.FlatAppearance.MouseDownBackColor;
            btn_line_refresh.FlatAppearance.MouseOverBackColor = button_line_refresh.FlatAppearance.MouseOverBackColor;
            btn_line_refresh.BackgroundImage = button_line_refresh.BackgroundImage;
            btn_line_refresh.BackgroundImageLayout = button_line_refresh.BackgroundImageLayout;
            btn_line_refresh.Bounds = button_line_refresh.Bounds;
            btn_line_refresh.ForeColor = button_line_refresh.ForeColor;
            btn_line_refresh.BackColor = button_line_refresh.BackColor;
            btn_line_refresh.Text = button_line_refresh.Text;
            btn_line_refresh.TextAlign = button_line_refresh.TextAlign;
            btn_line_refresh.Image = button_line_refresh.Image;
            btn_line_refresh.ImageAlign = button_line_refresh.ImageAlign;
            btn_line_refresh.Name = "button_refresh_" + category1 + "_" + type1;
            btn_line_refresh.Click += new EventHandler(button_line_refresh_Click);

            pan_left.Controls.AddRange(new Control[] { lab_current, btn_lin_dwg, btn_lin_zoom, btn_lin_select, lab_current, combo_linear, lab_lin_left, btn_line_refresh });


            Panel pan_right = new Panel();
            pan_right.Location = new System.Drawing.Point(pan_left.Width, pan_top.Bottom);
            pan_right.Size = new Size(pan_top.Width - pan_left.Width, pan_bottom.Top - pan_top.Bottom);
            pan_right.ForeColor = panel_lin_left.ForeColor;
            pan_right.BackColor = panel_lin_left.BackColor;
            pan_right.BorderStyle = BorderStyle.FixedSingle;

            DataGridView dgv1 = new DataGridView();

            dgv1.BorderStyle = BorderStyle.FixedSingle;
            dgv1.Dock = DockStyle.Fill;
            dgv1.BackgroundColor = dataGridView_pt.BackgroundColor;
            dgv1.ForeColor = dataGridView_pt.ForeColor;
            dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dgv1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Padding newpadding = new Padding(4, 0, 0, 0);
            dgv1.ColumnHeadersDefaultCellStyle.Padding = newpadding;
            dgv1.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dgv1.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
            dgv1.DefaultCellStyle.ForeColor = Color.White;
            dgv1.EnableHeadersVisualStyles = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.RowHeadersVisible = false;
            dgv1.Name = "dgv_" + category1 + "_" + type1;
            dgv1.CellMouseClick += new DataGridViewCellMouseEventHandler(dataGridView_extra_CellMouseClick);

            DataGridViewTextBoxColumn dgv_item = Functions.datagrid_textbox_column(col_item_no);
            dgv_item.Name = col_item_no;
            DataGridViewTextBoxColumn dgv_2dsta1 = Functions.datagrid_textbox_column(pipe_col_2d1);
            DataGridViewTextBoxColumn dgv_2dsta2 = Functions.datagrid_textbox_column(pipe_col_2d2);
            DataGridViewTextBoxColumn dgv_3dsta1 = Functions.datagrid_textbox_column(pipe_col_3d1);
            DataGridViewTextBoxColumn dgv_3dsta2 = Functions.datagrid_textbox_column(pipe_col_3d2);
            DataGridViewTextBoxColumn dgv_eqsta1 = Functions.datagrid_textbox_column(pipe_col_eq1);
            DataGridViewTextBoxColumn dgv_eqsta2 = Functions.datagrid_textbox_column(pipe_col_eq2);
            DataGridViewTextBoxColumn dgv_len2d = Functions.datagrid_textbox_column(pipe_col_len2d);
            DataGridViewTextBoxColumn dgv_len3d = Functions.datagrid_textbox_column(pipe_col_len3d);

            DataGridViewTextBoxColumn dgv_altdesc = Functions.datagrid_textbox_column(col_altdesc);
            dgv_altdesc.HeaderText = "Description";


            dgv1.Rows.Clear();

            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    if (dt1.Rows[0][pipe_col_2d1] != DBNull.Value && dt1.Rows[0][pipe_col_2d2] != DBNull.Value && dt1.Rows[0][pipe_col_3d1] == DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_2dsta1, dgv_2dsta2, dgv_len2d, dgv_altdesc);
                    }
                    if (dt1.Rows[0][pipe_col_3d1] != DBNull.Value && dt1.Rows[0][pipe_col_3d2] != DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_3dsta1, dgv_3dsta2, dgv_len3d, dgv_altdesc);
                    }
                    if (dt1.Rows[0][pipe_col_eq1] != DBNull.Value && dt1.Rows[0][pipe_col_eq2] != DBNull.Value && dt1.Rows[0][pipe_col_len2d] != DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_eqsta1, dgv_eqsta2, dgv_len2d, dgv_altdesc);
                    }
                    if (dt1.Rows[0][pipe_col_eq1] != DBNull.Value && dt1.Rows[0][pipe_col_eq2] != DBNull.Value && dt1.Rows[0][pipe_col_len3d] != DBNull.Value)
                    {
                        dgv1.Columns.Clear();
                        dgv1.Columns.AddRange(dgv_item, dgv_eqsta1, dgv_eqsta2, dgv_len3d, dgv_altdesc);
                    }
                }
                else
                {
                    dgv1.Columns.Clear();
                    dgv1.Columns.AddRange(dgv_item, dgv_2dsta1, dgv_2dsta2, dgv_len2d, dgv_altdesc);
                }
            }

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dt1;
            dgv1.CellClick += new DataGridViewCellEventHandler(dataGridView_pt_CellClick);


            pan_right.Controls.AddRange(new Control[] { dgv1 });

            tp.Controls.AddRange(new Control[] { pan_top, pan_bottom, pan_left, pan_right });


            return tp;
        }



        public void Add_FlatTab(string TabName)
        {
            System.Drawing.Font Header_Font = new System.Drawing.Font("Arial Black", 9.75f, FontStyle.Bold);
            System.Drawing.Font Label_Font = new System.Drawing.Font("Arial", 8.25f, FontStyle.Bold);
            Size Combobox_Size = new Size(207, 22);
            Size Button_Size = new Size(207, 28);

            flatTabControl1.TabPages.Add(TabName, TabName);
            TabPage tbpg_new = flatTabControl1.TabPages[flatTabControl1.TabPages.Count - 1];
            tbpg_new.BackColor = Color.FromArgb(37, 37, 38);
            tbpg_new.ForeColor = Color.FromArgb(0, 122, 204);
            tbpg_new.Padding = new Padding(0);
            tbpg_new.Margin = new Padding(0);

            Panel pnl_header = new Panel();
            pnl_header.Height = 29;
            pnl_header.BorderStyle = BorderStyle.FixedSingle;
            pnl_header.BackColor = Color.FromArgb(28, 28, 28);
            pnl_header.Margin = new Padding(3);
            pnl_header.Dock = DockStyle.Top;

            System.Windows.Forms.Label lbl_header = new System.Windows.Forms.Label();
            lbl_header.Text = TabName + " Material Data";
            lbl_header.Font = Header_Font;
            lbl_header.ForeColor = Color.FromArgb(0, 122, 204);
            lbl_header.Location = new System.Drawing.Point(3, 4);
            lbl_header.AutoSize = true;

            Panel pnl_btn_1 = new Panel();
            pnl_btn_1.Size = new Size(27, 27);
            pnl_btn_1.BorderStyle = BorderStyle.None;
            pnl_btn_1.BackColor = Color.FromArgb(28, 28, 28);
            pnl_btn_1.Margin = new Padding(3);
            pnl_btn_1.Padding = new Padding(2);
            pnl_btn_1.Dock = DockStyle.Right;

            System.Windows.Forms.Button btn_export = new System.Windows.Forms.Button();
            btn_export.FlatStyle = FlatStyle.Flat;
            btn_export.FlatAppearance.BorderColor = Color.DimGray;
            btn_export.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btn_export.FlatAppearance.MouseOverBackColor = Color.DarkOrange;
            btn_export.BackgroundImage = (Alignment_mdi.Properties.Resources.excel_icon_OUT);
            btn_export.BackgroundImageLayout = ImageLayout.Stretch;
            btn_export.Dock = DockStyle.Fill;

            pnl_btn_1.Controls.AddRange(new Control[] { btn_export });

            Panel pnl_btn_2 = new Panel();
            pnl_btn_2.Size = new Size(27, 27);
            pnl_btn_2.BorderStyle = BorderStyle.None;
            pnl_btn_2.BackColor = Color.FromArgb(28, 28, 28);
            pnl_btn_2.Margin = new Padding(3);
            pnl_btn_2.Padding = new Padding(2);
            pnl_btn_2.Dock = DockStyle.Right;

            System.Windows.Forms.Button btn_import = new System.Windows.Forms.Button();
            btn_import.FlatStyle = FlatStyle.Flat;
            btn_import.FlatAppearance.BorderColor = Color.DimGray;
            btn_import.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btn_import.FlatAppearance.MouseOverBackColor = Color.DarkOrange;
            btn_import.BackgroundImage = (Alignment_mdi.Properties.Resources.excel_icon_IN);
            btn_import.BackgroundImageLayout = ImageLayout.Stretch;
            btn_import.Dock = DockStyle.Fill;

            pnl_btn_2.Controls.AddRange(new Control[] { btn_import });

            pnl_header.Controls.AddRange(new Control[] { lbl_header, pnl_btn_1, pnl_btn_2 });

            Panel pnl1 = new Panel();
            pnl1.Width = 217;
            pnl1.BorderStyle = BorderStyle.FixedSingle;
            pnl1.Margin = new Padding(3);
            pnl1.Padding = new Padding(0);
            pnl1.Dock = DockStyle.Left;

            System.Windows.Forms.Label lbl_material = new System.Windows.Forms.Label();
            lbl_material.Text = "Current " + TabName + " Material";
            lbl_material.Font = Label_Font;
            lbl_material.ForeColor = Color.White;
            lbl_material.Location = new System.Drawing.Point(0, 0);
            lbl_material.AutoSize = false;
            lbl_material.Width = 217;
            lbl_material.TextAlign = ContentAlignment.MiddleCenter;

            ComboBox cmbx_material = new ComboBox();
            cmbx_material.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbx_material.Size = Combobox_Size;
            cmbx_material.Location = new System.Drawing.Point(3, 24);
            cmbx_material.BackColor = Color.FromArgb(51, 51, 55);
            cmbx_material.ForeColor = Color.White;
            cmbx_material.FlatStyle = FlatStyle.Flat;





            System.Windows.Forms.Button btn_2pts = new System.Windows.Forms.Button();
            btn_2pts.Text = "Select Location on Pipe";
            btn_2pts.Size = Button_Size;
            btn_2pts.Location = new System.Drawing.Point(3, 52);
            btn_2pts.FlatStyle = FlatStyle.Flat;
            btn_2pts.FlatAppearance.BorderColor = Color.DimGray;
            btn_2pts.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btn_2pts.FlatAppearance.MouseOverBackColor = Color.DarkOrange;
            btn_2pts.BackgroundImage = (Alignment_mdi.Properties.Resources.selectbluexs);
            btn_2pts.BackgroundImageLayout = ImageLayout.None;
            btn_2pts.BackColor = Color.FromArgb(51, 51, 55);
            btn_2pts.ForeColor = Color.White;
            //btn_2pts.Click += new EventHandler(button_pipe_pick_pts_Click);

            System.Windows.Forms.Button btn_Zoom_To = new System.Windows.Forms.Button();
            btn_Zoom_To.Text = "Zoom To";
            btn_Zoom_To.Size = Button_Size;
            btn_Zoom_To.Location = new System.Drawing.Point(3, 86);
            btn_Zoom_To.FlatStyle = FlatStyle.Flat;
            btn_Zoom_To.FlatAppearance.BorderColor = Color.DimGray;
            btn_Zoom_To.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btn_Zoom_To.FlatAppearance.MouseOverBackColor = Color.DarkOrange;
            btn_Zoom_To.BackgroundImage = (Alignment_mdi.Properties.Resources.Target);
            btn_Zoom_To.BackgroundImageLayout = ImageLayout.None;
            btn_Zoom_To.BackColor = Color.FromArgb(51, 51, 55);
            btn_Zoom_To.ForeColor = Color.White;

            System.Windows.Forms.Button btn_Mat_Select = new System.Windows.Forms.Button();
            btn_Mat_Select.Text = "Select Material";
            btn_Mat_Select.Size = Button_Size;
            btn_Mat_Select.Location = new System.Drawing.Point(3, 120);
            btn_Mat_Select.FlatStyle = FlatStyle.Flat;
            btn_Mat_Select.FlatAppearance.BorderColor = Color.DimGray;
            btn_Mat_Select.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btn_Mat_Select.FlatAppearance.MouseOverBackColor = Color.DarkOrange;
            //btn_Mat_Select.BackgroundImage = (Matl_Design.Properties.Resources.Target);
            //btn_Mat_Select.BackgroundImageLayout = ImageLayout.None;
            btn_Mat_Select.BackColor = Color.FromArgb(51, 51, 55);
            btn_Mat_Select.ForeColor = Color.White;

            pnl1.Controls.AddRange(new Control[] { lbl_material, cmbx_material, btn_2pts, btn_Zoom_To, btn_Mat_Select });

            DataGridView dgv1 = new DataGridView();
            dgv1.BorderStyle = BorderStyle.Fixed3D;
            dgv1.Dock = DockStyle.Fill;
            dgv1.BackgroundColor = Color.FromArgb(37, 37, 38);

            tbpg_new.Controls.AddRange(new Control[] { pnl1, dgv1, pnl_header });

        }

        private void clear_all_data_mat_lib_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("are sure you want to clear all the data?/r/nThis will clear also the other tables!", "material design tool", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
            {
                return;
            }

            try
            {
                dt_mat_library = null;
                ds_main.dt_pipe = null;

                dataGridView_mat_library.DataSource = dt_mat_library;
                dataGridView_pipe.DataSource = ds_main.dt_pipe;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void dataGridView_mat_lib_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                ContextMenuStrip_dt_mat_lib.Show(Cursor.Position);
                ContextMenuStrip_dt_mat_lib.Visible = true;
            }
            else
            {
                ContextMenuStrip_dt_mat_lib.Visible = false;
            }
        }

        private void delete_current_row_dt_pipe_Click(object sender, EventArgs e)
        {



            if (comboBox_default_mat.Text == "")
            {
                MessageBox.Show("No default material selected/r/nOperation aborted");
                return;
            }

            string mat0 = comboBox_default_mat.Text;

            try
            {

                if (dataGridView_pipe.RowCount > 0)
                {
                    int index_grid_pipe = dataGridView_pipe.CurrentCell.RowIndex;
                    if (index_grid_pipe == -1)
                    {
                        return;
                    }


                    bool is3D = ds_main.is3D;
                    dataGridView_pipe.Rows[index_grid_pipe].Cells[col_item_no].Value = mat0;

                    string prev_mat = Convert.ToString(ds_main.dt_pipe.Rows[ds_main.dt_pipe.Rows.Count - 1][col_item_no]);
                    for (int i = ds_main.dt_pipe.Rows.Count - 2; i >= 0; --i)
                    {
                        string mat3 = Convert.ToString(ds_main.dt_pipe.Rows[i][col_item_no]);

                        if (mat3 == prev_mat)
                        {
                            if (is3D == true)
                            {
                                ds_main.dt_pipe.Rows[i + 1][pipe_col_3d1] = ds_main.dt_pipe.Rows[i][pipe_col_3d1];
                            }
                            else
                            {
                                ds_main.dt_pipe.Rows[i + 1][pipe_col_2d1] = ds_main.dt_pipe.Rows[i][pipe_col_2d1];
                            }
                            ds_main.dt_pipe.Rows[i].Delete();

                        }
                        prev_mat = mat3;
                    }
                    draw_pipes();


                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void dataGridView_pipe_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                ContextMenuStrip_dt_pipe.Show(Cursor.Position);
                ContextMenuStrip_dt_pipe.Visible = true;
            }
            else
            {
                ContextMenuStrip_dt_pipe.Visible = false;
            }
        }

        private void dataGridView_pt_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            current_ct = "";
            DataGridView dgv1 = sender as DataGridView;
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                ContextMenuStrip_dt_pt.Show(Cursor.Position);
                ContextMenuStrip_dt_pt.Visible = true;

                current_ct = dgv1.Name.Replace("dgv_", "").ToUpper();
            }
            else
            {
                ContextMenuStrip_dt_pt.Visible = false;
            }
        }

        private void dataGridView_extra_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            current_ct = "";
            DataGridView dgv1 = sender as DataGridView;
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                ContextMenuStrip_dt_extra.Show(Cursor.Position);
                ContextMenuStrip_dt_extra.Visible = true;

                current_ct = dgv1.Name.Replace("dgv_", "").ToUpper();
            }
            else
            {
                ContextMenuStrip_dt_extra.Visible = false;
            }
        }

        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_matl_library);

            lista_butoane.Add(comboBox_default_mat);
            lista_butoane.Add(button_refresh_usa);


            lista_butoane.Add(button_add_row_to_library);
            lista_butoane.Add(button_save_mat_library);
            lista_butoane.Add(button_pipe_pick_pts);
            lista_butoane.Add(button_zoom_to_pipe);
            lista_butoane.Add(button_select_material_pipe);
            lista_butoane.Add(button_save_pipe);
            lista_butoane.Add(button_save_all);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_matl_library);

            lista_butoane.Add(comboBox_default_mat);
            lista_butoane.Add(button_refresh_usa);

            lista_butoane.Add(button_add_row_to_library);
            lista_butoane.Add(button_save_mat_library);

            lista_butoane.Add(button_pipe_pick_pts);
            lista_butoane.Add(button_zoom_to_pipe);
            lista_butoane.Add(button_select_material_pipe);
            lista_butoane.Add(button_save_pipe);
            lista_butoane.Add(button_save_all);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion

        #region hector work

        public void get_client_project_segment_pipe_diam()
        {
            ds_main.client1 = ds_main.tpage_main.get_textBox_client_name();
            ds_main.diam1 = ds_main.tpage_main.get_textBox_pipe_diam();
            ds_main.segment1 = ds_main.tpage_main.get_textBox_segment();
            ds_main.project1 = ds_main.tpage_main.get_textBox_project();
        }





        private void button_save_mat_library_Click(object sender, EventArgs e)
        {
            set_enable_false();
            get_client_project_segment_pipe_diam();
            dataGridView_mat_library.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            Transfer_datatable_to_file1(dt_mat_library);
            textBox_library.Text = ds_main.config_xls;
            textBox_library.ForeColor = Color.LightGreen;
            textBox_library.Font = font8;
            textBox_library.TextAlign = HorizontalAlignment.Right;
            set_enable_true();
        }



        public void Transfer_datatable_to_file1(System.Data.DataTable dt1)
        {

            sync_dt_mat_with_filter();

            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook_cfg = null;
            Microsoft.Office.Interop.Excel.Worksheet W_mat_lib = null;
            Microsoft.Office.Interop.Excel.Worksheet W_cfg = null;



            bool is_opened = false;
            bool save_and_close = false;

            try
            {

                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            if (Workbook2.FullName == ds_main.config_xls)
                            {
                                Workbook_cfg = Workbook2;
                                is_opened = true;
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                                {
                                    if (Wx.Name == "MDConfig")
                                    {
                                        W_cfg = Wx;
                                    }
                                    if (Wx.Name == "MatDesc")
                                    {
                                        W_mat_lib = Wx;
                                    }
                                }

                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    if (is_opened == false)
                    {
                        Workbook_cfg = Excel1.Workbooks.Open(ds_main.config_xls);
                        save_and_close = true;
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                        {
                            if (Wx.Name == "MDConfig")
                            {
                                W_cfg = Wx;
                            }
                            if (Wx.Name == "MatDesc")
                            {
                                W_mat_lib = Wx;
                            }
                        }
                    }
                }


                if (Workbook_cfg != null)
                {


                    if (W_cfg == null)
                    {
                        W_cfg = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W_cfg.Name = "MDConfig";
                    }

                    if (W_mat_lib == null)
                    {
                        W_mat_lib = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W_mat_lib.Name = "MatDesc";
                    }







                    if (dt1 != null && W_mat_lib != null)
                    {
                        if (dt1.Rows.Count > 0)
                        {
                            Create_header_material_library(W_mat_lib, ds_main.client1, ds_main.project1, ds_main.segment1, dt_mat_library);

                            W_mat_lib.Cells.NumberFormat = "General";
                            int maxRows = dt1.Rows.Count;
                            int maxCols = dt1.Columns.Count;
                            W_mat_lib.Range["A14:G1000"].ClearContents();
                            W_mat_lib.Range["A14:G1000"].ClearFormats();

                            Microsoft.Office.Interop.Excel.Range range1 = W_mat_lib.Range["A14:G" + (14 + maxRows - 1).ToString()];
                            object[,] values1 = new object[maxRows, maxCols];

                            for (int i = 0; i < maxRows; ++i)
                            {
                                for (int j = 0; j < maxCols; ++j)
                                {
                                    if (dt1.Rows[i][j] != DBNull.Value && j > 0)// i did not want to save mmid value
                                    {
                                        values1[i, j] = dt1.Rows[i][j];
                                    }
                                }
                            }
                            range1.Value2 = values1;

                        }

                        if (W_cfg != null)
                        {
                            W_cfg.Range["B6"].Value2 = ds_main.config_xls;
                        }


                        if (is_opened == true)
                        {
                            Workbook_cfg.Save();
                        }
                        if (save_and_close == true)
                        {
                            Workbook_cfg.Save();
                            Workbook_cfg.Close();
                        }
                    }


                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W_mat_lib != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_mat_lib);
                if (W_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cfg);
                if (Workbook_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_cfg);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }

        public static void Create_header_material_library(Microsoft.Office.Interop.Excel.Worksheet W1, string Client, string Project, string Segment, System.Data.DataTable dt_lib)
        {



            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B11"];
            object[,] valuesH = new object[11, 2];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";

            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at: " + DateTime.Now.TimeOfDay;
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;
            valuesH[6, 0] = "JOINT LENGTH";
            valuesH[7, 0] = "Engineering is responsible for the content and QAQC of this table.";
            valuesH[8, 0] = "This Table shows Material that has a Begin and End Station";
            valuesH[9, 0] = "Do not add any columns to this table, also do not add any rows above row 13";
            valuesH[10, 0] = "This Table is to be populated by Engineering and Programming as indicated below:";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:B7"];

            Functions.Color_border_range_inside(range1, 46);

            range1 = W1.Range["A8:G8"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A9:G9"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A10:G10"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 3); //red

            range1 = W1.Range["A11:G11"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 43); //green

            range1 = W1.Range["A12:G12"];
            object[,] values12 = new object[1, 7];
            values12[0, 0] = "N/A";
            values12[0, 1] = "ENG";
            values12[0, 2] = "ENG";
            values12[0, 3] = "ENG";
            values12[0, 4] = "ENG";
            values12[0, 5] = "ENG";
            values12[0, 6] = "USER";


            range1.Value2 = values12;
            Functions.Color_border_range_inside(range1, 43); //green

            range1 = W1.Range["C1:G7"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "Material Description";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Functions.Color_border_range_outside(range1, 0);


            W1.Range["A7"].Font.Bold = true;

            range1 = W1.Range["A13:G13"];
            object[,] values_dt = new object[1, dt_lib.Columns.Count];
            if (dt_lib != null && dt_lib.Columns.Count > 0)
            {
                for (int i = 0; i < dt_lib.Columns.Count; ++i)
                {
                    values_dt[0, i] = dt_lib.Columns[i].ColumnName;
                }
                range1.Value2 = values_dt;
                Functions.Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }


            W1.Range["A:B"].ColumnWidth = 13;
            W1.Range["C:C"].ColumnWidth = 60;
            W1.Range["D:E"].ColumnWidth = 13;
            W1.Range["F:F"].ColumnWidth = 25;
            W1.Range["G:G"].ColumnWidth = 13;


        }


        public System.Data.DataTable Creaza_mat_library_structure()
        {

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(col_mmid, typeof(string));
            dt1.Columns.Add(col_item_no, typeof(string));
            dt1.Columns.Add(col_descr, typeof(string));
            dt1.Columns.Add(col_category, typeof(string));
            dt1.Columns.Add(col_type, typeof(string));
            dt1.Columns.Add(col_layer, typeof(string));
            dt1.Columns.Add(col_MSblock, typeof(string));
            return dt1;
        }
        private void Button_load_matl_library_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(ds_main.config_xls) == false)
            {
                MessageBox.Show("No configuration file found", "MDtool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            set_enable_false();
            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W2 = null;
            Microsoft.Office.Interop.Excel.Worksheet W3 = null;
            Microsoft.Office.Interop.Excel.Worksheet W4 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            bool close_file = false;


            textBox_library.Text = "No library loaded";
            textBox_library.ForeColor = Color.Red;
            textBox_library.Font = font10;
            textBox_library.TextAlign = HorizontalAlignment.Left;



            using (System.Windows.Forms.OpenFileDialog fbd = new System.Windows.Forms.OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    load_W1_W2_W3_W4_for_mat_library(fbd.FileName, ref Excel1, ref Workbook1, ref W1, ref W2, ref W3, ref W4);
                    close_file = true;
                }
                else
                {
                    set_enable_true();

                    return;
                }
            }







            dt_mat_library = Creaza_mat_library_structure();

            ds_main.dt_pipe = null;
            ds_main.dt_points = null;
            ds_main.dt_extra = null;

            if (ds_main.tpage_centerline.dt_cl_display != null && ds_main.tpage_centerline.dt_cl_display.Rows.Count > 0)
            {
                for (int i = 0; i < ds_main.tpage_centerline.dt_cl_display.Rows.Count; ++i)
                {
                    ds_main.tpage_centerline.dt_cl_display.Rows[i][col_elbow] = DBNull.Value;
                    ds_main.tpage_centerline.dt_cl_display.Rows[i][col_mat_elbow] = DBNull.Value;
                }
            }

            ds_main.tpage_centerline.dt_cl_display_filtered = ds_main.tpage_centerline.dt_cl_display;

            load_bom(ds_main.config_xls, W1, W2, W3, W4);



            if (dt_mat_library == null || dt_mat_library.Rows.Count == 0)
            {
                textBox_library.Text = "No library loaded";
                textBox_library.ForeColor = Color.Red;
                textBox_library.Font = font10;
                textBox_library.TextAlign = HorizontalAlignment.Left;
                set_enable_true();
                return;
            }

            List<string> lista1 = Functions.get_blocks_from_current_drawing();

            for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
            {
                if (dt_mat_library.Rows[i][col_MSblock] != DBNull.Value)
                {
                    string bn = Convert.ToString(dt_mat_library.Rows[i][col_MSblock]);

                    if (lista1.Contains(bn) == false)
                    {
                        MessageBox.Show("the block " + bn + " not present in current drawing", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }


            try
            {
                ct_list = build_category_and_type_list_and_dt_ct();

                add_tab_pages();

                populate_datagridview_pipe();
                ds_main.tpage_centerline.populate_datagridview_cl();
                ds_main.tpage_centerline.add_elbows_mat_to_combobox(dt_mat_library);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                textBox_library.Text = "No library loaded";
                textBox_library.ForeColor = Color.Red;
                textBox_library.Font = font10;
                textBox_library.TextAlign = HorizontalAlignment.Left;

            }
            finally
            {

                if (close_file == true)
                {

                    if (Workbook1 != null) Workbook1.Close();

                    if (Excel1.Workbooks.Count == 0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }


                }

                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                if (W3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                if (W4 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W4);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }
            set_enable_true();

        }

        public void load_W1_W2_W3_W4_for_mat_library(string file1, ref Microsoft.Office.Interop.Excel.Application Excel1, ref Microsoft.Office.Interop.Excel.Workbook Workbook1,
                                                                            ref Microsoft.Office.Interop.Excel.Worksheet W1,
                                                                            ref Microsoft.Office.Interop.Excel.Worksheet W2,
                                                                            ref Microsoft.Office.Interop.Excel.Worksheet W3,
                                                                            ref Microsoft.Office.Interop.Excel.Worksheet W4)
        {


            try
            {
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel1.Visible = true;
                foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                {
                    if (Workbook2.FullName == file1)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                        {

                            if (Wx.Name == "MatDesc")
                            {
                                W1 = Wx;
                            }
                            if (Wx.Name == "MatPipe")
                            {
                                W2 = Wx;
                            }
                            if (Wx.Name == "MatPoints")
                            {
                                W3 = Wx;
                            }
                            if (Wx.Name == "MatOther")
                            {
                                W4 = Wx;
                            }
                        }
                    }


                }


            }
            catch (System.Exception)
            {
                Excel1 = new Microsoft.Office.Interop.Excel.Application();

            }

            if (W1 == null)
            {
                Workbook1 = Excel1.Workbooks.Open(file1);
                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                {
                    if (Wx.Name == "MatDesc")
                    {
                        W1 = Wx;
                    }
                    if (Wx.Name == "MatPipe")
                    {
                        W2 = Wx;
                    }
                    if (Wx.Name == "MatPoints")
                    {
                        W3 = Wx;
                    }
                    if (Wx.Name == "MatOther")
                    {
                        W4 = Wx;
                    }
                }
            }


        }

        public List<string> build_category_and_type_list_and_dt_ct()
        {

            List<string> lista1 = new List<string>();
            ct_list = new List<string>();



            for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
            {
                if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value && dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                {
                    string cat1 = Convert.ToString(dt_mat_library.Rows[j][col_category]).ToUpper().Replace(" ", "");
                    string type1 = Convert.ToString(dt_mat_library.Rows[j][col_type]).ToUpper().Replace(" ", "");
                    string descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);

                    string ct1 = (cat1 + "_" + type1).ToUpper().Replace(" ", "");
                    if (cat1.Replace(" ", "").Length > 0)
                    {
                        if (ct_list.Contains(ct1) == false)
                        {
                            ct_list.Add(ct1);

                            Array.Resize(ref dt_ct, ct_list.Count);

                            lista1.Add(ct1);


                        }
                    }
                }
            }

            for (int i = 0; i < ct_list.Count; ++i)
            {
                string ct1 = ct_list[i];

                List<string> lista_mat = new List<string>();

                for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                {
                    if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value && dt_mat_library.Rows[j][col_item_no] != DBNull.Value && dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                    {
                        string cat2 = Convert.ToString(dt_mat_library.Rows[j][col_category]).ToUpper().Replace(" ", "");
                        string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper().Replace(" ", "");
                        string type2 = Convert.ToString(dt_mat_library.Rows[j][col_type]).ToUpper().Replace(" ", "");
                        string ct2 = (cat2 + "_" + type2).ToUpper().Replace(" ", "");

                        if (ct2 == ct1)
                        {
                            lista_mat.Add(mat2);
                        }


                    }
                }

                System.Data.DataTable dt1 = new System.Data.DataTable();
                if (ds_main.dt_points != null || ds_main.dt_extra != null)
                {
                    if (ct1.ToUpper().Replace(" ", "").Contains("POINT") == true)
                    {
                        dt1 = ds_main.dt_points.Clone();
                        for (int k = 0; k < ds_main.dt_points.Rows.Count; ++k)
                        {
                            if (ds_main.dt_points.Rows[k][col_item_no] != DBNull.Value)
                            {
                                string mat2 = Convert.ToString(ds_main.dt_points.Rows[k][col_item_no]).ToUpper().Replace(" ", "");
                                if (lista_mat.Contains(mat2) == true)
                                {
                                    dt1.Rows.Add();
                                    for (int m = 0; m < dt1.Columns.Count; ++m)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][m] = ds_main.dt_points.Rows[k][m];
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        dt1 = ds_main.dt_extra.Clone();
                        for (int k = 0; k < ds_main.dt_extra.Rows.Count; ++k)
                        {
                            if (ds_main.dt_extra.Rows[k][col_item_no] != DBNull.Value)
                            {
                                string mat2 = Convert.ToString(ds_main.dt_extra.Rows[k][col_item_no]).ToUpper().Replace(" ", "");
                                if (lista_mat.Contains(mat2) == true)
                                {
                                    dt1.Rows.Add();
                                    for (int m = 0; m < dt1.Columns.Count; ++m)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][m] = ds_main.dt_extra.Rows[k][m];
                                    }
                                }
                            }
                        }
                    }


                    dt_ct[i] = dt1;

                }
            }
            return lista1;
        }

        public void add_linear_mat_to_combobox(ComboBox combo1)
        {
            int index1 = combo1.Items.IndexOf(combo1.Text);

            combo1.Items.Clear();
            for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
            {
                if (Convert.ToString(dt_mat_library.Rows[i][col_type]).ToUpper() == "LINEAR" && Convert.ToString(dt_mat_library.Rows[i][col_category]).ToUpper() == "PIPE")
                {
                    combo1.Items.Add(dt_mat_library.Rows[i][col_item_no]);
                }
            }

            if (combo1.Items.Count > index1)
            {
                combo1.SelectedIndex = index1;
            }

        }

        public System.Data.DataTable Creaza_dt_points_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(col_mmid, typeof(string));
            dt1.Columns.Add(col_item_no, typeof(string));
            dt1.Columns.Add(col_2dsta, typeof(double));
            dt1.Columns.Add(col_3dsta, typeof(double));
            dt1.Columns.Add(col_eqsta, typeof(double));
            dt1.Columns.Add(col_symbol, typeof(string));
            dt1.Columns.Add(col_altdesc, typeof(string));
            dt1.Columns.Add(col_x, typeof(double));
            dt1.Columns.Add(col_y, typeof(double));
            dt1.Columns.Add(col_block, typeof(string));
            dt1.Columns.Add(col_atr1, typeof(string));
            dt1.Columns.Add(col_atr2, typeof(string));
            dt1.Columns.Add(col_atr3, typeof(string));
            dt1.Columns.Add(col_atr4, typeof(string));
            dt1.Columns.Add(col_visibility, typeof(string));
            dt1.Columns.Add(col_layer, typeof(string));
            dt1.Columns.Add(col_MSblock, typeof(string));
            return dt1;
        }

        public void load_bom(string file1, Microsoft.Office.Interop.Excel.Worksheet W1, Microsoft.Office.Interop.Excel.Worksheet W2, Microsoft.Office.Interop.Excel.Worksheet W3, Microsoft.Office.Interop.Excel.Worksheet W4, bool load_pipes_and_others = false)
        {

            try
            {
                if (W1 != null)
                {
                    ds_main.dt_pipe = Creaza_dt_mat_linear_structure();
                    ds_main.dt_points = Creaza_dt_points_structure();
                    ds_main.dt_extra = Creaza_dt_mat_linear_structure();

                    int start1 = 14;

                    string xl_Item_No = "B";
                    string xl_descr = "C";
                    string xl_Cat = "D";
                    string xl_Type = "E";
                    string xl_Layer = "F";
                    string xl_Block = "G";


                    List<string> lista_col = new List<string>();
                    List<string> lista_colxl = new List<string>();

                    lista_col.Add(col_item_no);
                    lista_col.Add(col_descr);
                    lista_col.Add(col_category);
                    lista_col.Add(col_type);
                    lista_col.Add(col_layer);
                    lista_col.Add(col_MSblock);


                    lista_colxl.Add(xl_Item_No);
                    lista_colxl.Add(xl_descr);
                    lista_colxl.Add(xl_Cat);
                    lista_colxl.Add(xl_Type);
                    lista_colxl.Add(xl_Layer);
                    lista_colxl.Add(xl_Block);


                    dt_mat_library = Functions.build_data_table_from_excel(dt_mat_library, W1, start1, start1 + 1000, lista_col, lista_colxl);

                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        dt_mat_library.Rows[i][col_mmid] = i.ToString();
                    }

                }


                if (load_pipes_and_others == true)
                {
                    if (W2 != null)
                    {
                        int start1 = 14;

                        string xl_col1 = "B";
                        string xl_col2 = "C";
                        string xl_col3 = "D";
                        string xl_col4 = "E";
                        string xl_col5 = "F";
                        string xl_col8 = "I";
                        string xl_col9 = "J";
                        string xl_col11 = "L";
                        string xl_col12 = "M";
                        string xl_col13 = "N";
                        string xl_col14 = "O";
                        string xl_col15 = "V";

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_item_no);
                        lista_col.Add(pipe_col_2d1);
                        lista_col.Add(pipe_col_2d2);
                        lista_col.Add(pipe_col_3d1);
                        lista_col.Add(pipe_col_3d2);
                        lista_col.Add(pipe_col_len2d);
                        lista_col.Add(pipe_col_len3d);
                        lista_col.Add(col_xbeg);
                        lista_col.Add(col_ybeg);
                        lista_col.Add(col_xend);
                        lista_col.Add(col_yend);
                        lista_col.Add(col_layer);

                        lista_colxl.Add(xl_col1);
                        lista_colxl.Add(xl_col2);
                        lista_colxl.Add(xl_col3);
                        lista_colxl.Add(xl_col4);
                        lista_colxl.Add(xl_col5);
                        lista_colxl.Add(xl_col8);
                        lista_colxl.Add(xl_col9);
                        lista_colxl.Add(xl_col11);
                        lista_colxl.Add(xl_col12);
                        lista_colxl.Add(xl_col13);
                        lista_colxl.Add(xl_col14);
                        lista_colxl.Add(xl_col15);

                        ds_main.dt_pipe = Functions.build_data_table_from_excel(ds_main.dt_pipe, W2, start1, start1 + nr_max, lista_col, lista_colxl);
                        if (ds_main.dt_pipe.Rows.Count == 0) ds_main.dt_pipe = null;
                    }

                    if (W3 != null)
                    {
                        int start1 = 13;

                        string xl_colmmid = "A";
                        string xl_colitemno = "B";
                        string xl_col2dsta = "C";
                        string xl_col3dsta = "D";
                        string xl_coleqsta = "E";
                        string xl_colaltdesc = "G";
                        string xl_colx = "H";
                        string xl_coly = "I";
                        string xl_col_block = "P";
                        string xl_col_layer = "Q";

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_item_no);
                        lista_col.Add(col_mmid);
                        lista_col.Add(col_2dsta);
                        lista_col.Add(col_3dsta);
                        lista_col.Add(col_eqsta);
                        lista_col.Add(col_altdesc);
                        lista_col.Add(col_x);
                        lista_col.Add(col_y);
                        lista_col.Add(col_MSblock);
                        lista_col.Add(col_layer);

                        lista_colxl.Add(xl_colitemno);
                        lista_colxl.Add(xl_colmmid);
                        lista_colxl.Add(xl_col2dsta);
                        lista_colxl.Add(xl_col3dsta);
                        lista_colxl.Add(xl_coleqsta);
                        lista_colxl.Add(xl_colaltdesc);
                        lista_colxl.Add(xl_colx);
                        lista_colxl.Add(xl_coly);
                        lista_colxl.Add(xl_col_block);
                        lista_colxl.Add(xl_col_layer);




                        ds_main.dt_points = Functions.build_data_table_from_excel(ds_main.dt_points, W3, start1, start1 + nr_max, lista_col, lista_colxl);

                    }

                    if (W4 != null)
                    {
                        int start1 = 14;

                        string xl_col1 = "B";
                        string xl_col2 = "C";
                        string xl_col3 = "D";
                        string xl_col4 = "E";
                        string xl_col5 = "F";
                        string xl_col8 = "I";
                        string xl_col9 = "J";
                        string xl_col10 = "K";
                        string xl_col11 = "L";
                        string xl_col12 = "M";
                        string xl_col13 = "N";
                        string xl_col14 = "O";
                        string xl_col15 = "V";

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_item_no);
                        lista_col.Add(pipe_col_2d1);
                        lista_col.Add(pipe_col_2d2);
                        lista_col.Add(pipe_col_3d1);
                        lista_col.Add(pipe_col_3d2);
                        lista_col.Add(pipe_col_len2d);
                        lista_col.Add(pipe_col_len3d);
                        lista_col.Add(col_altdesc);
                        lista_col.Add(col_xbeg);
                        lista_col.Add(col_ybeg);
                        lista_col.Add(col_xend);
                        lista_col.Add(col_yend);
                        lista_col.Add(col_layer);

                        lista_colxl.Add(xl_col1);
                        lista_colxl.Add(xl_col2);
                        lista_colxl.Add(xl_col3);
                        lista_colxl.Add(xl_col4);
                        lista_colxl.Add(xl_col5);
                        lista_colxl.Add(xl_col8);
                        lista_colxl.Add(xl_col9);
                        lista_colxl.Add(xl_col10);
                        lista_colxl.Add(xl_col11);
                        lista_colxl.Add(xl_col12);
                        lista_colxl.Add(xl_col13);
                        lista_colxl.Add(xl_col14);
                        lista_colxl.Add(xl_col15);


                        ds_main.dt_extra = Functions.build_data_table_from_excel(ds_main.dt_extra, W4, start1, start1 + nr_max, lista_col, lista_colxl);

                    }
                }




            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }


        }

        public static List<string> Create_Tab_List_From_Table(System.Data.DataTable dt1, string col_name)
        {
            System.Data.DataTable dt2 = new System.Data.DataTable();
            DataView view = new DataView(dt1);
            dt2 = view.ToTable("", false, col_name);

            System.Data.DataTable dt3 = dt2.DefaultView.ToTable(/*distinct*/ true);

            List<string> list1 = new List<string>();

            for (int i = 0; i < dt3.Rows.Count; ++i)
            {
                list1.Add(Convert.ToString(dt3.Rows[i][0]));
            }

            return list1;
        }





        public static DataGridViewTextBoxColumn datagrid_to_datatable_textbox(System.Data.DataTable dt1, string col_name)
        {
            DataGridViewTextBoxColumn tbox_col = new DataGridViewTextBoxColumn();
            tbox_col.HeaderText = col_name;
            tbox_col.DataPropertyName = col_name;
            return tbox_col;
        }
        public static DataGridViewCheckBoxColumn datagrid_to_datatable_checkbox(System.Data.DataTable dt1, string col_name)
        {
            DataGridViewCheckBoxColumn tbox_col = new DataGridViewCheckBoxColumn();
            tbox_col.HeaderText = col_name;
            tbox_col.DataPropertyName = col_name;
            return tbox_col;
        }
        public DataGridViewComboBoxColumn datagrid_to_datatable_combobox_includes_all_items(System.Data.DataTable dt1, string data_to_col, string data_from_col)
        {
            DataGridViewComboBoxColumn cmbox_col = new DataGridViewComboBoxColumn();
            List<string> col_list = new List<string>();
            col_list = Create_Tab_List_From_Table(dt1, data_from_col);
            cmbox_col.DataSource = col_list;
            cmbox_col.HeaderText = data_to_col;
            cmbox_col.DataPropertyName = data_to_col;

            return cmbox_col;
        }

        public static DataGridViewComboBoxColumn datagrid_to_datatable_combobox_one_item(System.Data.DataTable dt1, string data_to_col, string data_from_col)
        {
            DataGridViewComboBoxColumn cmbox_col = new DataGridViewComboBoxColumn();
            List<string> col_list = new List<string>();
            col_list = Create_Tab_List_From_Table(dt1, data_from_col);
            cmbox_col.DataSource = col_list;
            cmbox_col.HeaderText = data_to_col;
            cmbox_col.DataPropertyName = data_to_col;

            return cmbox_col;
        }

        private void button_add_row_to_library_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt_mat_library == null)
                {
                    dt_mat_library = Creaza_mat_library_structure();
                }
                sync_dt_mat_with_filter();

                bool adauga = true;

                if (dt_mat_library.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        bool is_empty = true;
                        for (int j = 0; j < dt_mat_library.Columns.Count; ++j)
                        {
                            if (dt_mat_library.Rows[i][j] != DBNull.Value)
                            {
                                is_empty = false;
                            }
                        }
                        if (is_empty == true) adauga = false;
                    }
                }

                if (adauga == true)
                {
                    dt_mat_library.Rows.Add();
                    dt_mat_library.Rows[dt_mat_library.Rows.Count - 1][col_mmid] = Convert.ToString(dt_mat_library.Rows.Count - 1);


                    if (ct_list != null)
                    {
                        if (ct_list.Count == 1)
                        {
                            if (dt_filter.Rows.Count > 0)
                            {
                                dt_mat_library.Rows[dt_mat_library.Rows.Count - 1][col_category] = dt_filter.Rows[dt_filter.Rows.Count - 1][col_category];
                                dt_mat_library.Rows[dt_mat_library.Rows.Count - 1][col_type] = dt_filter.Rows[dt_filter.Rows.Count - 1][col_type];
                            }
                        }

                        add_tab_pages();
                    }
                    else
                    {
                        dt_filter = dt_mat_library.Copy();
                        display_mat_lib_on_dgv();
                    }
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void button_remove_row_to_library_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataRow row1 = dt_filter.Rows[dt_filter.Rows.Count - 1];

                for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                {
                    System.Data.DataRow row2 = dt_mat_library.Rows[i];

                    bool delete = true;
                    for (int j = 0; j < dt_mat_library.Columns.Count; ++j)
                    {
                        if (row1[j] != row2[j])
                        {
                            delete = false;
                            j = dt_mat_library.Columns.Count;
                        }
                    }

                    if (delete == true)
                    {
                        dt_mat_library.Rows.RemoveAt(i);
                        i = dt_mat_library.Rows.Count;
                    }
                }

                if (dt_filter.Rows.Count == 1)
                {
                    build_category_and_type_list_and_dt_ct();
                }
                add_tab_pages();

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }







        private void button_Filter_Click(object sender, EventArgs e)
        {
            sync_dt_mat_with_filter();
            if (dt_mat_library != null)
            {

                Filter_Box.X = this.MdiParent.Location.X + 911;
                Filter_Box.Y = this.MdiParent.Location.Y + 92;
                Form dialogbox = new Filter_Box();
                dialogbox.ShowDialog();
            }
        }



        private void comboBox_xl_DropDown(object sender, EventArgs e)
        {
            ComboBox combo1 = sender as ComboBox;
            Functions.Load_opened_workbooks_to_combobox(combo1);
            combo1.DropDownWidth = Functions.get_dropdown_width(combo1);

        }
        #endregion



        public void button_refresh_usa_Click(object sender, EventArgs e)
        {
            try
            {


                add_tab_pages();

                System.Data.DataTable dt_cl = ds_main.dt_centerline;
                if (dt_cl == null || dt_cl.Rows.Count == 0) return;

                if (comboBox_default_mat.Text == "")
                {

                    if (ds_main.dt_pipe == null || ds_main.dt_pipe.Rows.Count == 0)
                    {
                        this.MdiParent.WindowState = FormWindowState.Normal;
                        set_enable_true();
                        return;
                    }
                }
                else
                {
                    if ((ds_main.dt_pipe == null || ds_main.dt_pipe.Rows.Count == 0) && dt_cl != null && dt_cl.Rows.Count > 0)
                    {

                        bool is3D = ds_main.is3D;
                        double end_sta = -1;
                        if (is3D == true)
                        {
                            end_sta = Convert.ToDouble(dt_cl.Rows[dt_cl.Rows.Count - 1][col_3dsta]);
                        }
                        else
                        {
                            end_sta = Convert.ToDouble(dt_cl.Rows[dt_cl.Rows.Count - 1][col_2dsta]);
                        }

                        if (end_sta < Math.Round(end_sta, 2))
                        {
                            end_sta = Math.Round(end_sta, 2) - 0.01;
                        }
                        else
                        {
                            end_sta = Math.Round(end_sta, 2);
                        }

                        ds_main.dt_pipe = Creaza_dt_mat_linear_structure();
                        ds_main.dt_pipe.Rows.Add();
                        ds_main.dt_pipe.Rows[0][col_item_no] = comboBox_default_mat.Text;

                        if (is3D == true)
                        {
                            ds_main.dt_pipe.Rows[0][pipe_col_3d1] = 0;
                            ds_main.dt_pipe.Rows[0][pipe_col_3d2] = end_sta;
                            ds_main.dt_pipe.Rows[0][pipe_col_len3d] = end_sta;
                        }
                        else
                        {
                            ds_main.dt_pipe.Rows[0][pipe_col_2d1] = 0;
                            ds_main.dt_pipe.Rows[0][pipe_col_2d2] = end_sta;
                            ds_main.dt_pipe.Rows[0][pipe_col_len2d] = end_sta;
                        }
                    }
                }

                draw_pipes();
                this.MdiParent.WindowState = FormWindowState.Normal;

            }

            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }



        public System.Data.DataTable Creaza_dt_mat_linear_structure()
        {






            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(pipe_col_mmid, typeof(string));
            dt1.Columns.Add(col_item_no, typeof(string));
            dt1.Columns.Add(pipe_col_2d1, typeof(double));
            dt1.Columns.Add(pipe_col_2d2, typeof(double));
            dt1.Columns.Add(pipe_col_3d1, typeof(double));
            dt1.Columns.Add(pipe_col_3d2, typeof(double));
            dt1.Columns.Add(pipe_col_eq1, typeof(double));
            dt1.Columns.Add(pipe_col_eq2, typeof(double));
            dt1.Columns.Add(pipe_col_len2d, typeof(double));
            dt1.Columns.Add(pipe_col_len3d, typeof(double));
            dt1.Columns.Add(col_altdesc, typeof(string));
            dt1.Columns.Add(pipe_col11, typeof(double));
            dt1.Columns.Add(pipe_col12, typeof(double));
            dt1.Columns.Add(pipe_col13, typeof(double));
            dt1.Columns.Add(pipe_col14, typeof(double));
            dt1.Columns.Add(pipe_col_block, typeof(string));
            dt1.Columns.Add(pipe_col_mat, typeof(string));
            dt1.Columns.Add(pipe_col17, typeof(string));
            dt1.Columns.Add(pipe_col18, typeof(string));
            dt1.Columns.Add(pipe_col19, typeof(string));
            dt1.Columns.Add(pipe_col20, typeof(string));
            dt1.Columns.Add(col_layer, typeof(string));


            return dt1;
        }




        private void delete_existing_linear_curves(List<string> list_del)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            set_enable_false();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    foreach (ObjectId id1 in BTrecord)
                    {
                        Curve curve1 = Trans1.GetObject(id1, OpenMode.ForRead) as Curve;
                        if (curve1 != null)
                        {
                            if (list_del.Contains(curve1.Layer) == true)
                            {
                                curve1.UpgradeOpen();
                                curve1.Erase();
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }





        private void button_pipe_pick_pts_Click(object sender, System.EventArgs e)
        {
            System.Data.DataTable dt_cl = ds_main.dt_centerline;




            if (dt_cl == null || dt_cl.Rows.Count < 2)
            {
                MessageBox.Show("No centerline loaded\r\nOperation aborted", "material design", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (comboBox_default_mat.Text == "")
            {
                MessageBox.Show("Specify the default material\r\nOperation aborted", "material design", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (comboBox_pipe_mat.Text == "")
            {
                MessageBox.Show("Specify material\r\nOperation aborted", "material design", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    ObjectId p3did = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(dt_cl);
                        Trans1.Commit();

                        p3did = Poly3D.ObjectId;

                    }

                    System.Data.DataTable dt_od_pipe = new System.Data.DataTable();
                    dt_od_pipe.Columns.Add(pipe_us_od_item_no, typeof(string));
                    dt_od_pipe.Columns.Add(pipe_us_od_descr, typeof(string));
                    dt_od_pipe.Columns.Add(pipe_us_od_cat, typeof(string));
                    dt_od_pipe.Columns.Add(pipe_us_od_sta1, typeof(string));
                    dt_od_pipe.Columns.Add(pipe_us_od_sta2, typeof(string));
                    dt_od_pipe.Columns.Add("id", typeof(ObjectId));

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        this.MdiParent.WindowState = FormWindowState.Minimized;
                        Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);
                        Polyline3d Poly3D = Trans1.GetObject(p3did, OpenMode.ForWrite) as Polyline3d;

                        if (Poly3D == null)
                        {
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        ObjectIdCollection objectid_col = new ObjectIdCollection();
                        objectid_col.Add(Poly3D.ObjectId);
                        DrawOrderTable DrawOrderTable1 = Trans1.GetObject(BTrecord.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                        DrawOrderTable1.MoveToBottom(objectid_col);





                        //Point3d view_center = Functions.get_view_center();

                        //Point3d pt_on_cl = Poly2D.GetClosestPointTo(view_center, Vector3d.ZAxis, false);

                        //double dist1 = Math.Pow(Math.Pow(view_center.X - pt_on_cl.X, 2) + Math.Pow(view_center.Y - pt_on_cl.Y, 2), 0.5);
                        //if (dist1 > Poly2D.Length)
                        //{
                        //    Functions.zoom_to_Point(pt_on_cl, 1);
                        //}



                        double end_sta = Math.Round(Poly3D.Length, 2);
                        if (end_sta > Poly3D.Length)
                        {
                            end_sta = end_sta - 0.01;
                        }

                        bool is3D = ds_main.is3D;

                        if (ds_main.dt_pipe == null || ds_main.dt_pipe.Rows.Count == 0)
                        {
                            ds_main.dt_pipe = Creaza_dt_mat_linear_structure();
                            ds_main.dt_pipe.Rows.Add();
                            ds_main.dt_pipe.Rows[0][col_item_no] = comboBox_default_mat.Text;

                            if (is3D == false)
                            {
                                ds_main.dt_pipe.Rows[0][pipe_col_2d1] = 0;
                                ds_main.dt_pipe.Rows[0][pipe_col_2d2] = end_sta;
                                ds_main.dt_pipe.Rows[0][pipe_col_len2d] = end_sta;
                                ds_main.dt_pipe.Rows[0][col_xbeg] = Poly2D.StartPoint.X;
                                ds_main.dt_pipe.Rows[0][col_ybeg] = Poly2D.StartPoint.Y;
                                ds_main.dt_pipe.Rows[0][col_xend] = Poly2D.EndPoint.X;
                                ds_main.dt_pipe.Rows[0][col_yend] = Poly2D.EndPoint.Y;
                            }
                            else
                            {
                                ds_main.dt_pipe.Rows[0][pipe_col_3d1] = 0;
                                ds_main.dt_pipe.Rows[0][pipe_col_3d2] = end_sta;
                                ds_main.dt_pipe.Rows[0][pipe_col_len3d] = end_sta;
                                ds_main.dt_pipe.Rows[0][col_xbeg] = Poly3D.StartPoint.X;
                                ds_main.dt_pipe.Rows[0][col_ybeg] = Poly3D.StartPoint.Y;
                                ds_main.dt_pipe.Rows[0][col_xend] = Poly3D.EndPoint.X;
                                ds_main.dt_pipe.Rows[0][col_yend] = Poly3D.EndPoint.Y;
                            }
                        }

                        //Trans1.TransactionManager.QueueForGraphicsFlush();

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions pp1;
                        pp1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify start point");
                        pp1.AllowNone = false;

                        pp1.Keywords.Add("Middle");
                        pp1.Keywords.Add("Buffer from point");
                        pp1.Keywords.Add("Feature");

                        Point_res1 = Editor1.GetPoint(pp1);

                        if (Point_res1.Status != PromptStatus.OK && Point_res1.Status != PromptStatus.Keyword)
                        {
                            Poly3D.Erase();
                            Trans1.Commit();
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        double par1 = -1;
                        double par2 = -1;
                        double station_start = -1;
                        double station_end = -1;
                        if (Point_res1.Status == PromptStatus.Keyword)
                        {
                            #region keyword middle
                            PromptPointOptions ppm;
                            ppm = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify mid point:");
                            ppm.AllowNone = false;

                            if (Point_res1.StringResult.ToLower() == "middle")
                            {
                                PromptPointResult Point_resm = Editor1.GetPoint(ppm);
                                if (Point_resm.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }
                                Point3d point_middle = Point_resm.Value;



                                Autodesk.AutoCAD.EditorInput.PromptDistanceOptions Prompt_len = new Autodesk.AutoCAD.EditorInput.PromptDistanceOptions("\n" + "Specify length:");
                                Prompt_len.AllowNegative = false;
                                Prompt_len.AllowZero = true;
                                Prompt_len.AllowNone = true;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_len = ThisDrawing.Editor.GetDistance(Prompt_len);
                                if (Rezultat_len.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                double len1 = Rezultat_len.Value;
                                if (len1 < 1)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }
                                Point3d p1 = Poly2D.GetClosestPointTo(point_middle, Vector3d.ZAxis, false);



                                double param_m = Poly2D.GetParameterAtPoint(p1);
                                if (param_m > Poly3D.EndParam) param_m = Poly3D.EndParam;

                                double sta_m = Poly3D.GetDistanceAtParameter(param_m);
                                station_start = sta_m - len1 / 2;
                                station_end = sta_m + len1 / 2;

                                station_start = Math.Round(station_start, 0);
                                station_end = Math.Round(station_end, 0);
                                if (station_start < 0) station_start = 0;

                                if (sta_m + len1 / 2 == Poly3D.Length)
                                {
                                    station_end = Poly3D.Length - 0.0001;
                                }

                            }
                            #endregion

                            #region keyword buffer from point
                            if (Point_res1.StringResult.ToLower() == "buffer from point")
                            {
                                PromptPointOptions ppb1;
                                ppb1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify first point:");
                                ppb1.AllowNone = false;

                                PromptPointOptions ppb2;
                                ppb2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify second point:");
                                ppb2.AllowNone = false;


                                PromptPointResult Point_resb1 = Editor1.GetPoint(ppb1);
                                if (Point_resb1.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }
                                Point3d point_b1 = Point_resb1.Value;



                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_b1 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify buffer value:");
                                Prompt_b1.AllowNegative = false;
                                Prompt_b1.AllowZero = true;
                                Prompt_b1.AllowNone = true;
                                Prompt_b1.UseDefaultValue = true;
                                Prompt_b1.DefaultValue = 0;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_b1 = ThisDrawing.Editor.GetDouble(Prompt_b1);
                                if (Rezultat_b1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }



                                PromptPointResult Point_resb2 = Editor1.GetPoint(ppb2);
                                if (Point_resb2.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }
                                Point3d point_b2 = Point_resb2.Value;



                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_b2 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify buffer value:");
                                Prompt_b2.AllowNegative = false;
                                Prompt_b2.AllowZero = true;
                                Prompt_b2.AllowNone = true;
                                Prompt_b2.UseDefaultValue = true;
                                Prompt_b2.DefaultValue = 0;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_b2 = ThisDrawing.Editor.GetDouble(Prompt_b2);
                                if (Rezultat_b2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                double buffer1 = Rezultat_b1.Value;
                                Point3d p1 = Poly2D.GetClosestPointTo(point_b1, Vector3d.ZAxis, false);
                                double param_b1 = Poly2D.GetParameterAtPoint(p1);
                                if (param_b1 > Poly3D.EndParam) param_b1 = Poly3D.EndParam;
                                double sta_b1 = Poly3D.GetDistanceAtParameter(param_b1);




                                double buffer2 = Rezultat_b2.Value;
                                Point3d p2 = Poly2D.GetClosestPointTo(point_b2, Vector3d.ZAxis, false);
                                double param_b2 = Poly2D.GetParameterAtPoint(p2);
                                if (param_b2 > Poly3D.EndParam) param_b2 = Poly3D.EndParam;
                                double sta_b2 = Poly3D.GetDistanceAtParameter(param_b2);


                                if (sta_b2 > sta_b1)
                                {
                                    station_start = sta_b1 - buffer1;
                                    station_end = sta_b2 + buffer2;
                                }
                                else
                                {
                                    station_start = sta_b2 - buffer2;
                                    station_end = sta_b1 + buffer1;
                                }



                                station_start = Math.Round(station_start, 0);
                                station_end = Math.Round(station_end, 0);

                                if (station_start < 0) station_start = 0;


                                if (sta_b2 > sta_b1)
                                {
                                    if (sta_b2 + buffer2 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }
                                else
                                {
                                    if (sta_b1 + buffer1 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }



                            }
                            #endregion

                            #region keyword feature


                            if (Point_res1.StringResult.ToLower() == "feature")
                            {
                                Point3d point_feat1 = new Point3d();
                                Point3d point_feat2 = new Point3d();

                                Autodesk.AutoCAD.EditorInput.PromptEntityResult rezultat_feat1;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_feat1;
                                prompt_feat1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the feature:");
                                prompt_feat1.SetRejectMessage("\nSelect a polyline!");
                                prompt_feat1.AllowNone = true;
                                prompt_feat1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                rezultat_feat1 = ThisDrawing.Editor.GetEntity(prompt_feat1);

                                if (rezultat_feat1.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }

                                Polyline poly1 = Trans1.GetObject(rezultat_feat1.ObjectId, OpenMode.ForWrite) as Polyline;
                                poly1.Elevation = 0;

                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly1, Poly2D);
                                Point3dCollection colint2 = new Point3dCollection();

                                if (colint1.Count == 0)
                                {

                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "no intersection");
                                    set_enable_true();
                                    return;

                                }
                                else if (colint1.Count == 1)
                                {
                                    point_feat1 = colint1[0];

                                    Autodesk.AutoCAD.EditorInput.PromptEntityResult rezultat_feat2;
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_feat2;
                                    prompt_feat2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the other feature:");
                                    prompt_feat2.SetRejectMessage("\nSelect a polyline!");
                                    prompt_feat2.AllowNone = true;
                                    prompt_feat2.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                    rezultat_feat2 = ThisDrawing.Editor.GetEntity(prompt_feat2);

                                    if (rezultat_feat2.Status != PromptStatus.OK)
                                    {

                                        PromptPointOptions ppf2;
                                        ppf2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify second point:");
                                        ppf2.AllowNone = false;

                                        PromptPointResult Point_resf2 = Editor1.GetPoint(ppf2);
                                        if (Point_resf2.Status != PromptStatus.OK)
                                        {
                                            this.WindowState = FormWindowState.Normal;
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            set_enable_true();
                                            return;
                                        }
                                        Point3d point_f2 = Point_resf2.Value;
                                        point_feat2 = Poly2D.GetClosestPointTo(point_f2, Vector3d.ZAxis, false);
                                    }
                                    else
                                    {
                                        Polyline poly2 = Trans1.GetObject(rezultat_feat2.ObjectId, OpenMode.ForWrite) as Polyline;
                                        poly2.Elevation = 0;
                                        colint2 = Functions.Intersect_on_both_operands(poly2, Poly2D);
                                        if (colint2.Count == 0)
                                        {

                                            this.WindowState = FormWindowState.Normal;
                                            ThisDrawing.Editor.WriteMessage("\n" + "no intersection");
                                            set_enable_true();
                                            return;

                                        }
                                        else if (colint2.Count == 1)
                                        {
                                            point_feat2 = colint2[0];
                                        }
                                        else if (colint2.Count > 1)
                                        {
                                            PromptPointOptions pp_int2;
                                            pp_int2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify intersection point:");
                                            pp_int2.AllowNone = false;

                                            PromptPointResult rezult_close2 = Editor1.GetPoint(pp_int2);
                                            if (rezult_close2.Status != PromptStatus.OK)
                                            {
                                                this.WindowState = FormWindowState.Normal;
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                set_enable_true();
                                                return;
                                            }
                                            Point3d point_close2 = rezult_close2.Value;

                                            double dmin = 1000;
                                            for (int i = 0; i < colint2.Count; ++i)
                                            {
                                                Point3d pint2 = colint2[i];

                                                double d1 = Math.Pow(Math.Pow(point_close2.X - pint2.X, 2) + Math.Pow(point_close2.Y - pint2.Y, 2), 0.5);

                                                if (d1 < dmin)
                                                {
                                                    dmin = d1;
                                                    point_feat2 = pint2;
                                                }


                                            }




                                        }
                                    }



                                }
                                else if (colint1.Count == 2)
                                {
                                    point_feat1 = colint1[0];
                                    point_feat2 = colint1[1];

                                }
                                else if (colint1.Count > 2)
                                {
                                    PromptPointOptions pp_int1;
                                    pp_int1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify intersection point #1:");
                                    pp_int1.AllowNone = false;

                                    PromptPointResult rezult_close1 = Editor1.GetPoint(pp_int1);
                                    if (rezult_close1.Status != PromptStatus.OK)
                                    {
                                        this.WindowState = FormWindowState.Normal;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        set_enable_true();
                                        return;
                                    }
                                    Point3d point_close1 = rezult_close1.Value;

                                    double dmin1 = 1000;
                                    for (int i = 0; i < colint1.Count; ++i)
                                    {
                                        Point3d pint1 = colint1[i];

                                        double d1 = Math.Pow(Math.Pow(point_close1.X - pint1.X, 2) + Math.Pow(point_close1.Y - pint1.Y, 2), 0.5);

                                        if (d1 < dmin1)
                                        {
                                            dmin1 = d1;
                                            point_feat1 = pint1;
                                        }


                                    }


                                    PromptPointOptions pp_int2;
                                    pp_int2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify intersection point #2:");
                                    pp_int2.AllowNone = false;

                                    PromptPointResult rezult_close2 = Editor1.GetPoint(pp_int2);
                                    if (rezult_close2.Status != PromptStatus.OK)
                                    {
                                        this.WindowState = FormWindowState.Normal;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        set_enable_true();
                                        return;
                                    }
                                    Point3d point_close2 = rezult_close2.Value;

                                    double dmin2 = 1000;
                                    for (int i = 0; i < colint1.Count; ++i)
                                    {
                                        Point3d pint2 = colint1[i];

                                        double d2 = Math.Pow(Math.Pow(point_close2.X - pint2.X, 2) + Math.Pow(point_close2.Y - pint2.Y, 2), 0.5);

                                        if (d2 < dmin2)
                                        {
                                            dmin2 = d2;
                                            point_feat2 = pint2;
                                        }


                                    }

                                }








                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions prompt_buffer = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify buffer value:");
                                prompt_buffer.AllowNegative = false;
                                prompt_buffer.AllowZero = true;
                                prompt_buffer.AllowNone = true;
                                prompt_buffer.UseDefaultValue = true;
                                prompt_buffer.DefaultValue = 0;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult rez_buffer = ThisDrawing.Editor.GetDouble(prompt_buffer);
                                if (rez_buffer.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }








                                double buffer1 = rez_buffer.Value;

                                double param_b1 = Poly2D.GetParameterAtPoint(point_feat1);
                                if (param_b1 > Poly3D.EndParam) param_b1 = Poly3D.EndParam;


                                double sta_b1 = Poly3D.GetDistanceAtParameter(param_b1);


                                double param_b2 = Poly2D.GetParameterAtPoint(point_feat2);
                                if (param_b2 > Poly3D.EndParam) param_b2 = Poly3D.EndParam;
                                double sta_b2 = Poly3D.GetDistanceAtParameter(param_b2);


                                if (sta_b2 > sta_b1)
                                {
                                    station_start = sta_b1 - buffer1;
                                    station_end = sta_b2 + buffer1;
                                }
                                else
                                {
                                    station_start = sta_b2 - buffer1;
                                    station_end = sta_b1 + buffer1;
                                }



                                station_start = Math.Round(station_start, 0);
                                station_end = Math.Round(station_end, 0);

                                if (station_start < 0) station_start = 0;


                                if (sta_b2 > sta_b1)
                                {
                                    if (sta_b2 + buffer1 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }
                                else
                                {
                                    if (sta_b1 + buffer1 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }



                            }
                            #endregion


                        }
                        else
                        {
                            #region pick 2 pts
                            Point3d p1 = Poly2D.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);

                            par1 = Poly2D.GetParameterAtPoint(p1);
                            if (par1 > Poly3D.EndParam) par1 = Poly3D.EndParam;


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions pp2;
                            pp2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify end point");
                            pp2.AllowNone = false;
                            pp2.UseBasePoint = true;
                            pp2.BasePoint = p1;


                            Point_res2 = Editor1.GetPoint(pp2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                Poly3D.Erase();
                                Trans1.Commit();
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                return;
                            }

                            Point3d p2 = Poly2D.GetClosestPointTo(Point_res2.Value, Vector3d.ZAxis, false);
                            par2 = Poly2D.GetParameterAtPoint(p2);
                            if (par2 > Poly3D.EndParam) par2 = Poly3D.EndParam;

                            station_start = Math.Round(Poly3D.GetDistanceAtParameter(par1), 0);
                            station_end = Math.Round(Poly3D.GetDistanceAtParameter(par2), 0);
                            if (par2 == Poly3D.EndParam)
                            {
                                station_end = Poly3D.Length - 0.0001;
                            }
                            #endregion

                        }




                        if (station_start >= 0 && station_end >= 0)
                        {


                            if (station_start > end_sta)
                            {
                                station_start = end_sta;
                            }

                            if (station_end > end_sta)
                            {
                                station_end = end_sta;
                            }

                            Point3d pt1 = Poly3D.GetPointAtDist(station_start);
                            Point3d pt2 = Poly3D.GetPointAtDist(station_end);



                            insert_us_mat(ref ds_main.dt_pipe, station_start, station_end, pt1, pt2);
                        }




                        Poly3D.Erase();
                        Trans1.Commit();
                    }


                }


                draw_pipes();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        private void draw_pipes()
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            System.Data.DataTable dt_cl = ds_main.dt_centerline;



            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {

                ObjectId p3did = ObjectId.Null;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(dt_cl);
                    Trans1.Commit();

                    p3did = Poly3D.ObjectId;

                }

                System.Data.DataTable dt_od_pipe = new System.Data.DataTable();
                dt_od_pipe.Columns.Add(pipe_us_od_item_no, typeof(string));
                dt_od_pipe.Columns.Add(pipe_us_od_descr, typeof(string));
                dt_od_pipe.Columns.Add(pipe_us_od_cat, typeof(string));
                dt_od_pipe.Columns.Add(pipe_us_od_sta1, typeof(string));
                dt_od_pipe.Columns.Add(pipe_us_od_sta2, typeof(string));
                dt_od_pipe.Columns.Add("id", typeof(ObjectId));

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                    this.MdiParent.WindowState = FormWindowState.Minimized;
                    Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);
                    Polyline3d Poly3D = Trans1.GetObject(p3did, OpenMode.ForWrite) as Polyline3d;

                    if (Poly3D == null)
                    {
                        return;
                    }


                    List<string> list_del = new List<string>();
                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        if (dt_mat_library.Rows[i][col_layer] == DBNull.Value) dt_mat_library.Rows[i][col_layer] = "MAT_Pipe";
                        if (dt_mat_library.Rows[i][col_type] != DBNull.Value && dt_mat_library.Rows[i][col_category] != DBNull.Value && dt_mat_library.Rows[i][col_item_no] != DBNull.Value)
                        {
                            if (Convert.ToString(dt_mat_library.Rows[i][col_type]).ToUpper() == "LINEAR" && Convert.ToString(dt_mat_library.Rows[i][col_category]).ToUpper() == "PIPE")
                            {
                                string layer1 = Convert.ToString(dt_mat_library.Rows[i][col_layer]);
                                if (list_del.Contains(layer1) == false) list_del.Add(layer1);
                            }
                        }
                    }

                    if (extra_layer_to_be_deleted != "")
                    {
                        if (list_del.Contains(extra_layer_to_be_deleted) == false)
                        {
                            list_del.Add(extra_layer_to_be_deleted);
                        }
                        extra_layer_to_be_deleted = "";
                    }



                    delete_existing_linear_curves(list_del);

                    create_pipe_us_od_table(pipes_od);


                    double end_sta = Math.Round(Poly3D.Length, 2);
                    if (end_sta > Poly3D.Length)
                    {
                        end_sta = end_sta - 0.01;
                    }

                    bool is3D = ds_main.is3D;



                    for (int i = 0; i < ds_main.dt_pipe.Rows.Count; ++i)
                    {
                        double sta1 = -1;
                        double sta2 = -1;

                        if (is3D == true)
                        {
                            sta1 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_3d1]);
                            sta2 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_3d2]);
                        }
                        else
                        {
                            sta1 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_2d1]);
                            sta2 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_2d2]);
                        }

                        if (sta2 > end_sta) sta2 = end_sta;

                        double param1 = Poly3D.GetParameterAtDistance(sta1);
                        double param2 = Poly3D.GetParameterAtDistance(sta2);
                        if (param2 > Poly2D.EndParam) param2 = Poly2D.EndParam;

                        string mat1 = Convert.ToString(ds_main.dt_pipe.Rows[i][col_item_no]);
                        string layer1 = "MAT_Pipe";
                        string descr1 = "none";
                        string cat1 = "Pipe";

                        for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                        {

                            if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                            {
                                string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]);

                                string layer2 = Convert.ToString(dt_mat_library.Rows[j][col_layer]);

                                if (mat1 == mat2)
                                {
                                    layer1 = layer2;
                                    if (dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                                    {
                                        descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                                    }

                                    if (dt_mat_library.Rows[j][col_category] != DBNull.Value)
                                    {
                                        cat1 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                                    }

                                }
                            }
                        }

                        short cid = 1;

                        switch (mat1)
                        {
                            case "1":
                                cid = 9;
                                break;
                            case "2":
                                cid = 5;
                                break;
                            case "3":
                                cid = 4;
                                break;
                            case "4":
                                cid = 3;
                                break;
                            case "5":
                                cid = 2;
                                break;
                            case "6":
                                cid = 1;
                                break;
                            default:
                                cid = 6;
                                break;
                        }

                        Functions.Creaza_layer(layer1, cid, true);
                        Polyline poly1 = Functions.get_part_of_poly(Poly2D, param1, param2);
                        poly1.Layer = layer1;

                        BTrecord.AppendEntity(poly1);
                        Trans1.AddNewlyCreatedDBObject(poly1, true);

                        ds_main.dt_pipe.Rows[i][col_xbeg] = poly1.StartPoint.X;
                        ds_main.dt_pipe.Rows[i][col_ybeg] = poly1.StartPoint.Y;
                        ds_main.dt_pipe.Rows[i][col_xend] = poly1.EndPoint.X;
                        ds_main.dt_pipe.Rows[i][col_yend] = poly1.EndPoint.Y;

                        ds_main.dt_pipe.Rows[i][pipe_col_mmid] = poly1.ObjectId.Handle.Value.ToString();


                        dt_od_pipe.Rows.Add();
                        dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1]["id"] = poly1.ObjectId;
                        dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][pipe_us_od_sta1] = sta1;
                        dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][pipe_us_od_sta2] = sta2;
                        dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][pipe_us_od_descr] = descr1;
                        dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][pipe_us_od_item_no] = mat1;
                        dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][pipe_us_od_cat] = cat1;
                    }

                    Poly3D.Erase();
                    Trans1.Commit();
                }

                attach_od_to_us_pipes(dt_od_pipe);
            }



            populate_datagridview_pipe();

        }




        private void button_save_pipe_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                get_client_project_segment_pipe_diam();
                dataGridView_mat_library.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                Transfer_dt_pipe_to_file1();
                textBox_library.Text = ds_main.config_xls;
                textBox_library.ForeColor = Color.LightGreen;
                textBox_library.Font = font8;
                textBox_library.TextAlign = HorizontalAlignment.Right;
                set_enable_true();
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        public void Transfer_dt_pipe_to_file1()
        {



            sync_dt_mat_with_filter();


            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

            bool is_opened = false;


            try
            {
                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {


                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            if (Workbook2.FullName == ds_main.config_xls)
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                                {

                                    Workbook1 = Workbook2;
                                    if (Wx.Name == "MatPipe")
                                    {
                                        W1 = Wx;
                                    }
                                    is_opened = true;
                                }

                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W1.Name = "MatPipe";
                                }

                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }


                    if (is_opened == false)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                        Workbook1 = Excel1.Workbooks.Open(ds_main.config_xls);

                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name == "MatPipe")
                            {
                                W1 = Wx;
                            }
                        }

                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W1.Name = "MatPipe";
                        }
                    }




                }

                if (ds_main.dt_pipe != null && W1 != null)
                {
                    if (ds_main.dt_pipe.Rows.Count > 0)
                    {
                        Create_header_material_linear_file(W1, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_pipe);

                        int last_row = nr_max + 14;
                        W1.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_pipe.Rows.Count;
                        int maxCols = ds_main.dt_pipe.Columns.Count;
                        W1.Range["A14:V" + last_row.ToString()].ClearContents();
                        W1.Range["A14:V" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A14:V" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_pipe.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_pipe.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }
                    else
                    {
                        W1.Cells.Clear();
                        W1.Cells.ClearFormats();
                    }


                }
                else if (W1 != null)
                {
                    W1.Cells.Clear();
                    W1.Cells.ClearFormats();
                }

                if (is_opened == false)
                {

                    Workbook1.Save();
                    Workbook1.Close();
                }

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }

        public void transfer_dt_extra_to_file1()
        {



            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

            bool is_opened = false;

            try
            {
                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {


                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            if (Workbook2.FullName == ds_main.config_xls)
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                                {

                                    Workbook1 = Workbook2;
                                    if (Wx.Name == "MatOther")
                                    {
                                        W1 = Wx;
                                    }
                                    is_opened = true;
                                }
                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W1.Name = "MatOther";
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }


                    if (is_opened == false)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                        Workbook1 = Excel1.Workbooks.Open(ds_main.config_xls);
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name == "MatOther")
                            {
                                W1 = Wx;
                            }
                        }

                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W1.Name = "MatOther";
                        }
                    }

                }

                if (ds_main.dt_extra != null && W1 != null)
                {
                    if (ds_main.dt_extra.Rows.Count > 0)
                    {
                        Create_header_material_linear_file(W1, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_extra, "Material Linear Other");

                        int last_row = nr_max + 14;
                        W1.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_extra.Rows.Count;
                        int maxCols = ds_main.dt_extra.Columns.Count;
                        W1.Range["A14:V" + last_row.ToString()].ClearContents();
                        W1.Range["A14:V" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A14:V" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_extra.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_extra.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }
                    else
                    {
                        W1.Cells.Clear();
                        W1.Cells.ClearFormats();
                    }
                }
                else if (W1 != null)
                {
                    W1.Cells.Clear();
                    W1.Cells.ClearFormats();
                }

                if (is_opened == false)
                {
                    Workbook1.Save();
                    Workbook1.Close();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }

        public static void Create_header_material_linear_file(Microsoft.Office.Interop.Excel.Worksheet W1, string Client, string Project, string Segment, System.Data.DataTable dt_lin, string title1 = "Material Pipe")
        {



            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B11"];
            object[,] valuesH = new object[11, 2];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";

            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at: " + DateTime.Now.TimeOfDay;
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;
            valuesH[6, 0] = "JOINT LENGTH";
            valuesH[7, 0] = "Engineering is responsible for the content and QAQC of this table.";
            valuesH[8, 0] = "This Table shows Material that has a Begin and End Station";
            valuesH[9, 0] = "Do not add any columns to this table, also do not add any rows above row 13";
            valuesH[10, 0] = "This Table is to be populated by Engineering and Programming as indicated below:";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:B7"];

            Functions.Color_border_range_inside(range1, 46);

            range1 = W1.Range["A8:V8"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A9:V9"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A10:V10"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 3); //red

            range1 = W1.Range["A11:V11"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 43); //green

            range1 = W1.Range["A12:V12"];
            object[,] values12 = new object[1, 21];
            values12[0, 0] = "N/A";
            values12[0, 1] = "ENG";
            values12[0, 2] = "ENG";
            values12[0, 3] = "ENG";
            values12[0, 4] = "ENG";
            values12[0, 5] = "ENG";
            values12[0, 6] = "PROGRAM";
            values12[0, 7] = "PROGRAM";
            values12[0, 8] = "PROGRAM";
            values12[0, 9] = "PROGRAM";
            values12[0, 10] = "ENG";
            values12[0, 11] = "USER";
            values12[0, 12] = "USER";
            values12[0, 13] = "USER";
            values12[0, 14] = "USER";
            values12[0, 15] = "USER";
            values12[0, 16] = "USER";
            values12[0, 17] = "USER";
            values12[0, 18] = "USER";
            values12[0, 19] = "USER";
            values12[0, 20] = "USER";

            range1.Value2 = values12;
            Functions.Color_border_range_inside(range1, 43); //green

            range1 = W1.Range["C1:O7"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = title1;
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Functions.Color_border_range_outside(range1, 0);


            W1.Range["A7"].Font.Bold = true;

            range1 = W1.Range["A13:U13"];
            object[,] values_dt = new object[1, dt_lin.Columns.Count];
            if (dt_lin != null && dt_lin.Columns.Count > 0)
            {
                for (int i = 0; i < dt_lin.Columns.Count; ++i)
                {
                    values_dt[0, i] = dt_lin.Columns[i].ColumnName;
                }
                range1.Value2 = values_dt;
                Functions.Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }

            W1.Range["A:B"].ColumnWidth = 14;
            W1.Range["C:V"].ColumnWidth = 10;
            object[,] values_legend = new object[6, 1];
            values_legend[0, 0] = "attributes mandatory naming:";
            values_legend[1, 0] = "STA1, STA2, LEN";
            values_legend[2, 0] = "Between Block and Visibility columns you may add/remove any number of columns";
            values_legend[3, 0] = "Attributes names has to be unique (no duplicates in naming)";
            values_legend[4, 0] = "Attributes used by mat count:";
            values_legend[5, 0] = "QTY, LENGTH, LEN, STA1, STA2, MAT";
            W1.Range["P1:P6"].Value2 = values_legend;

            W1.Range["A:B"].ColumnWidth = 13;
            W1.Range["C:J"].ColumnWidth = 10;
            W1.Range["K:O"].ColumnWidth = 6;
            W1.Range["P:V"].ColumnWidth = 18;

        }



        public static void Create_header_material_points_file(Microsoft.Office.Interop.Excel.Worksheet W1, string Client, string Project, string Segment, System.Data.DataTable dt_pts)
        {
            W1.Range["A:B"].ColumnWidth = 13;
            W1.Range["C:E"].ColumnWidth = 10;
            W1.Range["F:I"].ColumnWidth = 6;
            W1.Range["J:Q"].ColumnWidth = 18;

            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B11"];
            Object[,] valuesH = new object[10, 2];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";

            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at: " + DateTime.Now.TimeOfDay;
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;
            valuesH[6, 0] = "Engineering is responsible for the content and QAQC of this table.";
            valuesH[7, 0] = "This Table shows Material that is identified as a single point";
            valuesH[8, 0] = "Do not add any columns to this table between columns A and J, also do not add any rows above row 13";
            valuesH[9, 0] = "This Table is to be populated by Engineering and Programming as indicated below:";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:B6"];

            Functions.Color_border_range_inside(range1, 46);

            range1 = W1.Range["A7:Q7"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A8:Q8"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A9:Q9"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 3); //red


            range1 = W1.Range["A10:Q10"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 43); //green

            range1 = W1.Range["A11:Q11"];
            object[,] values12 = new object[1, 15];
            values12[0, 0] = "N/A";
            values12[0, 1] = "ENG";
            values12[0, 2] = "ENG";
            values12[0, 3] = "ENG";
            values12[0, 4] = "PROGRAM";
            values12[0, 5] = "PROGRAM";
            values12[0, 6] = "ENG";
            values12[0, 7] = "USER";
            values12[0, 8] = "USER";
            values12[0, 9] = "USER";
            values12[0, 10] = "USER";
            values12[0, 11] = "USER";
            values12[0, 12] = "USER";
            values12[0, 13] = "USER";
            values12[0, 14] = "USER";


            range1.Value2 = values12;
            Functions.Color_border_range_inside(range1, 43); //green

            range1 = W1.Range["C1:I6"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "Material as points";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Functions.Color_border_range_outside(range1, 0);

            range1 = W1.Range["A12:Q12"];
            object[,] values_dt = new object[1, dt_pts.Columns.Count];
            if (dt_pts != null && dt_pts.Columns.Count > 0)
            {
                for (int i = 0; i < dt_pts.Columns.Count; ++i)
                {
                    values_dt[0, i] = dt_pts.Columns[i].ColumnName;
                }
                range1.Value2 = values_dt;
                Functions.Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }

            W1.Range["A:B"].ColumnWidth = 14;
            W1.Range["C:Q"].ColumnWidth = 10;
            object[,] values_legend = new object[6, 1];
            values_legend[0, 0] = "attributes mandatory naming:";
            values_legend[1, 0] = "STA, MAT";
            values_legend[2, 0] = "Between Block and Visibility columns you may add/remove any number of columns";
            values_legend[3, 0] = "Attributes names has to be unique (no duplicates in naming)";
            values_legend[4, 0] = "Attributes used by mat count:";
            values_legend[5, 0] = "QTY, LENGTH, LEN, MAT";
            W1.Range["J1:J6"].Value2 = values_legend;

        }




        private void insert_us_mat(ref System.Data.DataTable dt1, double sm1, double sm2, Point3d pt1, Point3d pt2)
        {
            bool is3D = ds_main.is3D;

            if (dt1 != null && dt1.Rows.Count > 0)
            {

                if (sm1 > sm2)
                {
                    double t = sm1;
                    sm1 = sm2;
                    sm2 = t;
                }

                for (int i = dt1.Rows.Count - 1; i >= 0; --i)
                {
                    double sta1 = -1;
                    double sta2 = -1;

                    if (dt1.Rows[i][pipe_col_3d1] != DBNull.Value && dt1.Rows[i][pipe_col_3d2] != DBNull.Value)
                    {
                        sta1 = Convert.ToDouble(dt1.Rows[i][pipe_col_3d1]);
                        sta2 = Convert.ToDouble(dt1.Rows[i][pipe_col_3d2]);
                    }
                    else if (dt1.Rows[i][pipe_col_2d1] != DBNull.Value && dt1.Rows[i][pipe_col_2d2] != DBNull.Value)
                    {
                        sta1 = Convert.ToDouble(dt1.Rows[i][pipe_col_2d1]);
                        sta2 = Convert.ToDouble(dt1.Rows[i][pipe_col_2d2]);
                    }

                    if (sta1 >= sm1 && sta2 <= sm2)
                    {
                        dt1.Rows[i].Delete();
                    }

                    else if (sta2 > sm2 && sta1 >= sm1 && sta1 < sm2)
                    {
                        if (is3D == true)
                        {
                            dt1.Rows[i][pipe_col_3d1] = sm2;
                            dt1.Rows[i][pipe_col_len3d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                            dt1.Rows[i][col_xbeg] = pt2.X;
                            dt1.Rows[i][col_ybeg] = pt2.Y;

                        }
                        else
                        {
                            dt1.Rows[i][pipe_col_2d1] = sm2;
                            dt1.Rows[i][pipe_col_len2d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                            dt1.Rows[i][col_xbeg] = pt2.X;
                            dt1.Rows[i][col_ybeg] = pt2.Y;
                        }

                    }

                    else if (sta2 <= sm2 && sta1 < sm1 && sta2 > sm1)
                    {

                        if (is3D == true)
                        {
                            dt1.Rows[i][pipe_col_3d2] = sm1;
                            dt1.Rows[i][pipe_col_len3d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);
                            dt1.Rows[i][col_xend] = pt1.X;
                            dt1.Rows[i][col_yend] = pt1.Y;
                        }
                        else
                        {
                            dt1.Rows[i][pipe_col_2d2] = sm1;
                            dt1.Rows[i][pipe_col_len2d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);
                            dt1.Rows[i][col_xend] = pt1.X;
                            dt1.Rows[i][col_yend] = pt1.Y;

                        }
                    }
                    else if (sta2 > sm2 && sta1 < sm1 && sta2 > sm1)
                    {

                        System.Data.DataRow row1 = dt1.NewRow();
                        row1.ItemArray = dt1.Rows[i].ItemArray;


                        if (is3D == true)
                        {
                            row1[pipe_col_3d1] = sm2;
                            row1[pipe_col_len3d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                            row1[col_xbeg] = pt2.X;
                            row1[col_ybeg] = pt2.Y;
                        }
                        else
                        {
                            row1[pipe_col_2d1] = sm2;
                            row1[pipe_col_len2d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                            row1[col_xbeg] = pt2.X;
                            row1[col_ybeg] = pt2.Y;
                        }



                        if (is3D == true)
                        {
                            dt1.Rows[i][pipe_col_3d2] = sm1;
                            dt1.Rows[i][pipe_col_len3d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);
                            dt1.Rows[i][col_xend] = pt1.X;
                            dt1.Rows[i][col_yend] = pt1.Y;


                        }
                        else
                        {
                            dt1.Rows[i][pipe_col_2d2] = sm1;
                            dt1.Rows[i][pipe_col_len2d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);
                            dt1.Rows[i][col_xend] = pt1.X;
                            dt1.Rows[i][col_yend] = pt1.Y;

                        }

                        dt1.Rows.Add(row1);
                    }

                }

                dt1.Rows.Add();


                if (is3D == true)
                {
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_3d1] = sm1;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_3d2] = sm2;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_len3d] = Convert.ToDecimal(sm2) - Convert.ToDecimal(sm1);

                }
                else
                {
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_2d1] = sm1;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_2d2] = sm2;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_len2d] = Convert.ToDecimal(sm2) - Convert.ToDecimal(sm1);

                }

                dt1.Rows[dt1.Rows.Count - 1][col_item_no] = comboBox_pipe_mat.Text;
                dt1.Rows[dt1.Rows.Count - 1][col_xbeg] = pt1.X;
                dt1.Rows[dt1.Rows.Count - 1][col_ybeg] = pt1.Y;
                dt1.Rows[dt1.Rows.Count - 1][col_xend] = pt2.X;
                dt1.Rows[dt1.Rows.Count - 1][col_yend] = pt2.Y;

                if (is3D == true)
                {
                    dt1 = Functions.Sort_data_table(dt1, pipe_col_3d1);
                }
                else
                {
                    dt1 = Functions.Sort_data_table(dt1, pipe_col_2d1);

                }

                string prev_mat = Convert.ToString(dt1.Rows[dt1.Rows.Count - 1][col_item_no]);
                for (int i = dt1.Rows.Count - 2; i >= 0; --i)
                {
                    string mat1 = Convert.ToString(dt1.Rows[i][col_item_no]);

                    if (mat1 == prev_mat)
                    {
                        if (is3D == true)
                        {
                            dt1.Rows[i + 1][pipe_col_3d1] = dt1.Rows[i][pipe_col_3d1];
                        }
                        else
                        {
                            dt1.Rows[i + 1][pipe_col_2d1] = dt1.Rows[i][pipe_col_2d1];
                        }
                        dt1.Rows[i].Delete();

                    }
                    prev_mat = mat1;
                }
            }
        }

        private void populate_dt_extra(ref System.Data.DataTable dt1, double sm1, double sm2, string mat1, Point3d pt1, Point3d pt2)
        {
            bool is3D = ds_main.is3D;


            if (dt1 != null)
            {

                if (sm1 > sm2)
                {
                    double t = sm1;
                    sm1 = sm2;
                    sm2 = t;
                }

                string cat1 = "**";
                string type1 = "**";
                string descr1 = "**";

                for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                {
                    if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                    {


                        if (dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                        {
                            descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                        }

                        if (dt_mat_library.Rows[j][col_category] != DBNull.Value)
                        {
                            cat1 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                        }

                        if (dt_mat_library.Rows[j][col_type] != DBNull.Value)
                        {
                            type1 = Convert.ToString(dt_mat_library.Rows[j][col_type]);
                        }
                    }
                }


                for (int i = dt1.Rows.Count - 1; i >= 0; --i)
                {
                    double sta1 = -1;
                    double sta2 = -1;

                    if (dt1.Rows[i][col_item_no] != DBNull.Value && Convert.ToString(dt1.Rows[i][col_item_no]).ToUpper().Replace(" ", "") == mat1)
                    {
                        if (dt1.Rows[i][pipe_col_3d1] != DBNull.Value && dt1.Rows[i][pipe_col_3d2] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt1.Rows[i][pipe_col_3d1]);
                            sta2 = Convert.ToDouble(dt1.Rows[i][pipe_col_3d2]);
                        }
                        else if (dt1.Rows[i][pipe_col_2d1] != DBNull.Value && dt1.Rows[i][pipe_col_2d2] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt1.Rows[i][pipe_col_2d1]);
                            sta2 = Convert.ToDouble(dt1.Rows[i][pipe_col_2d2]);
                        }

                        if (sta1 >= sm1 && sta2 <= sm2)
                        {
                            dt1.Rows[i].Delete();
                        }

                        else if (sta2 > sm2 && sta1 >= sm1 && sta1 < sm2)
                        {
                            if (is3D == true)
                            {
                                dt1.Rows[i][pipe_col_3d1] = sm2;


                                dt1.Rows[i][pipe_col_len3d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                            }
                            else
                            {
                                dt1.Rows[i][pipe_col_2d1] = sm2;
                                dt1.Rows[i][pipe_col_len2d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);

                            }

                            dt1.Rows[i][col_xbeg] = pt2.X;
                            dt1.Rows[i][col_ybeg] = pt2.Y;

                        }

                        else if (sta2 <= sm2 && sta1 < sm1 && sta2 > sm1)
                        {

                            if (is3D == true)
                            {
                                dt1.Rows[i][pipe_col_3d2] = sm1;
                                dt1.Rows[i][pipe_col_len3d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);
                            }
                            else
                            {
                                dt1.Rows[i][pipe_col_2d2] = sm1;
                                dt1.Rows[i][pipe_col_len2d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);
                            }

                            dt1.Rows[i][col_xend] = pt1.X;
                            dt1.Rows[i][col_yend] = pt1.Y;
                        }
                        else if (sta2 > sm2 && sta1 < sm1 && sta2 > sm1)
                        {

                            System.Data.DataRow row1 = dt1.NewRow();
                            row1.ItemArray = dt1.Rows[i].ItemArray;


                            if (is3D == true)
                            {
                                row1[pipe_col_3d1] = sm2;
                                row1[pipe_col_len3d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                                row1[col_xbeg] = pt2.X;
                                row1[col_ybeg] = pt2.Y;

                            }
                            else
                            {
                                row1[pipe_col_2d1] = sm2;
                                row1[pipe_col_len2d] = Convert.ToDecimal(sta2) - Convert.ToDecimal(sm2);
                                row1[col_xbeg] = pt2.X;
                                row1[col_ybeg] = pt2.Y;
                            }



                            if (is3D == true)
                            {
                                dt1.Rows[i][pipe_col_3d2] = sm1;
                                dt1.Rows[i][pipe_col_len3d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);


                                dt1.Rows[i][col_xend] = pt1.X;
                                dt1.Rows[i][col_yend] = pt1.Y;

                            }
                            else
                            {
                                dt1.Rows[i][pipe_col_2d2] = sm1;
                                dt1.Rows[i][pipe_col_len2d] = Convert.ToDecimal(sm1) - Convert.ToDecimal(sta1);

                                dt1.Rows[i][col_xend] = pt1.X;
                                dt1.Rows[i][col_yend] = pt1.Y;

                            }




                            dt1.Rows.Add(row1);
                        }
                    }
                }

                dt1.Rows.Add();


                if (is3D == true)
                {
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_3d1] = sm1;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_3d2] = sm2;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_len3d] = Convert.ToDecimal(sm2) - Convert.ToDecimal(sm1);


                }
                else
                {
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_2d1] = sm1;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_2d2] = sm2;
                    dt1.Rows[dt1.Rows.Count - 1][pipe_col_len2d] = Convert.ToDecimal(sm2) - Convert.ToDecimal(sm1);

                }

                dt1.Rows[dt1.Rows.Count - 1][col_xbeg] = pt1.X;
                dt1.Rows[dt1.Rows.Count - 1][col_ybeg] = pt1.Y;
                dt1.Rows[dt1.Rows.Count - 1][col_xend] = pt2.X;
                dt1.Rows[dt1.Rows.Count - 1][col_yend] = pt2.Y;
                dt1.Rows[dt1.Rows.Count - 1][col_item_no] = mat1;
                dt1.Rows[dt1.Rows.Count - 1][col_altdesc] = descr1;
                dt1.Rows[dt1.Rows.Count - 1][pipe_col_mmid] = "**" + cat1 + "_" + type1;


                if (is3D == true)
                {
                    dt1 = Functions.Sort_data_table(dt1, pipe_col_3d1);
                }
                else
                {
                    dt1 = Functions.Sort_data_table(dt1, pipe_col_2d1);

                }

                string prev_mat = Convert.ToString(dt1.Rows[dt1.Rows.Count - 1][col_item_no]);
                for (int i = dt1.Rows.Count - 2; i >= 0; --i)
                {
                    string mat2 = Convert.ToString(dt1.Rows[i][col_item_no]);

                    if (mat2 == prev_mat)
                    {
                        if (is3D == true)
                        {
                            dt1.Rows[i + 1][pipe_col_3d1] = dt1.Rows[i][pipe_col_3d1];
                        }
                        else
                        {
                            dt1.Rows[i + 1][pipe_col_2d1] = dt1.Rows[i][pipe_col_2d1];
                        }
                        dt1.Rows[i].Delete();

                    }
                    prev_mat = mat2;
                }
            }
        }

        public void create_pipe_us_od_table(string table_name)
        {


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();



                        List1.Add(pipe_us_od_item_no);
                        List2.Add("Material");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(pipe_us_od_descr);
                        List2.Add(pipe_us_od_descr);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(pipe_us_od_cat);
                        List2.Add(pipe_us_od_cat);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(pipe_us_od_sta1);
                        List2.Add("Start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add(pipe_us_od_sta2);
                        List2.Add("End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);


                        Functions.Get_object_data_table(table_name, "Generated by MD", List1, List2, List3);


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        private void attach_od_to_us_pipes(System.Data.DataTable dt_od)
        {



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string pipe_type = null;
                        if (dt_od.Rows[i][pipe_us_od_item_no] != DBNull.Value)
                        {
                            pipe_type = Convert.ToString(dt_od.Rows[i][pipe_us_od_item_no]);
                        }
                        lista_val.Add(pipe_type);

                        string descr = null;
                        if (dt_od.Rows[i][pipe_us_od_descr] != DBNull.Value)
                        {
                            descr = Convert.ToString(dt_od.Rows[i][pipe_us_od_descr]);
                        }
                        lista_val.Add(descr);

                        string cat1 = null;
                        if (dt_od.Rows[i][pipe_us_od_cat] != DBNull.Value)
                        {
                            cat1 = Convert.ToString(dt_od.Rows[i][pipe_us_od_cat]);
                        }
                        lista_val.Add(cat1);


                        double sta1 = -1;
                        if (dt_od.Rows[i][pipe_us_od_sta1] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt_od.Rows[i][pipe_us_od_sta1]);
                        }
                        lista_val.Add(sta1);


                        double sta2 = -1;
                        if (dt_od.Rows[i][pipe_us_od_sta2] != DBNull.Value)
                        {
                            sta2 = Convert.ToDouble(dt_od.Rows[i][pipe_us_od_sta2]);
                        }
                        lista_val.Add(sta2);


                        Polyline atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, pipes_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        public void create_point_od_table(string table_name)
        {


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();



                        List1.Add(col_item_no);
                        List2.Add(col_item_no);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_sta);
                        List2.Add(col_sta);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add(col_descr);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(table_name, "Generated by MD", List1, List2, List3);


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        private void attach_od_to_points(System.Data.DataTable dt_od)
        {



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string mat1 = null;
                        if (dt_od.Rows[i][col_item_no] != DBNull.Value)
                        {
                            mat1 = Convert.ToString(dt_od.Rows[i][col_item_no]);
                        }
                        lista_val.Add(mat1);


                        double sta1 = -1;
                        if (dt_od.Rows[i][col_sta] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt_od.Rows[i][col_sta]);
                        }
                        lista_val.Add(sta1);

                        string descr = null;
                        if (dt_od.Rows[i][col_descr] != DBNull.Value)
                        {
                            descr = Convert.ToString(dt_od.Rows[i][col_descr]);
                        }
                        lista_val.Add(descr);








                        BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForWrite) as BlockReference;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, points_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        public void populate_datagridview_pipe()
        {
            bool is3D = ds_main.is3D;


            if (ds_main.dt_pipe != null && ds_main.dt_pipe.Rows.Count > 0)
            {
                DataGridViewTextBoxColumn DG_Col_mat = datagrid_to_datatable_textbox(ds_main.dt_pipe, col_item_no);

                DataGridViewTextBoxColumn DG_Col_sta1 = datagrid_to_datatable_textbox(ds_main.dt_pipe, pipe_col_2d1);
                DataGridViewTextBoxColumn DG_Col_sta2 = datagrid_to_datatable_textbox(ds_main.dt_pipe, pipe_col_2d2);
                DataGridViewTextBoxColumn DG_Col_len = datagrid_to_datatable_textbox(ds_main.dt_pipe, pipe_col_len2d);
                if (is3D == true)
                {

                    DG_Col_sta1 = datagrid_to_datatable_textbox(ds_main.dt_pipe, pipe_col_3d1);
                    DG_Col_sta2 = datagrid_to_datatable_textbox(ds_main.dt_pipe, pipe_col_3d2);
                    DG_Col_len = datagrid_to_datatable_textbox(ds_main.dt_pipe, pipe_col_len3d);
                }



                if (ft_pipe == 0)
                {
                    dataGridView_pipe.Columns.AddRange(DG_Col_mat, DG_Col_sta1, DG_Col_sta2, DG_Col_len);
                    dataGridView_pipe.Columns[0].Name = col_item_no;

                    if (is3D == true)
                    {

                        dataGridView_pipe.Columns[1].Name = pipe_col_3d1;
                        dataGridView_pipe.Columns[2].Name = pipe_col_3d2;
                        dataGridView_pipe.Columns[3].Name = pipe_col_len3d;
                    }
                    else
                    {
                        dataGridView_pipe.Columns[1].Name = pipe_col_2d1;
                        dataGridView_pipe.Columns[2].Name = pipe_col_2d2;
                        dataGridView_pipe.Columns[3].Name = pipe_col_len2d;
                    }


                    ft_pipe = 1;
                }

                dataGridView_pipe.AutoGenerateColumns = false;
                dataGridView_pipe.DataSource = ds_main.dt_pipe;
                dataGridView_pipe.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                //DG_Col_OD_Table.FlatStyle = FlatStyle.Flat;
                //DG_Col_OD_Table.Width = 100;
                dataGridView_pipe.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_pipe.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_pipe.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Padding newpadding = new Padding(4, 0, 0, 0);
                dataGridView_pipe.ColumnHeadersDefaultCellStyle.Padding = newpadding;
                dataGridView_pipe.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_pipe.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
                dataGridView_pipe.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_pipe.EnableHeadersVisualStyles = false;

                dataGridView_pipe.Columns[0].Width = 75;
                dataGridView_pipe.Columns[1].Width = 150;
                dataGridView_pipe.Columns[2].Width = 150;
                dataGridView_pipe.Columns[3].Width = 150;
            }
            else
            {
                dataGridView_pipe.DataSource = null;
            }



        }

        private void button_zoom_to_pipe_Click(object sender, EventArgs e)
        {

            try
            {
                if (ds_main.dt_pipe != null && ds_main.dt_pipe.Rows.Count > 0)
                {
                    set_enable_false();
                    int row_idx = dataGridView_pipe.SelectedCells[0].RowIndex;
                    if (row_idx >= 0)
                    {

                        bool is3D = ds_main.is3D;
                        string handle1 = null;
                        double sta1 = -1;
                        double sta2 = -1;

                        if (is3D == true)
                        {
                            sta1 = Convert.ToDouble(dataGridView_pipe.Rows[row_idx].Cells[pipe_col_3d1].Value);
                            sta2 = Convert.ToDouble(dataGridView_pipe.Rows[row_idx].Cells[pipe_col_3d2].Value);
                        }
                        else
                        {
                            sta1 = Convert.ToDouble(dataGridView_pipe.Rows[row_idx].Cells[pipe_col_2d1].Value);
                            sta2 = Convert.ToDouble(dataGridView_pipe.Rows[row_idx].Cells[pipe_col_2d2].Value);
                        }

                        for (int i = 0; i < ds_main.dt_pipe.Rows.Count; ++i)
                        {

                            double sta11 = -1;
                            double sta22 = -1;

                            if (is3D == true)
                            {
                                sta11 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_3d1]);
                                sta22 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_3d2]);
                            }
                            else
                            {
                                sta11 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_2d1]);
                                sta22 = Convert.ToDouble(ds_main.dt_pipe.Rows[i][pipe_col_2d2]);
                            }
                            if (sta11 == sta1 && sta22 == sta2)
                            {

                                if (ds_main.dt_pipe.Rows[i][pipe_col_mmid] != DBNull.Value)
                                {
                                    handle1 = Convert.ToString(ds_main.dt_pipe.Rows[i][pipe_col_mmid]);
                                }



                            }

                        }


                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    Functions.zoom_to_object(id1);
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_select_material_pipe_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (dataGridView_pipe.Rows.Count > 0)
            {
                set_enable_false();

                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            bool ask_for_selection = false;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_object = (Autodesk.AutoCAD.EditorInput.PromptSelectionResult)Editor1.SelectImplied();

                            if (Rezultat_object.Status == PromptStatus.OK)
                            {
                                if (Rezultat_object.Value.Count == 0)
                                {
                                    ask_for_selection = true;
                                }
                                if (Rezultat_object.Value.Count > 1)
                                {
                                    MessageBox.Show("There is more than one object selected," + "\r\n" + "the first object in selection will be the one that will be current in table");
                                    ask_for_selection = false;
                                }
                            }
                            else ask_for_selection = true;



                            if (ask_for_selection == true)
                            {
                                this.MdiParent.WindowState = FormWindowState.Minimized;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_object = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_object.MessageForAdding = "\nSelect a pipe";
                                Prompt_object.SingleOnly = true;
                                Rezultat_object = Editor1.GetSelection(Prompt_object);

                            }


                            if (Rezultat_object.Status != PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                set_enable_true();
                                return;
                            }
                            this.MdiParent.WindowState = FormWindowState.Normal;


                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object.Value[0].ObjectId, OpenMode.ForRead);
                            string handle1 = Ent1.ObjectId.Handle.Value.ToString();
                            string handle2 = null;
                            bool is3D = ds_main.is3D;


                            for (int i = 0; i < dataGridView_pipe.Rows.Count; ++i)
                            {

                                double sta1 = -1;
                                double sta2 = -1;

                                if (is3D == true)
                                {
                                    sta1 = Convert.ToDouble(dataGridView_pipe.Rows[i].Cells[pipe_col_3d1].Value);
                                    sta2 = Convert.ToDouble(dataGridView_pipe.Rows[i].Cells[pipe_col_3d2].Value);
                                }
                                else
                                {
                                    sta1 = Convert.ToDouble(dataGridView_pipe.Rows[i].Cells[pipe_col_2d1].Value);
                                    sta2 = Convert.ToDouble(dataGridView_pipe.Rows[i].Cells[pipe_col_2d2].Value);
                                }


                                for (int j = 0; j < ds_main.dt_pipe.Rows.Count; ++j)
                                {

                                    double sta11 = -1;
                                    double sta22 = -1;

                                    if (is3D == true)
                                    {
                                        sta11 = Convert.ToDouble(ds_main.dt_pipe.Rows[j][pipe_col_3d1]);
                                        sta22 = Convert.ToDouble(ds_main.dt_pipe.Rows[j][pipe_col_3d2]);
                                    }
                                    else
                                    {
                                        sta11 = Convert.ToDouble(ds_main.dt_pipe.Rows[j][pipe_col_2d1]);
                                        sta22 = Convert.ToDouble(ds_main.dt_pipe.Rows[j][pipe_col_2d2]);
                                    }
                                    if (sta11 == sta1 && sta22 == sta2)
                                    {

                                        if (ds_main.dt_pipe.Rows[j][pipe_col_mmid] != DBNull.Value)
                                        {
                                            handle2 = Convert.ToString(ds_main.dt_pipe.Rows[j][pipe_col_mmid]);
                                        }

                                        if (handle1 == handle2)
                                        {
                                            dataGridView_pipe.CurrentCell = dataGridView_pipe.Rows[i].Cells[0];
                                            i = dataGridView_pipe.Rows.Count;
                                            j = ds_main.dt_pipe.Rows.Count;
                                        }
                                    }
                                }
                            }










                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            set_enable_true();
        }

        public void populate_textbox_cl(string continut)
        {
            if (continut != "")
            {
                textBox_cl.Text = continut;
                textBox_cl.ForeColor = Color.LightGreen;
                textBox_cl.Font = font8;
                textBox_cl.TextAlign = HorizontalAlignment.Right;
            }
            else
            {
                textBox_cl.Text = "No centerline loaded";
                textBox_cl.ForeColor = Color.Red;
                textBox_cl.Font = font10;
                textBox_cl.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void button_place_point_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            System.Data.DataTable dt_cl = ds_main.dt_centerline;

            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    System.Data.DataTable dt_od_pt = new System.Data.DataTable();
                    dt_od_pt.Columns.Add(col_item_no, typeof(string));
                    dt_od_pt.Columns.Add(col_sta, typeof(string));
                    dt_od_pt.Columns.Add(col_descr, typeof(string));
                    dt_od_pt.Columns.Add("id", typeof(ObjectId));

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        string blockname = "";
                        string layer1 = "0";
                        string descr1 = "**";


                        System.Windows.Forms.Button bt1 = sender as System.Windows.Forms.Button;
                        string ct1 = bt1.Name.Replace("button_draw_", "");
                        string mat1 = "123424";
                        foreach (TabPage tab1 in flatTabControl1.TabPages)
                        {
                            foreach (Panel panel1 in tab1.Controls)
                            {
                                foreach (Control ctrl1 in panel1.Controls)
                                {
                                    ComboBox combo1 = ctrl1 as ComboBox;
                                    if (combo1 != null && combo1.Name.Replace("combo_", "") == ct1)
                                    {
                                        mat1 = combo1.Text;
                                    }
                                }
                            }
                        }

                        create_point_od_table(points_od);

                        for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                        {
                            if (dt_mat_library.Rows[i][col_item_no] != DBNull.Value)
                            {
                                string mat2 = Convert.ToString(dt_mat_library.Rows[i][col_item_no]);

                                if (mat1 == mat2 && dt_mat_library.Rows[i][col_MSblock] != DBNull.Value)
                                {
                                    blockname = Convert.ToString(dt_mat_library.Rows[i][col_MSblock]);
                                    if (dt_mat_library.Rows[i][col_layer] != DBNull.Value)
                                    {
                                        layer1 = Convert.ToString(dt_mat_library.Rows[i][col_layer]);
                                        Functions.Creaza_layer(layer1, 1, true);
                                    }

                                    if (dt_mat_library.Rows[i][col_descr] != DBNull.Value)
                                    {
                                        descr1 = Convert.ToString(dt_mat_library.Rows[i][col_descr]);
                                    }

                                    i = dt_mat_library.Rows.Count;
                                }
                            }
                        }




                        if (blockname != "")
                        {
                            if (BlockTable1.Has(blockname) == true)
                            {
                                BlockTableRecord btr = Trans1.GetObject(BlockTable1[blockname], OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null)
                                {
                                    Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);

                                    //BlockReference br = new BlockReference(new Point3d(), btr.ObjectId);
                                    //br.Layer = layer1;
                                    //br.ColorIndex = 256;

                                    //BTrecord.AppendEntity(br);
                                    //Trans1.AddNewlyCreatedDBObject(br, true);

                                    //Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();

                                    //BlockJig1 myJig = new BlockJig1(Trans1, br, attInfo, Poly2D);
                                    //if (myJig.Run() != PromptStatus.OK)
                                    //{
                                    //    Editor1.SetImpliedSelection(Empty_array);
                                    //    Editor1.WriteMessage("\nCommand:");
                                    //    set_enable_true();
                                    //    return;
                                    //}

                                    BlockReference br1 = null;
                                    jig_actions.insert_block(ref br1, blockname, layer1, Poly2D);

                                    Point3d pt_ins = br1.Position;

                                    Point3d pt_on_poly = Poly2D.GetClosestPointTo(new Point3d(pt_ins.X, pt_ins.Y, 0), Vector3d.ZAxis, false);
                                    double sta1 = Poly2D.GetDistAtPoint(pt_on_poly);

                                    dt_od_pt.Rows.Add();
                                    dt_od_pt.Rows[dt_od_pt.Rows.Count - 1]["id"] = br1.ObjectId;
                                    dt_od_pt.Rows[dt_od_pt.Rows.Count - 1][col_item_no] = mat1;
                                    dt_od_pt.Rows[dt_od_pt.Rows.Count - 1][col_sta] = sta1;
                                    dt_od_pt.Rows[dt_od_pt.Rows.Count - 1][col_descr] = descr1;


                                    if (ct_list.Contains(ct1) == true)
                                    {
                                        System.Data.DataTable dt1 = dt_ct[ct_list.IndexOf(ct1)];

                                        if (dt1 != null)
                                        {
                                            ds_main.dt_points.Rows.Add();
                                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_mmid] = "**" + ct1;
                                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_altdesc] = descr1;
                                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_item_no] = mat1;
                                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_2dsta] = sta1;
                                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_x] = pt_on_poly.X;
                                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_y] = pt_on_poly.Y;

                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][col_mmid] = "**" + ct1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_altdesc] = descr1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_item_no] = mat1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_2dsta] = sta1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = pt_on_poly.X;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = pt_on_poly.Y;

                                            dt1 = Functions.Sort_data_table(dt1, col_2dsta);
                                            ds_main.dt_points = Functions.Sort_data_table(ds_main.dt_points, col_2dsta);

                                        }
                                    }
                                    Trans1.Commit();
                                }
                            }
                        }
                    }
                    attach_od_to_points(dt_od_pt);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }




        private void button_save_point_Click(object sender, EventArgs e)
        {
            // here is saving all dt_points the ct1 is just created in case you want to save only 1 tab.
            set_enable_false();
            get_client_project_segment_pipe_diam();
            System.Windows.Forms.Button bt1 = sender as System.Windows.Forms.Button;
            string ct1 = bt1.Name.ToLower().Replace("button_save_", "").Replace(" ", "").ToUpper();

            if (ct_list.Contains(ct1) == true)
            {
                if (ct1 == "ELL_POINT")
                {

                }

                System.Data.DataTable dt1 = dt_ct[ct_list.IndexOf(ct1)];
                List<string> lista1 = new List<string>();

                for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                {
                    if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value && dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                    {
                        if ((Convert.ToString(dt_mat_library.Rows[j][col_category]) + "_" + Convert.ToString(dt_mat_library.Rows[j][col_type])).ToUpper().Replace(" ", "") == ct1)
                        {
                            string mat1 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper().Replace(" ", "");
                            lista1.Add(mat1);
                        }
                    }
                }

                for (int i = ds_main.dt_points.Rows.Count - 1; i >= 0; --i)
                {
                    if (ds_main.dt_points.Rows[i][col_item_no] != DBNull.Value)
                    {
                        string mat1 = Convert.ToString(ds_main.dt_points.Rows[i][col_item_no]).ToUpper().Replace(" ", "");
                        if (lista1.Contains(mat1) == true)
                        {
                            ds_main.dt_points.Rows[i].Delete();
                        }
                    }
                }

                for (int j = 0; j < dt1.Rows.Count; ++j)
                {
                    ds_main.dt_points.Rows.Add();
                    for (int k = 0; k < dt1.Columns.Count; ++k)
                    {
                        ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][k] = dt1.Rows[j][k];
                    }
                }

                if (ds_main.dt_points.Rows.Count > 0)
                {
                    string col1 = col_2dsta;
                    if (ds_main.dt_points.Rows[0][col_3dsta] != DBNull.Value) col1 = col_3dsta;

                    ds_main.dt_points = Functions.Sort_data_table(ds_main.dt_points, col1);

                    Transfer_dt_points_to_file1();
                }

            }

            set_enable_true();
        }
        public void Transfer_dt_points_to_file1()
        {


            string tabname = "MatPoints";

            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

            bool is_opened = false;


            try
            {
                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {


                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            if (Workbook2.FullName == ds_main.config_xls)
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                                {

                                    Workbook1 = Workbook2;
                                    if (Wx.Name == tabname)
                                    {
                                        W1 = Wx;
                                    }
                                    is_opened = true;
                                }
                            }
                            if (W1 == null)
                            {
                                W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                W1.Name = tabname;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }


                    if (is_opened == false)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                        Workbook1 = Excel1.Workbooks.Open(ds_main.config_xls);

                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name == tabname)
                            {
                                W1 = Wx;
                            }
                        }

                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W1.Name = tabname;
                        }
                    }



                }


                System.Data.DataTable dt_pts = ds_main.dt_points;



                if (dt_pts != null && W1 != null)
                {
                    if (dt_pts.Rows.Count > 0)
                    {
                        if (ds_main.dt_centerline != null && ds_main.dt_centerline.Rows.Count > 1)
                        {
                            Polyline Poly2D = Functions.Build_2d_poly_for_scanning(ds_main.dt_centerline);
                            for (int i = 0; i < dt_pts.Rows.Count; ++i)
                            {
                                if (dt_pts.Rows[i][col_2dsta] != DBNull.Value && (dt_pts.Rows[i][col_x] == DBNull.Value || dt_pts.Rows[i][col_y] == DBNull.Value))
                                {
                                    double sta1 = Convert.ToDouble(dt_pts.Rows[i][col_2dsta]);
                                    if (sta1 < 0) sta1 = 0;
                                    if (sta1 >= Poly2D.Length) sta1 = Poly2D.Length - 0.001;

                                    Point3d pt_on_poly = Poly2D.GetPointAtDist(sta1);
                                    dt_pts.Rows[i][col_x] = pt_on_poly.X;
                                    dt_pts.Rows[i][col_y] = pt_on_poly.Y;
                                }
                            }
                        }

                        Create_header_material_points_file(W1, ds_main.client1, ds_main.project1, ds_main.segment1, dt_pts);
                        int last_row = nr_max + 13;
                        W1.Cells.NumberFormat = "General";
                        int maxRows = dt_pts.Rows.Count;
                        int maxCols = dt_pts.Columns.Count;
                        W1.Range["A13:Q" + last_row.ToString()].ClearContents();
                        W1.Range["A13:Q" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A13:Q" + (13 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (dt_pts.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = dt_pts.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;
                    }
                    else
                    {
                        W1.Cells.Clear();
                        W1.Cells.ClearFormats();
                    }



                }
                else if (W1 != null)
                {
                    W1.Cells.Clear();
                    W1.Cells.ClearFormats();
                }

                if (is_opened == false)
                {
                    Workbook1.Save();
                    Workbook1.Close();
                }

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }




        private void button_pt_zoom_Click(object sender, EventArgs e)
        {
            try
            {
                if (ds_main.dt_points != null && ds_main.dt_points.Rows.Count > 0)
                {


                    System.Windows.Forms.Button bt1 = sender as System.Windows.Forms.Button;
                    string ct1 = bt1.Name.Replace("btn_zoom_to_", "");

                    DataGridView dgv1 = null;


                    foreach (TabPage tab1 in flatTabControl1.TabPages)
                    {
                        foreach (Panel panel1 in tab1.Controls)
                        {
                            foreach (Control ctrl1 in panel1.Controls)
                            {
                                DataGridView dgv2 = ctrl1 as DataGridView;
                                if (dgv2 != null && dgv2.Name.Replace("dgv_", "") == ct1)
                                {
                                    dgv1 = dgv2;
                                }
                            }
                        }
                    }

                    set_enable_false();

                    if (dgv1.SelectedCells[0].RowIndex > -1)
                    {

                        double sta1 = -1;




                        if (dgv1.Rows[dgv1.SelectedCells[0].RowIndex].Cells[1].Value != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dgv1.Rows[dgv1.SelectedCells[0].RowIndex].Cells[1].Value);
                        }


                        for (int i = 0; i < ds_main.dt_points.Rows.Count; ++i)
                        {

                            double sta2 = -1;


                            if (ds_main.dt_points.Rows[i][col_2dsta] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_2dsta]);
                            }

                            if (ds_main.dt_points.Rows[i][col_3dsta] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_3dsta]);
                            }

                            if (Math.Round(sta1, 2) == Math.Round(sta2, 2))
                            {

                                if (ds_main.dt_points.Rows[i][col_x] != DBNull.Value && ds_main.dt_points.Rows[i][col_y] != DBNull.Value)
                                {
                                    double x = Convert.ToDouble(ds_main.dt_points.Rows[i][col_x]);
                                    double y = Convert.ToDouble(ds_main.dt_points.Rows[i][col_y]);

                                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                    {
                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                        {

                                            Functions.zoom_to_Point(new Point3d(x, y, 0), 0.3);

                                        }
                                    }
                                }

                                i = ds_main.dt_points.Rows.Count;

                            }

                        }



                    }
                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_pt_select_Click(object sender, EventArgs e)
        {

            Button buton1 = sender as Button;
            string nume1 = buton1.Name.Replace("btn_select_", "").ToUpper();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (dataGridView_pipe.Rows.Count > 0)
            {
                set_enable_false();

                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);





                            bool ask_for_selection = false;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_object = (Autodesk.AutoCAD.EditorInput.PromptSelectionResult)Editor1.SelectImplied();

                            if (Rezultat_object.Status == PromptStatus.OK)
                            {
                                if (Rezultat_object.Value.Count == 0)
                                {
                                    ask_for_selection = true;
                                }
                                if (Rezultat_object.Value.Count > 1)
                                {
                                    MessageBox.Show("There is more than one object selected," + "\r\n" + "the first object in selection will be the one that will be current in table");
                                    ask_for_selection = false;
                                }
                            }
                            else ask_for_selection = true;



                            if (ask_for_selection == true)
                            {
                                this.MdiParent.WindowState = FormWindowState.Minimized;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_object = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_object.MessageForAdding = "\nSelect " + nume1;
                                Prompt_object.SingleOnly = true;
                                Rezultat_object = Editor1.GetSelection(Prompt_object);
                            }


                            if (Rezultat_object.Status != PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                set_enable_true();
                                return;
                            }
                            this.MdiParent.WindowState = FormWindowState.Normal;

                            double sta1 = -1;

                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object.Value[0].ObjectId, OpenMode.ForRead);
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                            {
                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];

                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;

                                    for (int i = 0; i < Record1.Count; ++i)
                                    {

                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                        string Nume_field = Field_def1.Name;
                                        string Valoare1 = Record1[i].StrValue;
                                        if (Nume_field == "STA")
                                        {
                                            if (Functions.IsNumeric(Valoare1.Replace("+", "")) == true)
                                            {
                                                sta1 = Convert.ToDouble(Valoare1.Replace("+", ""));
                                            }
                                        }
                                    }
                                }
                            }

                            DataGridView dgv1 = null;

                            foreach (TabPage tab1 in flatTabControl1.TabPages)
                            {
                                foreach (Panel panel1 in tab1.Controls)
                                {
                                    foreach (Control ctrl1 in panel1.Controls)
                                    {
                                        DataGridView dgv2 = ctrl1 as DataGridView;
                                        if (dgv2 != null && dgv2.Name.Replace("dgv_", "") == nume1)
                                        {
                                            dgv1 = dgv2;
                                        }
                                    }
                                }
                            }

                            if (dgv1 != null)
                            {
                                for (int i = 0; i < dgv1.Rows.Count; ++i)
                                {
                                    double sta2 = Convert.ToDouble(dgv1.Rows[i].Cells[1].Value);
                                    if (Math.Round(sta1, 2) == Math.Round(sta2, 2))
                                    {
                                        dgv1.CurrentCell = dgv1.Rows[i].Cells[1];
                                        i = dgv1.Rows.Count;
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            set_enable_true();
        }



        private void button_lin_save_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                get_client_project_segment_pipe_diam();
                transfer_dt_extra_to_file1();
                set_enable_true();
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button_lin_zoom_Click(object sender, EventArgs e)
        {
            try
            {
                if (ds_main.dt_extra != null && ds_main.dt_extra.Rows.Count > 0)
                {


                    System.Windows.Forms.Button bt1 = sender as System.Windows.Forms.Button;
                    string ct1 = bt1.Name.Replace("btn_zoom_to_", "");

                    DataGridView dgv1 = null;


                    foreach (TabPage tab1 in flatTabControl1.TabPages)
                    {
                        foreach (Panel panel1 in tab1.Controls)
                        {
                            foreach (Control ctrl1 in panel1.Controls)
                            {
                                DataGridView dgv2 = ctrl1 as DataGridView;
                                if (dgv2 != null && dgv2.Name.Replace("dgv_", "") == ct1)
                                {
                                    dgv1 = dgv2;
                                }
                            }
                        }
                    }


                    set_enable_false();

                    if (dgv1.SelectedCells[0].RowIndex > -1)
                    {
                        double sta1 = -1;
                        double sta2 = -1;

                        if (dgv1.Rows[dgv1.SelectedCells[0].RowIndex].Cells[1].Value != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dgv1.Rows[dgv1.SelectedCells[0].RowIndex].Cells[1].Value);
                        }

                        if (dgv1.Rows[dgv1.SelectedCells[0].RowIndex].Cells[2].Value != DBNull.Value)
                        {
                            sta2 = Convert.ToDouble(dgv1.Rows[dgv1.SelectedCells[0].RowIndex].Cells[2].Value);
                        }

                        for (int i = 0; i < ds_main.dt_extra.Rows.Count; ++i)
                        {

                            double sta3 = -1;
                            double sta4 = -1;

                            if (ds_main.dt_extra.Rows[i][col_2dbeg] != DBNull.Value)
                            {
                                sta3 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_2dbeg]);
                            }

                            if (ds_main.dt_extra.Rows[i][col_3dbeg] != DBNull.Value)
                            {
                                sta3 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_3dbeg]);
                            }

                            if (ds_main.dt_extra.Rows[i][col_2dend] != DBNull.Value)
                            {
                                sta4 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_2dend]);
                            }

                            if (ds_main.dt_extra.Rows[i][col_3dend] != DBNull.Value)
                            {
                                sta4 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_3dend]);
                            }

                            if (Math.Round(sta1, 2) == Math.Round(sta3, 2) && Math.Round(sta2, 2) == Math.Round(sta4, 2))
                            {
                                if (ds_main.dt_extra.Rows[i][col_xbeg] != DBNull.Value && ds_main.dt_extra.Rows[i][col_xend] != DBNull.Value &&
                                    ds_main.dt_extra.Rows[i][col_ybeg] != DBNull.Value && ds_main.dt_extra.Rows[i][col_yend] != DBNull.Value)
                                {
                                    double x1 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_xbeg]);
                                    double y1 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_ybeg]);

                                    double x2 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_xend]);
                                    double y2 = Convert.ToDouble(ds_main.dt_extra.Rows[i][col_yend]);

                                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                    {
                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                        {
                                            Functions.zoom_to_Point(new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0), 0.1);
                                        }
                                    }
                                }
                                i = ds_main.dt_extra.Rows.Count;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_lin_select_Click(object sender, EventArgs e)
        {

            Button buton1 = sender as Button;
            string nume1 = buton1.Name.Replace("btn_select_", "").ToUpper();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (dataGridView_pipe.Rows.Count > 0)
            {
                set_enable_false();

                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);





                            bool ask_for_selection = false;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_object = (Autodesk.AutoCAD.EditorInput.PromptSelectionResult)Editor1.SelectImplied();

                            if (Rezultat_object.Status == PromptStatus.OK)
                            {
                                if (Rezultat_object.Value.Count == 0)
                                {
                                    ask_for_selection = true;
                                }
                                if (Rezultat_object.Value.Count > 1)
                                {
                                    MessageBox.Show("There is more than one object selected," + "\r\n" + "the first object in selection will be the one that will be current in table");
                                    ask_for_selection = false;
                                }
                            }
                            else ask_for_selection = true;



                            if (ask_for_selection == true)
                            {
                                this.MdiParent.WindowState = FormWindowState.Minimized;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_object = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_object.MessageForAdding = "\nSelect " + nume1;
                                Prompt_object.SingleOnly = true;
                                Rezultat_object = Editor1.GetSelection(Prompt_object);

                            }


                            if (Rezultat_object.Status != PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                set_enable_true();
                                return;
                            }
                            this.MdiParent.WindowState = FormWindowState.Normal;

                            double sta1 = -1;
                            double sta2 = -1;

                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object.Value[0].ObjectId, OpenMode.ForRead);
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                            {
                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];

                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;

                                    for (int i = 0; i < Record1.Count; ++i)
                                    {
                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                        string Nume_field = Field_def1.Name;
                                        string Valoare1 = Record1[i].StrValue;
                                        if (Nume_field == "BeginSta")
                                        {
                                            if (Functions.IsNumeric(Valoare1.Replace("+", "")) == true)
                                            {
                                                sta1 = Convert.ToDouble(Valoare1.Replace("+", ""));
                                            }
                                        }
                                        if (Nume_field == "EndSta")
                                        {
                                            if (Functions.IsNumeric(Valoare1.Replace("+", "")) == true)
                                            {
                                                sta2 = Convert.ToDouble(Valoare1.Replace("+", ""));
                                            }
                                        }
                                    }
                                }
                            }

                            DataGridView dgv1 = null;

                            foreach (TabPage tab1 in flatTabControl1.TabPages)
                            {
                                foreach (Panel panel1 in tab1.Controls)
                                {
                                    foreach (Control ctrl1 in panel1.Controls)
                                    {
                                        DataGridView dgv2 = ctrl1 as DataGridView;
                                        if (dgv2 != null && dgv2.Name.Replace("dgv_", "") == nume1)
                                        {
                                            dgv1 = dgv2;
                                        }
                                    }
                                }
                            }

                            if (dgv1 != null)
                            {
                                for (int i = 0; i < dgv1.Rows.Count; ++i)
                                {
                                    double sta3 = Convert.ToDouble(dgv1.Rows[i].Cells[1].Value);
                                    double sta4 = Convert.ToDouble(dgv1.Rows[i].Cells[2].Value);
                                    if (Math.Round(sta1, 2) == Math.Round(sta3, 2) && Math.Round(sta2, 2) == Math.Round(sta4, 2))
                                    {
                                        dgv1.CurrentCell = dgv1.Rows[i].Cells[1];
                                        i = dgv1.Rows.Count;
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            set_enable_true();
        }

        private void button_pt_refresh_Click(object sender, EventArgs e)
        {
            Button bt1 = sender as Button;

            if (ct_list != null && ct_list.Count > 0)
            {
                for (int i = 0; i < ct_list.Count; ++i)
                {
                    if (bt1 != null && bt1.Name.ToUpper() == "BUTTON_REFRESH_" + ct_list[i])
                    {
                        draw_points(ct_list[i]);
                    }
                }
            }

        }

        public void draw_points(string ct1)
        {
            System.Data.DataTable dt_cl = ds_main.dt_centerline;

            if (dt_cl == null || dt_cl.Rows.Count < 2)
            {
                MessageBox.Show("No centerline loaded\r\nOperation aborted.", "Points Materials", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                System.Data.DataTable dt_od_pt = new System.Data.DataTable();
                dt_od_pt.Columns.Add(col_item_no, typeof(string));
                dt_od_pt.Columns.Add(col_sta, typeof(string));
                dt_od_pt.Columns.Add(col_descr, typeof(string));
                dt_od_pt.Columns.Add("id", typeof(ObjectId));

                List<string> lista1 = Functions.get_blocks_from_current_drawing();

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    BlockTableRecord btr = null;

                    if (ds_main.dt_points != null && ds_main.dt_points.Rows.Count > 0)
                    {
                        if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
                        {

                            Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);

                            for (int i = 0; i < ds_main.dt_points.Rows.Count; ++i)
                            {
                                if (ds_main.dt_points.Rows[i][col_item_no] != DBNull.Value)
                                {
                                    string mat1 = Convert.ToString(ds_main.dt_points.Rows[i][col_item_no]);
                                    string ct2 = "**";

                                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                    {
                                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                                        {
                                            if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value)
                                            {
                                                ct2 = (Convert.ToString(dt_mat_library.Rows[j][col_category]) + "_" + Convert.ToString(dt_mat_library.Rows[j][col_type])).ToUpper();
                                            }
                                        }
                                    }


                                    if (ct2 == ct1)
                                    {
                                        for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                        {
                                            if (dt_mat_library.Rows[j][col_MSblock] != DBNull.Value && dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                                            {
                                                string block_name = Convert.ToString(dt_mat_library.Rows[j][col_MSblock]);

                                                if (lista1.Contains(block_name) == false)
                                                {
                                                    MessageBox.Show("the block " + block_name + " not present in the current drawing", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                }
                                                else
                                                {
                                                    string layer1 = "0";
                                                    if (dt_mat_library.Rows[j][col_layer] != DBNull.Value)
                                                    {
                                                        layer1 = Convert.ToString(dt_mat_library.Rows[j][col_layer]);
                                                    }

                                                    if (BlockTable1.Has(block_name) == true)
                                                    {
                                                        btr = Trans1.GetObject(BlockTable1[block_name], OpenMode.ForRead) as BlockTableRecord;
                                                    }

                                                    if (btr != null)
                                                    {

                                                        foreach (ObjectId id1 in BTrecord)
                                                        {
                                                            BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                                            if (block1 != null)
                                                            {
                                                                string nume1 = Functions.get_block_name(block1);

                                                                if (block1.Layer == layer1 && nume1 == block_name)
                                                                {
                                                                    block1.UpgradeOpen();
                                                                    block1.Erase();
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }



                                }

                            }

                            create_point_od_table(points_od);




                            for (int i = 0; i < ds_main.dt_points.Rows.Count; ++i)
                            {

                                if (ds_main.dt_points.Rows[i][col_item_no] != DBNull.Value)
                                {
                                    string mat1 = Convert.ToString(ds_main.dt_points.Rows[i][col_item_no]);
                                    string ct2 = "**";
                                    string descr1 = "**";

                                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                    {
                                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                                        {
                                            string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]);

                                            if (mat1 == mat2)
                                            {
                                                if (dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                                                {
                                                    descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                                                }

                                                j = dt_mat_library.Rows.Count;
                                            }
                                        }
                                    }


                                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                    {
                                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                                        {
                                            if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value)
                                            {
                                                ct2 = (Convert.ToString(dt_mat_library.Rows[j][col_category]) + "_" + Convert.ToString(dt_mat_library.Rows[j][col_type])).ToUpper();

                                                if (ct2 == ct1)
                                                {
                                                    string block_name = Convert.ToString(dt_mat_library.Rows[j][col_MSblock]);

                                                    if (lista1.Contains(block_name) == true)
                                                    {
                                                        string layer1 = "0";
                                                        if (dt_mat_library.Rows[j][col_layer] != DBNull.Value)
                                                        {
                                                            layer1 = Convert.ToString(dt_mat_library.Rows[j][col_layer]);
                                                            Functions.Creaza_layer(layer1, 1, true);
                                                        }

                                                        if (BlockTable1.Has(block_name) == true)
                                                        {
                                                            btr = Trans1.GetObject(BlockTable1[block_name], OpenMode.ForRead) as BlockTableRecord;
                                                        }






                                                        double sta1 = -1.23456;

                                                        if (ds_main.dt_points.Rows[i][col_2dsta] != DBNull.Value)
                                                        {
                                                            sta1 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_2dsta]);
                                                        }
                                                        if (ds_main.dt_points.Rows[i][col_3dsta] != DBNull.Value)
                                                        {
                                                            sta1 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_3dsta]);
                                                        }



                                                        if (sta1 >= 0)
                                                        {

                                                            if (sta1 > Poly2D.Length) sta1 = Poly2D.Length - 0.001;
                                                            if (btr != null)
                                                            {
                                                                BlockReference br = new BlockReference(Poly2D.GetPointAtDist(sta1), btr.ObjectId);
                                                                br.Layer = layer1;
                                                                BTrecord.AppendEntity(br);
                                                                Trans1.AddNewlyCreatedDBObject(br, true);

                                                                dt_od_pt.Rows.Add();
                                                                dt_od_pt.Rows[dt_od_pt.Rows.Count - 1]["id"] = br.ObjectId;
                                                                dt_od_pt.Rows[dt_od_pt.Rows.Count - 1][col_item_no] = mat1;
                                                                dt_od_pt.Rows[dt_od_pt.Rows.Count - 1][col_sta] = sta1;
                                                                dt_od_pt.Rows[dt_od_pt.Rows.Count - 1][col_descr] = descr1;

                                                            }
                                                        }
                                                    }

                                                }

                                            }
                                        }
                                    }
                                }


                            }
                        }
                    }



                    // Commit changes if user accepted, otherwise discard
                    Trans1.Commit();
                }

                attach_od_to_points(dt_od_pt);

            }



        }

        private void dataGridView_pt_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex > -1)
            {

                if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        if (dt_mat_library.Rows[i][col_item_no] != DBNull.Value && dt_mat_library.Rows[i][col_category] != DBNull.Value && dt_mat_library.Rows[i][col_type] != DBNull.Value)
                        {

                            string category1 = Convert.ToString(dt_mat_library.Rows[i][col_category]);
                            string type1 = Convert.ToString(dt_mat_library.Rows[i][col_type]);

                            string ct = category1 + "_" + type1;

                            DataGridView dgv1 = sender as DataGridView;
                            if (dgv1 != null)
                            {
                                if (dgv1.Name == "dgv_" + ct)
                                {
                                    if (dgv1.Columns[e.ColumnIndex].Name == col_item_no)
                                    {
                                        foreach (TabPage tab1 in flatTabControl1.TabPages)
                                        {
                                            foreach (Panel panel1 in tab1.Controls)
                                            {
                                                foreach (Control ctrl1 in panel1.Controls)
                                                {
                                                    ComboBox combo1 = ctrl1 as ComboBox;
                                                    if (combo1 != null && combo1.Name.Replace("combo_", "") == ct)
                                                    {
                                                        DataGridViewComboBoxCell cbox = new DataGridViewComboBoxCell();
                                                        cbox.Style.BackColor = Color.FromArgb(51, 51, 55);
                                                        cbox.Style.ForeColor = Color.White;
                                                        cbox.Style.SelectionBackColor = Color.FromArgb(51, 51, 55);
                                                        cbox.Style.SelectionForeColor = Color.White;
                                                        cbox.Style.Padding = new Padding(4, 0, 0, 0);
                                                        dgv1[e.ColumnIndex, e.RowIndex] = cbox;
                                                        if (combo1.Items.Count > 0)
                                                        {
                                                            cbox.DataSource = combo1.Items;
                                                        }
                                                        i = dt_mat_library.Rows.Count;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (dgv1.Columns[e.ColumnIndex].Name == col_layer)
                                    {
                                        foreach (TabPage tab1 in flatTabControl1.TabPages)
                                        {
                                            foreach (Panel panel1 in tab1.Controls)
                                            {
                                                foreach (Control ctrl1 in panel1.Controls)
                                                {
                                                    DataGridViewComboBoxCell cbox = new DataGridViewComboBoxCell();
                                                    cbox.Style.BackColor = Color.FromArgb(51, 51, 55);
                                                    cbox.Style.ForeColor = Color.White;
                                                    cbox.Style.SelectionBackColor = Color.FromArgb(51, 51, 55);
                                                    cbox.Style.SelectionForeColor = Color.White;
                                                    cbox.Style.Padding = new Padding(4, 0, 0, 0);

                                                    List<string> lista1 = get_layers_from_current_drawing();
                                                    if (lista1.Count > 0)
                                                    {
                                                        cbox.DataSource = lista1;
                                                    }

                                                    dgv1[e.ColumnIndex, e.RowIndex] = cbox;
                                                    i = dt_mat_library.Rows.Count;
                                                }
                                            }
                                        }
                                    }

                                    if (dgv1.Columns[e.ColumnIndex].Name == col_MSblock)
                                    {
                                        foreach (TabPage tab1 in flatTabControl1.TabPages)
                                        {
                                            foreach (Panel panel1 in tab1.Controls)
                                            {
                                                foreach (Control ctrl1 in panel1.Controls)
                                                {
                                                    DataGridViewComboBoxCell cbox = new DataGridViewComboBoxCell();
                                                    cbox.Style.BackColor = Color.FromArgb(51, 51, 55);
                                                    cbox.Style.ForeColor = Color.White;
                                                    cbox.Style.SelectionBackColor = Color.FromArgb(51, 51, 55);
                                                    cbox.Style.SelectionForeColor = Color.White;
                                                    cbox.Style.Padding = new Padding(4, 0, 0, 0);

                                                    string existing_val = "";
                                                    if (dgv1[e.ColumnIndex, e.RowIndex].Value != DBNull.Value)
                                                    {
                                                        existing_val = Convert.ToString(dgv1[e.ColumnIndex, e.RowIndex].Value);
                                                    }

                                                    List<string> lista1 = Functions.get_blocks_from_current_drawing();
                                                    if (lista1.Contains(existing_val) == false)
                                                    {
                                                        dgv1[e.ColumnIndex, e.RowIndex].Value = DBNull.Value;
                                                    }

                                                    if (lista1.Count > 0)
                                                    {
                                                        cbox.DataSource = lista1;
                                                    }

                                                    dgv1[e.ColumnIndex, e.RowIndex] = cbox;
                                                    i = dt_mat_library.Rows.Count;
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }

        private void button_line_refresh_Click(object sender, EventArgs e)
        {
            Button bt1 = sender as Button;

            if (ct_list != null && ct_list.Count > 0)
            {
                for (int i = 0; i < ct_list.Count; ++i)
                {
                    if (bt1 != null && bt1.Name.ToUpper() == "BUTTON_REFRESH_" + ct_list[i])
                    {
                        refresh_extra(ct_list[i]);
                    }
                }
            }
        }

        public void refresh_extra(string ct1)
        {
            System.Data.DataTable dt_cl = ds_main.dt_centerline;

            if (dt_cl == null || dt_cl.Rows.Count < 2)
            {
                MessageBox.Show("No centerline loaded\r\nOperation aborted.", "Points Materials", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            System.Data.DataTable dt_od_extra = new System.Data.DataTable();
            dt_od_extra.Columns.Add(pipe_us_od_item_no, typeof(string));
            dt_od_extra.Columns.Add(pipe_us_od_descr, typeof(string));
            dt_od_extra.Columns.Add(pipe_us_od_cat, typeof(string));
            dt_od_extra.Columns.Add(pipe_us_od_sta1, typeof(string));
            dt_od_extra.Columns.Add(pipe_us_od_sta2, typeof(string));
            dt_od_extra.Columns.Add("id", typeof(ObjectId));

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    if (ds_main.dt_extra != null && ds_main.dt_extra.Rows.Count > 0)
                    {
                        if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
                        {

                            Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);


                            create_pipe_us_od_table(extra_od);

                            List<string> lista_mat = new List<string>();
                            List<string> lista_descr = new List<string>();
                            List<string> lista_layer = new List<string>();
                            List<string> lista_category = new List<string>();

                            for (int i = 0; i < ds_main.dt_extra.Rows.Count; ++i)
                            {
                                if (ds_main.dt_extra.Rows[i][col_item_no] != DBNull.Value)
                                {
                                    string mat1 = Convert.ToString(ds_main.dt_extra.Rows[i][col_item_no]);

                                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                    {
                                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                                        {
                                            if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value && dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                                            {
                                                string cat1 = Convert.ToString(dt_mat_library.Rows[j][col_category]).ToUpper().Replace(" ", "");
                                                string type1 = Convert.ToString(dt_mat_library.Rows[j][col_type]).ToUpper().Replace(" ", "");
                                                string descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);

                                                if (cat1 + "_" + type1 == ct1)
                                                {
                                                    lista_mat.Add(mat1);
                                                    lista_descr.Add(descr1);
                                                    string layer1 = "0";
                                                    if (dt_mat_library.Rows[j][col_layer] != DBNull.Value)
                                                    {
                                                        layer1 = Convert.ToString(dt_mat_library.Rows[j][col_layer]);
                                                    }
                                                    lista_layer.Add(layer1);
                                                    lista_category.Add(cat1);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            foreach (ObjectId id1 in BTrecord)
                            {
                                Curve poly1 = Trans1.GetObject(id1, OpenMode.ForRead) as Curve;
                                if (poly1 != null)
                                {

                                    if (lista_layer.Contains(poly1.Layer) == true)
                                    {
                                        poly1.UpgradeOpen();
                                        poly1.Erase();
                                    }
                                }
                            }


                            for (int i = 0; i < ds_main.dt_extra.Rows.Count; ++i)
                            {


                                if (ds_main.dt_extra.Rows[i][col_item_no] != DBNull.Value)
                                {
                                    string mat1 = Convert.ToString(ds_main.dt_extra.Rows[i][col_item_no]);
                                    if (lista_mat.Contains(mat1) == true)
                                    {
                                        double sta1 = -1.23456;
                                        double sta2 = -1.23456;

                                        if (ds_main.dt_extra.Rows[i][pipe_col_2d1] != DBNull.Value)
                                        {
                                            sta1 = Convert.ToDouble(ds_main.dt_extra.Rows[i][pipe_col_2d1]);
                                        }
                                        if (ds_main.dt_extra.Rows[i][pipe_col_3d1] != DBNull.Value)
                                        {
                                            sta1 = Convert.ToDouble(ds_main.dt_extra.Rows[i][pipe_col_3d1]);
                                        }
                                        if (ds_main.dt_extra.Rows[i][pipe_col_2d2] != DBNull.Value)
                                        {
                                            sta2 = Convert.ToDouble(ds_main.dt_extra.Rows[i][pipe_col_2d2]);
                                        }
                                        if (ds_main.dt_extra.Rows[i][pipe_col_3d2] != DBNull.Value)
                                        {
                                            sta2 = Convert.ToDouble(ds_main.dt_extra.Rows[i][pipe_col_3d2]);
                                        }

                                        if (sta1 >= 0 && sta2 > 0)
                                        {
                                            if (sta1 > Poly2D.Length) sta1 = Poly2D.Length - 0.001;
                                            if (sta2 > Poly2D.Length) sta2 = Poly2D.Length - 0.001;

                                            double param1 = Poly2D.GetParameterAtDistance(sta1);
                                            double param2 = Poly2D.GetParameterAtDistance(sta2)
                                                ;
                                            Functions.Creaza_layer(lista_layer[lista_mat.IndexOf(mat1)], 1, true);
                                            Polyline poly1 = Functions.get_part_of_poly(Poly2D, param1, param2);
                                            poly1.Layer = lista_layer[lista_mat.IndexOf(mat1)];
                                            BTrecord.AppendEntity(poly1);
                                            Trans1.AddNewlyCreatedDBObject(poly1, true);

                                            dt_od_extra.Rows.Add();
                                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1]["id"] = poly1.ObjectId;
                                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_sta1] = sta1;
                                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_sta2] = sta2;
                                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_descr] = lista_descr[lista_mat.IndexOf(mat1)];
                                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_item_no] = mat1;
                                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_cat] = lista_category[lista_mat.IndexOf(mat1)];
                                        }
                                    }
                                }
                            }
                        }
                    }



                    // Commit changes if user accepted, otherwise discard
                    Trans1.Commit();
                }

                attach_od_to_us_pipes(dt_od_extra);

            }



        }

        private void button_lin_dwg_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt_cl = ds_main.dt_centerline;



            if (dt_cl == null || dt_cl.Rows.Count < 2)
            {
                MessageBox.Show("No centerline loaded\r\nOperation aborted", "material design", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            Button bt1 = sender as System.Windows.Forms.Button;
            string ct1 = bt1.Name.Replace("button_lin_draw_", "");
            string mat1 = "**";
            string layer1 = "**";
            string descr1 = "none";
            string cat1 = "**";

            foreach (TabPage tab1 in flatTabControl1.TabPages)
            {
                foreach (Panel panel1 in tab1.Controls)
                {
                    foreach (Control ctrl1 in panel1.Controls)
                    {
                        ComboBox combo1 = ctrl1 as ComboBox;
                        if (combo1 != null && combo1.Name.Replace("combo_lin_", "") == ct1)
                        {
                            mat1 = combo1.Text;
                        }
                    }
                }
            }

            if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
            {
                for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                {
                    if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && dt_mat_library.Rows[j][col_layer] != DBNull.Value)
                    {
                        string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]);
                        string layer2 = Convert.ToString(dt_mat_library.Rows[j][col_layer]);

                        if (mat1.ToUpper() == mat2.ToUpper())
                        {
                            layer1 = layer2;
                            Functions.Creaza_layer(layer1, 1, true);

                            if (dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                            {
                                descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                            }

                            if (dt_mat_library.Rows[j][col_category] != DBNull.Value)
                            {
                                cat1 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                            }

                            j = dt_mat_library.Rows.Count;
                        }
                    }
                }
            }

            if (mat1 == "**")
            {
                MessageBox.Show("Specify the current item number\r\nOperation aborted", "material design", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (layer1 == "**")
            {
                MessageBox.Show("the dwg layer specified is not valid\r\nOperation aborted", "material design", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    ObjectId p3did = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(dt_cl);
                        Trans1.Commit();

                        p3did = Poly3D.ObjectId;

                    }

                    System.Data.DataTable dt_od_extra = new System.Data.DataTable();
                    dt_od_extra.Columns.Add(pipe_us_od_item_no, typeof(string));
                    dt_od_extra.Columns.Add(pipe_us_od_descr, typeof(string));
                    dt_od_extra.Columns.Add(pipe_us_od_cat, typeof(string));
                    dt_od_extra.Columns.Add(pipe_us_od_sta1, typeof(string));
                    dt_od_extra.Columns.Add(pipe_us_od_sta2, typeof(string));
                    dt_od_extra.Columns.Add("id", typeof(ObjectId));

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        this.MdiParent.WindowState = FormWindowState.Minimized;
                        Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);
                        Polyline3d Poly3D = Trans1.GetObject(p3did, OpenMode.ForWrite) as Polyline3d;

                        create_pipe_us_od_table(extra_od);

                        if (Poly3D == null)
                        {
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        ObjectIdCollection objectid_col = new ObjectIdCollection();
                        objectid_col.Add(Poly3D.ObjectId);
                        DrawOrderTable DrawOrderTable1 = Trans1.GetObject(BTrecord.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                        DrawOrderTable1.MoveToBottom(objectid_col);



                        double end_sta = Math.Round(Poly3D.Length, 2);
                        if (end_sta > Poly3D.Length)
                        {
                            end_sta = end_sta - 0.01;
                        }

                        bool is3D = ds_main.is3D;

                        if (ds_main.dt_extra == null)
                        {
                            ds_main.dt_extra = Creaza_dt_mat_linear_structure();
                        }

                        //Trans1.TransactionManager.QueueForGraphicsFlush();

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions pp1;
                        pp1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify start point");
                        pp1.AllowNone = false;

                        pp1.Keywords.Add("Middle");
                        pp1.Keywords.Add("Buffer from point");
                        pp1.Keywords.Add("Feature");

                        Point_res1 = Editor1.GetPoint(pp1);

                        if (Point_res1.Status != PromptStatus.OK && Point_res1.Status != PromptStatus.Keyword)
                        {
                            Poly3D.Erase();
                            Trans1.Commit();
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        double par1 = -1;
                        double par2 = -1;
                        double station_start = -1;
                        double station_end = -1;
                        if (Point_res1.Status == PromptStatus.Keyword)
                        {
                            #region keyword middle
                            PromptPointOptions ppm;
                            ppm = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify mid point:");
                            ppm.AllowNone = false;

                            if (Point_res1.StringResult.ToLower() == "middle")
                            {
                                PromptPointResult Point_resm = Editor1.GetPoint(ppm);
                                if (Point_resm.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }
                                Point3d point_middle = Point_resm.Value;



                                Autodesk.AutoCAD.EditorInput.PromptDistanceOptions Prompt_len = new Autodesk.AutoCAD.EditorInput.PromptDistanceOptions("\n" + "Specify length:");
                                Prompt_len.AllowNegative = false;
                                Prompt_len.AllowZero = true;
                                Prompt_len.AllowNone = true;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_len = ThisDrawing.Editor.GetDistance(Prompt_len);
                                if (Rezultat_len.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                double len1 = Rezultat_len.Value;
                                if (len1 < 1)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }
                                Point3d p1 = Poly2D.GetClosestPointTo(point_middle, Vector3d.ZAxis, false);



                                double param_m = Poly2D.GetParameterAtPoint(p1);
                                if (param_m > Poly3D.EndParam) param_m = Poly3D.EndParam;

                                double sta_m = Poly3D.GetDistanceAtParameter(param_m);
                                station_start = sta_m - len1 / 2;
                                station_end = sta_m + len1 / 2;

                                station_start = Math.Round(station_start, 0);
                                station_end = Math.Round(station_end, 0);
                                if (station_start < 0) station_start = 0;

                                if (sta_m + len1 / 2 == Poly3D.Length)
                                {
                                    station_end = Poly3D.Length - 0.0001;
                                }

                            }
                            #endregion

                            #region keyword buffer from point
                            if (Point_res1.StringResult.ToLower() == "buffer from point")
                            {
                                PromptPointOptions ppb1;
                                ppb1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify first point:");
                                ppb1.AllowNone = false;

                                PromptPointOptions ppb2;
                                ppb2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify second point:");
                                ppb2.AllowNone = false;


                                PromptPointResult Point_resb1 = Editor1.GetPoint(ppb1);
                                if (Point_resb1.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }
                                Point3d point_b1 = Point_resb1.Value;



                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_b1 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify buffer value:");
                                Prompt_b1.AllowNegative = false;
                                Prompt_b1.AllowZero = true;
                                Prompt_b1.AllowNone = true;
                                Prompt_b1.UseDefaultValue = true;
                                Prompt_b1.DefaultValue = 0;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_b1 = ThisDrawing.Editor.GetDouble(Prompt_b1);
                                if (Rezultat_b1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }



                                PromptPointResult Point_resb2 = Editor1.GetPoint(ppb2);
                                if (Point_resb2.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }
                                Point3d point_b2 = Point_resb2.Value;



                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_b2 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify buffer value:");
                                Prompt_b2.AllowNegative = false;
                                Prompt_b2.AllowZero = true;
                                Prompt_b2.AllowNone = true;
                                Prompt_b2.UseDefaultValue = true;
                                Prompt_b2.DefaultValue = 0;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_b2 = ThisDrawing.Editor.GetDouble(Prompt_b2);
                                if (Rezultat_b2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                double buffer1 = Rezultat_b1.Value;
                                Point3d p1 = Poly2D.GetClosestPointTo(point_b1, Vector3d.ZAxis, false);
                                double param_b1 = Poly2D.GetParameterAtPoint(p1);
                                if (param_b1 > Poly3D.EndParam) param_b1 = Poly3D.EndParam;
                                double sta_b1 = Poly3D.GetDistanceAtParameter(param_b1);




                                double buffer2 = Rezultat_b2.Value;
                                Point3d p2 = Poly2D.GetClosestPointTo(point_b2, Vector3d.ZAxis, false);
                                double param_b2 = Poly2D.GetParameterAtPoint(p2);
                                if (param_b2 > Poly3D.EndParam) param_b2 = Poly3D.EndParam;
                                double sta_b2 = Poly3D.GetDistanceAtParameter(param_b2);


                                if (sta_b2 > sta_b1)
                                {
                                    station_start = sta_b1 - buffer1;
                                    station_end = sta_b2 + buffer2;
                                }
                                else
                                {
                                    station_start = sta_b2 - buffer2;
                                    station_end = sta_b1 + buffer1;
                                }



                                station_start = Math.Round(station_start, 0);
                                station_end = Math.Round(station_end, 0);

                                if (station_start < 0) station_start = 0;


                                if (sta_b2 > sta_b1)
                                {
                                    if (sta_b2 + buffer2 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }
                                else
                                {
                                    if (sta_b1 + buffer1 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }



                            }
                            #endregion

                            #region keyword feature


                            if (Point_res1.StringResult.ToLower() == "feature")
                            {
                                Point3d point_feat1 = new Point3d();
                                Point3d point_feat2 = new Point3d();

                                Autodesk.AutoCAD.EditorInput.PromptEntityResult rezultat_feat1;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_feat1;
                                prompt_feat1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the feature:");
                                prompt_feat1.SetRejectMessage("\nSelect a polyline!");
                                prompt_feat1.AllowNone = true;
                                prompt_feat1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                rezultat_feat1 = ThisDrawing.Editor.GetEntity(prompt_feat1);

                                if (rezultat_feat1.Status != PromptStatus.OK)
                                {
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    set_enable_true();
                                    return;
                                }

                                Polyline poly1 = Trans1.GetObject(rezultat_feat1.ObjectId, OpenMode.ForWrite) as Polyline;
                                poly1.Elevation = 0;

                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly1, Poly2D);
                                Point3dCollection colint2 = new Point3dCollection();

                                if (colint1.Count == 0)
                                {

                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "no intersection");
                                    set_enable_true();
                                    return;

                                }
                                else if (colint1.Count == 1)
                                {
                                    point_feat1 = colint1[0];

                                    Autodesk.AutoCAD.EditorInput.PromptEntityResult rezultat_feat2;
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_feat2;
                                    prompt_feat2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the other feature:");
                                    prompt_feat2.SetRejectMessage("\nSelect a polyline!");
                                    prompt_feat2.AllowNone = true;
                                    prompt_feat2.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                    rezultat_feat2 = ThisDrawing.Editor.GetEntity(prompt_feat2);

                                    if (rezultat_feat2.Status != PromptStatus.OK)
                                    {

                                        PromptPointOptions ppf2;
                                        ppf2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify second point:");
                                        ppf2.AllowNone = false;

                                        PromptPointResult Point_resf2 = Editor1.GetPoint(ppf2);
                                        if (Point_resf2.Status != PromptStatus.OK)
                                        {
                                            this.WindowState = FormWindowState.Normal;
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            set_enable_true();
                                            return;
                                        }
                                        Point3d point_f2 = Point_resf2.Value;
                                        point_feat2 = Poly2D.GetClosestPointTo(point_f2, Vector3d.ZAxis, false);
                                    }
                                    else
                                    {
                                        Polyline poly2 = Trans1.GetObject(rezultat_feat2.ObjectId, OpenMode.ForWrite) as Polyline;
                                        poly2.Elevation = 0;
                                        colint2 = Functions.Intersect_on_both_operands(poly2, Poly2D);
                                        if (colint2.Count == 0)
                                        {

                                            this.WindowState = FormWindowState.Normal;
                                            ThisDrawing.Editor.WriteMessage("\n" + "no intersection");
                                            set_enable_true();
                                            return;

                                        }
                                        else if (colint2.Count == 1)
                                        {
                                            point_feat2 = colint2[0];
                                        }
                                        else if (colint2.Count > 1)
                                        {
                                            PromptPointOptions pp_int2;
                                            pp_int2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify intersection point:");
                                            pp_int2.AllowNone = false;

                                            PromptPointResult rezult_close2 = Editor1.GetPoint(pp_int2);
                                            if (rezult_close2.Status != PromptStatus.OK)
                                            {
                                                this.WindowState = FormWindowState.Normal;
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                set_enable_true();
                                                return;
                                            }
                                            Point3d point_close2 = rezult_close2.Value;

                                            double dmin = 1000;
                                            for (int i = 0; i < colint2.Count; ++i)
                                            {
                                                Point3d pint2 = colint2[i];

                                                double d1 = Math.Pow(Math.Pow(point_close2.X - pint2.X, 2) + Math.Pow(point_close2.Y - pint2.Y, 2), 0.5);

                                                if (d1 < dmin)
                                                {
                                                    dmin = d1;
                                                    point_feat2 = pint2;
                                                }


                                            }




                                        }
                                    }



                                }
                                else if (colint1.Count == 2)
                                {
                                    point_feat1 = colint1[0];
                                    point_feat2 = colint1[1];

                                }
                                else if (colint1.Count > 2)
                                {
                                    PromptPointOptions pp_int1;
                                    pp_int1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify intersection point #1:");
                                    pp_int1.AllowNone = false;

                                    PromptPointResult rezult_close1 = Editor1.GetPoint(pp_int1);
                                    if (rezult_close1.Status != PromptStatus.OK)
                                    {
                                        this.WindowState = FormWindowState.Normal;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        set_enable_true();
                                        return;
                                    }
                                    Point3d point_close1 = rezult_close1.Value;

                                    double dmin1 = 1000;
                                    for (int i = 0; i < colint1.Count; ++i)
                                    {
                                        Point3d pint1 = colint1[i];

                                        double d1 = Math.Pow(Math.Pow(point_close1.X - pint1.X, 2) + Math.Pow(point_close1.Y - pint1.Y, 2), 0.5);

                                        if (d1 < dmin1)
                                        {
                                            dmin1 = d1;
                                            point_feat1 = pint1;
                                        }


                                    }


                                    PromptPointOptions pp_int2;
                                    pp_int2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify intersection point #2:");
                                    pp_int2.AllowNone = false;

                                    PromptPointResult rezult_close2 = Editor1.GetPoint(pp_int2);
                                    if (rezult_close2.Status != PromptStatus.OK)
                                    {
                                        this.WindowState = FormWindowState.Normal;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        set_enable_true();
                                        return;
                                    }
                                    Point3d point_close2 = rezult_close2.Value;

                                    double dmin2 = 1000;
                                    for (int i = 0; i < colint1.Count; ++i)
                                    {
                                        Point3d pint2 = colint1[i];

                                        double d2 = Math.Pow(Math.Pow(point_close2.X - pint2.X, 2) + Math.Pow(point_close2.Y - pint2.Y, 2), 0.5);

                                        if (d2 < dmin2)
                                        {
                                            dmin2 = d2;
                                            point_feat2 = pint2;
                                        }


                                    }

                                }








                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions prompt_buffer = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify buffer value:");
                                prompt_buffer.AllowNegative = false;
                                prompt_buffer.AllowZero = true;
                                prompt_buffer.AllowNone = true;
                                prompt_buffer.UseDefaultValue = true;
                                prompt_buffer.DefaultValue = 0;
                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult rez_buffer = ThisDrawing.Editor.GetDouble(prompt_buffer);
                                if (rez_buffer.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Poly3D.Erase();
                                    Trans1.Commit();
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }








                                double buffer1 = rez_buffer.Value;

                                double param_b1 = Poly2D.GetParameterAtPoint(point_feat1);
                                if (param_b1 > Poly3D.EndParam) param_b1 = Poly3D.EndParam;


                                double sta_b1 = Poly3D.GetDistanceAtParameter(param_b1);


                                double param_b2 = Poly2D.GetParameterAtPoint(point_feat2);
                                if (param_b2 > Poly3D.EndParam) param_b2 = Poly3D.EndParam;
                                double sta_b2 = Poly3D.GetDistanceAtParameter(param_b2);


                                if (sta_b2 > sta_b1)
                                {
                                    station_start = sta_b1 - buffer1;
                                    station_end = sta_b2 + buffer1;
                                }
                                else
                                {
                                    station_start = sta_b2 - buffer1;
                                    station_end = sta_b1 + buffer1;
                                }



                                station_start = Math.Round(station_start, 0);
                                station_end = Math.Round(station_end, 0);

                                if (station_start < 0) station_start = 0;


                                if (sta_b2 > sta_b1)
                                {
                                    if (sta_b2 + buffer1 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }
                                else
                                {
                                    if (sta_b1 + buffer1 == Poly3D.Length)
                                    {
                                        station_end = Poly3D.Length - 0.0001;
                                    }
                                }



                            }
                            #endregion
                        }
                        else
                        {
                            #region pick 2 pts
                            Point3d p1 = Poly2D.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);
                            par1 = Poly2D.GetParameterAtPoint(p1);
                            if (par1 > Poly3D.EndParam) par1 = Poly3D.EndParam;


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions pp2;
                            pp2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify end point");
                            pp2.AllowNone = false;
                            pp2.UseBasePoint = true;
                            pp2.BasePoint = p1;


                            Point_res2 = Editor1.GetPoint(pp2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                Poly3D.Erase();
                                Trans1.Commit();
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                return;
                            }

                            Point3d p2 = Poly2D.GetClosestPointTo(Point_res2.Value, Vector3d.ZAxis, false);
                            par2 = Poly2D.GetParameterAtPoint(p2);
                            if (par2 > Poly3D.EndParam) par2 = Poly3D.EndParam;

                            station_start = Math.Round(Poly3D.GetDistanceAtParameter(par1), 0);
                            station_end = Math.Round(Poly3D.GetDistanceAtParameter(par2), 0);
                            if (par2 == Poly3D.EndParam)
                            {
                                station_end = Poly3D.Length - 0.0001;
                            }
                            #endregion
                        }

                        if (station_start >= 0 && station_end >= 0)
                        {
                            if (station_start > end_sta)
                            {
                                station_start = end_sta;
                            }

                            if (station_end > end_sta)
                            {
                                station_end = end_sta;
                            }

                            Point3d pt_start = Poly3D.GetPointAtDist(station_start);
                            Point3d pt_end = Poly3D.GetPointAtDist(station_end);

                            populate_dt_extra(ref ds_main.dt_extra, station_start, station_end, mat1, pt_start, pt_end);
                            dt_ct[ct_list.IndexOf(ct1)].Rows.Clear();
                            for (int i = 0; i < ds_main.dt_extra.Rows.Count; ++i)
                            {
                                if (ds_main.dt_extra.Rows[i][col_altdesc] != DBNull.Value)
                                {
                                    string cat2 = "**";
                                    string type2 = "**";

                                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                    {
                                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                                        {
                                            if (dt_mat_library.Rows[j][col_category] != DBNull.Value)
                                            {
                                                cat2 = Convert.ToString(dt_mat_library.Rows[j][col_category]);
                                            }

                                            if (dt_mat_library.Rows[j][col_type] != DBNull.Value)
                                            {
                                                type2 = Convert.ToString(dt_mat_library.Rows[j][col_type]);
                                            }
                                        }
                                    }

                                    string ct2 = cat2 + "_" + type2;
                                    if (ct1 == ct2)
                                    {
                                        System.Data.DataRow row1 = dt_ct[ct_list.IndexOf(ct1)].NewRow();
                                        row1.ItemArray = ds_main.dt_extra.Rows[i].ItemArray;
                                        dt_ct[ct_list.IndexOf(ct1)].Rows.Add(row1);
                                    }
                                }
                            }

                            double param1 = Poly3D.GetParameterAtDistance(station_start);
                            double param2 = Poly3D.GetParameterAtDistance(station_end);
                            if (param2 > Poly2D.EndParam) param2 = Poly2D.EndParam;

                            Polyline poly1 = Functions.get_part_of_poly(Poly2D, param1, param2);
                            poly1.Layer = layer1;

                            BTrecord.AppendEntity(poly1);
                            Trans1.AddNewlyCreatedDBObject(poly1, true);

                            dt_od_extra.Rows.Add();
                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1]["id"] = poly1.ObjectId;
                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_sta1] = station_start;
                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_sta2] = station_end;
                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_descr] = descr1;
                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_item_no] = mat1;
                            dt_od_extra.Rows[dt_od_extra.Rows.Count - 1][pipe_us_od_cat] = cat1;

                            // Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_ct[list_ct.IndexOf(ct1)]);
                        }

                        Poly3D.Erase();
                        Trans1.Commit();
                    }
                    attach_od_to_us_pipes(dt_od_extra);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        private void button_new_mat_library_Click(object sender, EventArgs e)
        {
            dt_mat_library = Creaza_mat_library_structure();
            ds_main.dt_points = Creaza_dt_points_structure();
            ds_main.dt_pipe = Creaza_dt_mat_linear_structure();
            ds_main.dt_extra = Creaza_dt_mat_linear_structure();

            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W2 = null;
            Microsoft.Office.Interop.Excel.Worksheet W3 = null;
            Microsoft.Office.Interop.Excel.Worksheet W4 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

            bool is_opened = false;
            bool do_not_save = false;

            try
            {
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }

                Excel1.Visible = true;

                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName == ds_main.config_xls)
                        {
                            Workbook1 = Workbook2;
                            is_opened = true;
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                            {
                                if (Wx.Name == "MatDesc")
                                {
                                    W1 = Wx;
                                }
                                if (Wx.Name == "MatPipe")
                                {
                                    W2 = Wx;
                                }
                                if (Wx.Name == "MatPoints")
                                {
                                    W3 = Wx;
                                }
                                if (Wx.Name == "MatOther")
                                {
                                    W4 = Wx;
                                }
                            }
                        }
                    }

                    if (is_opened == false)
                    {
                        Workbook1 = Excel1.Workbooks.Open(ds_main.config_xls);
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name == "MatDesc")
                            {
                                W1 = Wx;
                            }
                            if (Wx.Name == "MatPipe")
                            {
                                W2 = Wx;
                            }
                            if (Wx.Name == "MatPoints")
                            {
                                W3 = Wx;
                            }
                            if (Wx.Name == "MatOther")
                            {
                                W4 = Wx;
                            }
                        }
                    }

                }
                else
                {
                    Workbook1 = Excel1.Workbooks.Add();
                    do_not_save = true;
                }

                if (W4 == null)
                {
                    W4 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W4.Name = "MatOther";
                }

                if (W3 == null)
                {
                    W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W3.Name = "MatPoints";
                }

                if (W2 == null)
                {
                    W2 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W2.Name = "MatPipe";
                }

                if (W1 == null)
                {
                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W1.Name = "MatDesc";
                }







                string client1 = ds_main.tpage_main.get_textbox_client_name();
                string project1 = ds_main.tpage_main.get_textbox_project();
                string segment1 = ds_main.tpage_main.get_textbox_segment();

                Create_header_material_library(W1, client1, project1, segment1, dt_mat_library);
                Create_header_material_linear_file(W2, client1, project1, segment1, ds_main.dt_pipe);
                Create_header_material_points_file(W3, client1, project1, segment1, ds_main.dt_points);
                Create_header_material_linear_file(W4, client1, project1, segment1, ds_main.dt_extra, "Material Linear Other");

                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1 = dt_mat_library.Clone();
                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "BELOW IS AN EXAMPLE OF A MATERIAL DESCRIPTION TABLE";
                dt1.Rows.Add();
                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "1";
                dt1.Rows[dt1.Rows.Count - 1][2] = "PIPE, 42.000 OD X 0.443 WT, API-5L, X75, DSAW, FBE COATED";
                dt1.Rows[dt1.Rows.Count - 1][3] = "PIPE";
                dt1.Rows[dt1.Rows.Count - 1][4] = "LINEAR";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_Pipe_01_Linear";
                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "2";
                dt1.Rows[dt1.Rows.Count - 1][2] = "PIPE, 42.000 OD X 0.469 WT, API-5L, X75, DSAW, FBE COATED";
                dt1.Rows[dt1.Rows.Count - 1][3] = "PIPE";
                dt1.Rows[dt1.Rows.Count - 1][4] = "LINEAR";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_Pipe_02_Linear";
                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "3";
                dt1.Rows[dt1.Rows.Count - 1][2] = "LAUNCHER";
                dt1.Rows[dt1.Rows.Count - 1][3] = "FAB";
                dt1.Rows[dt1.Rows.Count - 1][4] = "LINEAR";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_FAB_03_Linear";
                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "4";
                dt1.Rows[dt1.Rows.Count - 1][2] = "ELBOW TYPE 1";
                dt1.Rows[dt1.Rows.Count - 1][3] = "ELL";
                dt1.Rows[dt1.Rows.Count - 1][4] = "POINT";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_ELL_04_POINT";
                dt1.Rows[dt1.Rows.Count - 1][6] = "ELBOW_BLOCK_1";
                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "5";
                dt1.Rows[dt1.Rows.Count - 1][2] = "ELBOW TYPE 2";
                dt1.Rows[dt1.Rows.Count - 1][3] = "ELL";
                dt1.Rows[dt1.Rows.Count - 1][4] = "POINT";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_ELL_05_POINT";
                dt1.Rows[dt1.Rows.Count - 1][6] = "ELBOW_BLOCK_2";

                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "6";
                dt1.Rows[dt1.Rows.Count - 1][2] = "DITCH PLUG";
                dt1.Rows[dt1.Rows.Count - 1][3] = "DP";
                dt1.Rows[dt1.Rows.Count - 1][4] = "POINT";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_DP_06_POINT";
                dt1.Rows[dt1.Rows.Count - 1][6] = "DP_BLOCK";

                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "7";
                dt1.Rows[dt1.Rows.Count - 1][2] = "TEST LEAD";
                dt1.Rows[dt1.Rows.Count - 1][3] = "TL";
                dt1.Rows[dt1.Rows.Count - 1][4] = "POINT";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_TL_07_POINT";
                dt1.Rows[dt1.Rows.Count - 1][6] = "TL_BLOCK";

                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1][1] = "8";
                dt1.Rows[dt1.Rows.Count - 1][2] = "MLV";
                dt1.Rows[dt1.Rows.Count - 1][3] = "FAB";
                dt1.Rows[dt1.Rows.Count - 1][4] = "LINEAR";
                dt1.Rows[dt1.Rows.Count - 1][5] = "MAT_FAB_08_Linear";

                Functions.Transfer_datatable_to_existing_excel_spreadsheet(W1, dt1, "A", 14);




                if (do_not_save == true)
                {
                    Workbook1.Save();
                }


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                if (W3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                if (W4 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W4);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }

        private void comboBox_linear_current_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox1 = sender as ComboBox;
            string mat1 = comboBox1.Text.ToUpper();
            string ct1 = comboBox1.Name.Replace("combo_lin_", "").ToUpper();
            foreach (TabPage tab1 in flatTabControl1.TabPages)
            {
                foreach (Panel panel1 in tab1.Controls)
                {
                    foreach (Control ctrl1 in panel1.Controls)
                    {
                        Label label1 = ctrl1 as Label;
                        if (label1 != null && label1.Name.Replace("lab_lin_left_", "").ToUpper() == ct1)
                        {
                            for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                            {
                                if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                                {
                                    string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper();
                                    if (mat1 == mat2)
                                    {
                                        label1.Text = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                                        j = dt_mat_library.Rows.Count;
                                    }


                                }
                            }
                        }
                    }
                }
            }
        }

        private void comboBox_current_pt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox1 = sender as ComboBox;
            string mat1 = comboBox1.Text.ToUpper();
            string ct1 = comboBox1.Name.Replace("combo_", "").ToUpper();
            foreach (TabPage tab1 in flatTabControl1.TabPages)
            {
                foreach (Panel panel1 in tab1.Controls)
                {
                    foreach (Control ctrl1 in panel1.Controls)
                    {
                        Label label1 = ctrl1 as Label;
                        if (label1 != null && label1.Name.Replace("lab_pt_left_", "").ToUpper() == ct1)
                        {
                            for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                            {
                                if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                                {
                                    string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper();
                                    if (mat1 == mat2)
                                    {
                                        label1.Text = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                                        j = dt_mat_library.Rows.Count;
                                    }


                                }
                            }
                        }
                    }
                }
            }
        }

        private void dataGridView_mat_library_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > -1)
            {
                if (dataGridView_mat_library.Columns[e.ColumnIndex].Name == col_MSblock)
                {
                    DataGridViewComboBoxCell cbox = new DataGridViewComboBoxCell();
                    cbox.Style.BackColor = Color.FromArgb(51, 51, 55);
                    cbox.Style.ForeColor = Color.White;
                    cbox.Style.SelectionBackColor = Color.FromArgb(51, 51, 55);
                    cbox.Style.SelectionForeColor = Color.White;
                    cbox.Style.Padding = new Padding(4, 0, 0, 0);

                    string existing_val = "";
                    if (dataGridView_mat_library[e.ColumnIndex, e.RowIndex].Value != DBNull.Value)
                    {
                        existing_val = Convert.ToString(dataGridView_mat_library[e.ColumnIndex, e.RowIndex].Value);
                    }

                    List<string> lista1 = Functions.get_blocks_from_current_drawing();
                    if (lista1.Contains(existing_val) == false)
                    {
                        dataGridView_mat_library[e.ColumnIndex, e.RowIndex].Value = DBNull.Value;
                    }

                    dataGridView_mat_library[e.ColumnIndex, e.RowIndex] = cbox;



                    if (lista1.Count > 0)
                    {
                        cbox.DataSource = lista1;
                    }

                }

                if (dataGridView_mat_library.Columns[e.ColumnIndex].Name == col_layer)
                {
                    DataGridViewComboBoxCell cbox = new DataGridViewComboBoxCell();
                    cbox.Style.BackColor = Color.FromArgb(51, 51, 55);
                    cbox.Style.ForeColor = Color.White;
                    cbox.Style.SelectionBackColor = Color.FromArgb(51, 51, 55);
                    cbox.Style.SelectionForeColor = Color.White;
                    cbox.Style.Padding = new Padding(4, 0, 0, 0);

                    string existing_val = "";
                    if (e.RowIndex >= 0)
                    {
                        if (dataGridView_mat_library[e.ColumnIndex, e.RowIndex].Value != DBNull.Value)
                        {
                            existing_val = Convert.ToString(dataGridView_mat_library[e.ColumnIndex, e.RowIndex].Value);
                        }

                        dataGridView_mat_library[e.ColumnIndex, e.RowIndex] = cbox;
                    }



                    List<string> lista1 = get_layers_from_current_drawing();
                    if (existing_val != "" && lista1.Contains(existing_val) == false)
                    {
                        lista1.Add(existing_val);
                    }

                    if (lista1.Count > 0)
                    {
                        cbox.DataSource = lista1;

                    }

                }

            }
        }

        private void dataGridView_mat_library_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < dt_filter.Rows.Count; ++i)
            {
                if (dt_filter.Rows[i][col_item_no] != DBNull.Value)
                {
                    string mat1 = Convert.ToString(dt_filter.Rows[i][col_item_no]);
                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                    {
                        if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                        {
                            string mat2 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]);

                            if (mat1 == mat2)
                            {
                                dt_mat_library.Rows[j][col_MSblock] = dt_filter.Rows[i][col_MSblock];
                                dt_mat_library.Rows[j][col_layer] = dt_filter.Rows[i][col_layer];
                            }

                        }
                    }
                }
            }
        }



        private List<string> get_layers_from_current_drawing()
        {
            List<string> lista1 = new List<string>();

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                    foreach (ObjectId layer_id in LayerTable1)
                    {
                        LayerTableRecord layer1 = Trans1.GetObject(layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as LayerTableRecord;

                        if (layer1.Name.Contains("*") == false && layer1.Name.Contains("|") == false)
                        {
                            lista1.Add(layer1.Name);
                        }
                    }
                    Trans1.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }


            return lista1;

        }

        private void button_save_all_Click(object sender, EventArgs e)
        {

            try
            {

                set_enable_false();
                get_client_project_segment_pipe_diam();
                string ct1 = "ELL_POINT";

                if (ct_list.Contains(ct1) == true)
                {

                    System.Data.DataTable dt1 = dt_ct[ct_list.IndexOf(ct1)];
                    List<string> lista1 = new List<string>();

                    for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                    {
                        if (dt_mat_library.Rows[j][col_category] != DBNull.Value && dt_mat_library.Rows[j][col_type] != DBNull.Value && dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                        {
                            if ((Convert.ToString(dt_mat_library.Rows[j][col_category]) + "_" + Convert.ToString(dt_mat_library.Rows[j][col_type])).ToUpper().Replace(" ", "") == ct1)
                            {
                                string mat1 = Convert.ToString(dt_mat_library.Rows[j][col_item_no]).ToUpper().Replace(" ", "");
                                lista1.Add(mat1);
                            }
                        }
                    }

                    for (int i = ds_main.dt_points.Rows.Count - 1; i >= 0; --i)
                    {
                        if (ds_main.dt_points.Rows[i][col_item_no] != DBNull.Value)
                        {
                            string mat1 = Convert.ToString(ds_main.dt_points.Rows[i][col_item_no]).ToUpper().Replace(" ", "");
                            if (lista1.Contains(mat1) == true)
                            {
                                ds_main.dt_points.Rows[i].Delete();
                            }
                        }
                    }

                    for (int j = 0; j < dt1.Rows.Count; ++j)
                    {
                        ds_main.dt_points.Rows.Add();
                        for (int k = 0; k < dt1.Columns.Count; ++k)
                        {
                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][k] = dt1.Rows[j][k];
                        }
                    }

                    if (ds_main.dt_points.Rows.Count > 0)
                    {
                        string col1 = col_2dsta;
                        if (ds_main.dt_points.Rows[0][col_3dsta] != DBNull.Value) col1 = col_3dsta;

                        ds_main.dt_points = Functions.Sort_data_table(ds_main.dt_points, col1);


                    }
                }

                if (ds_main.dt_centerline != null && ds_main.dt_centerline.Rows.Count > 1)
                {
                    Polyline Poly2D = Functions.Build_2d_poly_for_scanning(ds_main.dt_centerline);
                    for (int i = 0; i < ds_main.dt_points.Rows.Count; ++i)
                    {
                        if (ds_main.dt_points.Rows[i][col_2dsta] != DBNull.Value && (ds_main.dt_points.Rows[i][col_x] == DBNull.Value || ds_main.dt_points.Rows[i][col_y] == DBNull.Value))
                        {
                            double sta1 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_2dsta]);
                            if (sta1 < 0) sta1 = 0;
                            if (sta1 >= Poly2D.Length) sta1 = Poly2D.Length - 0.001;

                            Point3d pt_on_poly = Poly2D.GetPointAtDist(sta1);
                            ds_main.dt_points.Rows[i][col_x] = pt_on_poly.X;
                            ds_main.dt_points.Rows[i][col_y] = pt_on_poly.Y;
                        }
                    }
                }

                Transfer_all_datatables_to_file1();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }



        public void Transfer_all_datatables_to_file1()
        {

            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W2 = null;
            Microsoft.Office.Interop.Excel.Worksheet W3 = null;
            Microsoft.Office.Interop.Excel.Worksheet W4 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

            bool is_opened = false;


            try
            {

                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {


                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            if (Workbook2.FullName == ds_main.config_xls)
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                                {

                                    Workbook1 = Workbook2;
                                    if (Wx.Name == "MatDesc")
                                    {
                                        W1 = Wx;
                                    }
                                    if (Wx.Name == "MatPipe")
                                    {
                                        W2 = Wx;
                                    }
                                    if (Wx.Name == "MatPoints")
                                    {
                                        W3 = Wx;
                                    }
                                    if (Wx.Name == "MatOther")
                                    {
                                        W4 = Wx;
                                    }

                                    is_opened = true;
                                }
                                if (W4 == null)
                                {
                                    W4 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W4.Name = "MatOther";
                                }
                                if (W3 == null)
                                {
                                    W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W3.Name = "MatPoints";
                                }
                                if (W2 == null)
                                {
                                    W2 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W2.Name = "MatPipe";
                                }
                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W1.Name = "MatDesc";
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    if (is_opened == false)
                    {
                        Workbook1 = Excel1.Workbooks.Open(ds_main.config_xls);

                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name == "MatDesc")
                            {
                                W1 = Wx;
                            }
                            if (Wx.Name == "MatPipe")
                            {
                                W2 = Wx;
                            }
                            if (Wx.Name == "MatPoints")
                            {
                                W3 = Wx;
                            }
                            if (Wx.Name == "MatOther")
                            {
                                W4 = Wx;
                            }
                        }
                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W1.Name = "MatDesc";
                        }
                        if (W2 == null)
                        {
                            W2 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W2.Name = "MatPipe";
                        }
                        if (W3 == null)
                        {
                            W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W3.Name = "MatPoints";
                        }

                        if (W4 == null)
                        {
                            W4 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W4.Name = "MatOther";
                        }
                    }

                }


                if (W1 != null || W2 != null || W3 != null || W4 != null)
                {
                    if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
                    {
                        Create_header_material_library(W1, ds_main.client1, ds_main.project1, ds_main.segment1, dt_mat_library);

                        W1.Cells.NumberFormat = "General";
                        int maxRows = dt_mat_library.Rows.Count;
                        int maxCols = dt_mat_library.Columns.Count;
                        W1.Range["A14:G1000"].ClearContents();
                        W1.Range["A14:G1000"].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A14:G" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (dt_mat_library.Rows[i][j] != DBNull.Value && j > 0)// i did not want to save mmid value
                                {
                                    values1[i, j] = dt_mat_library.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }


                    if (ds_main.dt_pipe != null && ds_main.dt_pipe.Rows.Count > 0)
                    {
                        Create_header_material_linear_file(W2, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_pipe);

                        int last_row = nr_max + 14;
                        W2.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_pipe.Rows.Count;
                        int maxCols = ds_main.dt_pipe.Columns.Count;
                        W2.Range["A14:V" + last_row.ToString()].ClearContents();
                        W2.Range["A14:V" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W2.Range["A14:V" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_pipe.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_pipe.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }
                    else
                    {
                        W2.Cells.Clear();
                        W2.Cells.ClearFormats();
                    }



                    if (ds_main.dt_points != null && ds_main.dt_points.Rows.Count > 0)
                    {


                        Create_header_material_points_file(W3, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_points);
                        int last_row = nr_max + 13;
                        W3.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_points.Rows.Count;
                        int maxCols = ds_main.dt_points.Columns.Count;
                        W3.Range["A13:Q" + last_row.ToString()].ClearContents();
                        W3.Range["A13:Q" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W3.Range["A13:Q" + (13 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_points.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_points.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;
                    }
                    else
                    {
                        W3.Cells.Clear();
                        W3.Cells.ClearFormats();
                    }
                    if (ds_main.dt_extra != null && ds_main.dt_extra.Rows.Count > 0)
                    {
                        Create_header_material_linear_file(W4, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_extra, "Material Linear Other");

                        int last_row = nr_max + 14;
                        W4.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_extra.Rows.Count;
                        int maxCols = ds_main.dt_extra.Columns.Count;
                        W4.Range["A14:V" + last_row.ToString()].ClearContents();
                        W4.Range["A14:V" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W4.Range["A14:V" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_extra.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_extra.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }
                    else
                    {
                        W4.Cells.Clear();
                        W4.Cells.ClearFormats();
                    }

                    if (is_opened == true)
                    {
                        Workbook1.Save();
                    }
                    else
                    {
                        Workbook1.Save();
                        Workbook1.Close();
                    }



                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                if (W3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                if (W4 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W4);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }


        public void set_textBox_library_to_red()
        {
            textBox_library.Text = "No library loaded";
            textBox_library.ForeColor = Color.Red;
            textBox_library.Font = font10;
            textBox_library.TextAlign = HorizontalAlignment.Left;
        }

    }
}

