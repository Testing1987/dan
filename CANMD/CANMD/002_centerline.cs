using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;

namespace Alignment_mdi
{
    public partial class Centerline_form : Form
    {
        //Global Variables

        public static Centerline_form tpage_centerline = null;
        string folder_cl = "";
        string cl_excel_name = "centerline.xlsx";
        public static string cl_path = "";
        public static string Col_MMid = "MMID";
        public static string Col_Type = "Type";
        public static string Col_x = "X";
        public static string Col_y = "Y";
        public static string Col_z = "Z";
        public static string Col_2DSta = "2DSta";
        public static string Col_3DSta = "3DSta";
        public static string Col_EqSta = "EqSta";
        public static string Col_BackSta = "BackSta";
        public static string Col_AheadSta = "AheadSta";
        public static string Col_DeflAng = "DeflAng";
        public static string Col_DeflAngDMS = "DeflAngDMS";
        public static string Col_Bearing = "Bearing";
        public static string Col_Distance = "Distance";
        public static string Col_DisplaySta = "DisplaySta";
        public static string Col_DisplayPI = "DisplayPI";
        public static string Col_DisplayProf = "DisplayProf";
        public static string Col_Symbol = "Symbol";
        string col_Type = "Type";
        string col_Item_No = "ItemNo";
        string col_descr = "Descr";
        string col_elbow = "ELBOW";
        string col_mat_elbow = "Material_Elbow";
        string layer_elbow = "_md_elbow";
        string col_MSblock = "MS Block";
        string col_sta = "STA";




        System.Data.DataTable dt_mat_library = null;



        public Centerline_form()
        {
            InitializeComponent();
            tpage_centerline = this;

        }





        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_xl_centerline);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_xl_centerline);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion





        private void button_load_xl_centerline_Click(object sender, EventArgs e)
        {
            string file1 = "";
            try
            {
                using (OpenFileDialog fbd = new OpenFileDialog())
                {
                    fbd.Multiselect = false;
                    fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        file1 = fbd.FileName;

                        Load_centerline(file1);



                        System.Data.DataTable dt_cl = ds_main.dt_centerline;


                        if (dt_cl != null && dt_cl.Rows.Count > 1)
                        {
                            for (int i = 0; i < dt_cl.Rows.Count; ++i)
                            {
                                if (dt_cl.Rows[i][Col_z] != DBNull.Value)
                                {
                                    double z = Convert.ToDouble(dt_cl.Rows[i][Col_z]);
                                    if (z != 0)
                                    {
                                        ds_main.is3D = true;
                                        i = dt_cl.Rows.Count;
                                    }
                                }
                            }
                            for (int i = 1; i < dt_cl.Rows.Count - 1; ++i)
                            {
                                if (dt_cl.Rows[i][Col_x] != DBNull.Value && dt_cl.Rows[i][Col_y] != DBNull.Value)
                                {
                                    if (dt_cl.Rows[i - 1][Col_x] != DBNull.Value && dt_cl.Rows[i - 1][Col_y] != DBNull.Value)
                                    {
                                        if (dt_cl.Rows[i + 1][Col_x] != DBNull.Value && dt_cl.Rows[i + 1][Col_y] != DBNull.Value)
                                        {
                                            double x1 = Convert.ToDouble(dt_cl.Rows[i - 1][Col_x]);
                                            double y1 = Convert.ToDouble(dt_cl.Rows[i - 1][Col_y]);
                                            double x2 = Convert.ToDouble(dt_cl.Rows[i][Col_x]);
                                            double y2 = Convert.ToDouble(dt_cl.Rows[i][Col_y]);
                                            double x3 = Convert.ToDouble(dt_cl.Rows[i + 1][Col_x]);
                                            double y3 = Convert.ToDouble(dt_cl.Rows[i + 1][Col_y]);

                                            double defl = Functions.Get_deflection_angle_rad(x1, y1, x2, y2, x3, y3) * 180 / Math.PI;
                                            dt_cl.Rows[i][Col_DeflAng] = defl;
                                            string defl_dms = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);
                                            dt_cl.Rows[i][Col_DeflAngDMS] = defl_dms;
                                        }

                                    }
                                }
                            }

                        }

                        else
                        {
                            ds_main.dt_centerline = null;
                        }


                    }
                    else
                    {
                        ds_main.dt_centerline = null;

                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            populate_datagridview_cl();
            cl_path = file1;
            set_label_contents(cl_path);
        }


        public void Load_centerline(string file1)
        {


            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                bool is_opened = false;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                        {
                            if (Workbook2.FullName == file1)
                            {
                                Workbook1 = Workbook2;
                                if (Wx.Name == "CenterLine")
                                {
                                    W1 = Wx;
                                }
                                is_opened = true;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }

                bool save_file = false;

                if (is_opened == false)
                {
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = false;
                    Workbook1 = Excel1.Workbooks.Open(file1);

                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                    {
                        if (Wx.Name == "CenterLine")
                        {
                            W1 = Wx;
                        }
                    }
                    if (W1 == null)
                    {
                        W1 = Workbook1.Worksheets[1];
                        W1.Name = "CenterLine";
                        save_file = true;
                    }
                }

                try
                {
                    bool CSF = true;
                    if (ds_main.is_usa == true)
                    {
                        CSF = false;
                    }

                    ds_main.dt_centerline = Build_Data_table_centerline_from_excel(W1, 10, CSF);

                    if (is_opened == false)
                    {
                        if (save_file == true) Workbook1.Save();
                        Workbook1.Close();

                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
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
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable Build_Data_table_centerline_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row, bool CSF)
        {
            System.Data.DataTable Data_table_centerline = Creaza_centerline_datatable_structure();
            if (CSF == true)
            {
                string Col_CSF = "CSF";
                string Col_rr = "Reroute#";
                Data_table_centerline.Columns.Add(Col_CSF, typeof(double));
                Data_table_centerline.Columns.Add(Col_rr, typeof(string));
            }

            string Col1 = "C";
            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;
            ds_main.tpage_main.set_textbox_client_name(W1.Range["B1"].Value2);
            ds_main.tpage_main.set_textbox_project(W1.Range["B2"].Value2);
            ds_main.tpage_main.set_textbox_segment(W1.Range["B3"].Value2);
            ds_main.tpage_main.set_textbox_pipe_diam(W1.Range["D4"].Value2);

            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    Data_table_centerline.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }


            if (is_data == false)
            {
                MessageBox.Show("no data found in the CENTERLINE file");
                return Data_table_centerline;
            }

            int NrR = Data_table_centerline.Rows.Count;
            int NrC = Data_table_centerline.Columns.Count;

            if (is_data == true)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];
                object[,] values = new object[NrR - 1, NrC - 1];
                values = range1.Value2;
                for (int i = 0; i < Data_table_centerline.Rows.Count; ++i)
                {
                    for (int j = 0; j < Data_table_centerline.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        Data_table_centerline.Rows[i][j] = Valoare;
                    }
                }
            }
            return Data_table_centerline;
        }

        public static System.Data.DataTable Creaza_centerline_datatable_structure()
        {
            System.Type type_MMid = typeof(string);
            System.Type type_Type = typeof(string);
            System.Type type_x = typeof(double);
            System.Type type_y = typeof(double);
            System.Type type_z = typeof(double);
            System.Type type_2DSta = typeof(double);
            System.Type type_3DSta = typeof(double);
            System.Type type_EqSta = typeof(double);
            System.Type type_BackSta = typeof(double);
            System.Type type_AheadSta = typeof(double);
            System.Type type_DeflAng = typeof(double);
            System.Type type_DeflAngDMS = typeof(string);
            System.Type type_Bearing = typeof(string);
            System.Type type_Distance = typeof(double);
            System.Type type_DisplaySta = typeof(double);
            System.Type type_DisplayPI = typeof(int);
            System.Type type_DisplayProf = typeof(int);
            System.Type type_Symbol = typeof(string);


            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_Type);
            Lista1.Add(Col_x);
            Lista1.Add(Col_y);
            Lista1.Add(Col_z);
            Lista1.Add(Col_2DSta);
            Lista1.Add(Col_3DSta);
            Lista1.Add(Col_EqSta);
            Lista1.Add(Col_BackSta);
            Lista1.Add(Col_AheadSta);
            Lista1.Add(Col_DeflAng);
            Lista1.Add(Col_DeflAngDMS);
            Lista1.Add(Col_Bearing);
            Lista1.Add(Col_Distance);
            Lista1.Add(Col_DisplaySta);
            Lista1.Add(Col_DisplayPI);
            Lista1.Add(Col_DisplayProf);
            Lista1.Add(Col_Symbol);

            Lista2.Add(type_MMid);
            Lista2.Add(type_Type);
            Lista2.Add(type_x);
            Lista2.Add(type_y);
            Lista2.Add(type_z);
            Lista2.Add(type_2DSta);
            Lista2.Add(type_3DSta);
            Lista2.Add(type_EqSta);
            Lista2.Add(type_BackSta);
            Lista2.Add(type_AheadSta);
            Lista2.Add(type_DeflAng);
            Lista2.Add(type_DeflAngDMS);
            Lista2.Add(type_Bearing);
            Lista2.Add(type_Distance);
            Lista2.Add(type_DisplaySta);
            Lista2.Add(type_DisplayPI);
            Lista2.Add(type_DisplayProf);
            Lista2.Add(type_Symbol);


            System.Data.DataTable dt_cl = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt_cl.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt_cl;
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
            DataGridViewCheckBoxColumn chck_box_col = new DataGridViewCheckBoxColumn();
            chck_box_col.HeaderText = col_name;
            chck_box_col.DataPropertyName = col_name;
            return chck_box_col;
        }

        public DataGridViewComboBoxColumn datagrid_to_datatable_combobox(System.Data.DataTable dt1)
        {
            DataGridViewComboBoxColumn cmbox_col = new DataGridViewComboBoxColumn();

            List<string> col_list = new List<string>();
            if (dt1 != null && dt1.Rows.Count > 0)
            {
                for (int i = 0; i < dt1.Columns.Count; ++i)
                {
                    if (dt1.Rows[i][col_Item_No] != DBNull.Value)
                    {
                        col_list.Add(Convert.ToString(dt1.Rows[i][col_Item_No]));
                    }
                }
            }


            cmbox_col.DataSource = col_list;
            cmbox_col.FlatStyle = FlatStyle.Flat;
            //cmbox_col.DataPropertyName = col_name;
            cmbox_col.HeaderText = "Elbow Material";
            cmbox_col.ValueType = typeof(string);
            cmbox_col.Name = "MatElbow";

            return cmbox_col;
        }

        public void populate_datagridview_cl()
        {
            bool is3D = ds_main.is3D;

            System.Data.DataTable dt_cl = ds_main.dt_centerline;

            if (dt_cl != null && dt_cl.Rows.Count > 0)
            {


                DataGridViewTextBoxColumn dg_col_x = datagrid_to_datatable_textbox(dt_cl, Col_x);
                DataGridViewTextBoxColumn dg_col_y = datagrid_to_datatable_textbox(dt_cl, Col_y);
                DataGridViewTextBoxColumn dg_col_z = datagrid_to_datatable_textbox(dt_cl, Col_z);
                DataGridViewTextBoxColumn dg_col_sta = datagrid_to_datatable_textbox(dt_cl, Col_2DSta);
                DataGridViewTextBoxColumn dg_col_defl = datagrid_to_datatable_textbox(dt_cl, Col_DeflAngDMS);
                DataGridViewCheckBoxColumn dg_col_elbow = null;
                if (dt_cl.Columns.Contains(col_elbow) == true)
                {
                    dg_col_elbow = datagrid_to_datatable_checkbox(dt_cl, col_elbow);
                }
                if (is3D == true)
                {

                    dg_col_sta = datagrid_to_datatable_textbox(dt_cl, Col_3DSta);
                }


                if (dt_cl.Columns.Contains(col_mat_elbow) == false)
                {
                    dt_cl.Columns.Add(col_mat_elbow, typeof(string));
                }

                dataGridView_cl.Columns.Clear();

                if (dt_cl.Columns.Contains(col_elbow) == true)
                {
                    dataGridView_cl.Columns.AddRange(dg_col_x, dg_col_y, dg_col_z, dg_col_sta, dg_col_defl, dg_col_elbow);
                    dataGridView_cl.Columns[5].Name = col_elbow;

                }
                else
                {
                    dataGridView_cl.Columns.AddRange(dg_col_x, dg_col_y, dg_col_z, dg_col_sta, dg_col_defl);
                }

                dataGridView_cl.Columns[0].Name = Col_x;
                dataGridView_cl.Columns[1].Name = Col_y;
                dataGridView_cl.Columns[2].Name = Col_z;
                dataGridView_cl.Columns[3].Name = col_sta;
                dataGridView_cl.Columns[4].Name = Col_DeflAngDMS;





                dataGridView_cl.AutoGenerateColumns = false;
                dataGridView_cl.DataSource = dt_cl;
                dataGridView_cl.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                //DG_Col_OD_Table.FlatStyle = FlatStyle.Flat;
                //DG_Col_OD_Table.Width = 100;
                dataGridView_cl.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_cl.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_cl.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Padding newpadding = new Padding(4, 0, 0, 0);
                dataGridView_cl.ColumnHeadersDefaultCellStyle.Padding = newpadding;
                dataGridView_cl.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_cl.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
                dataGridView_cl.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_cl.EnableHeadersVisualStyles = false;
                dataGridView_cl.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dataGridView_cl.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dataGridView_cl.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dataGridView_cl.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


                dataGridView_cl.Columns[0].Width = 100;
                dataGridView_cl.Columns[1].Width = 100;
                dataGridView_cl.Columns[2].Width = 75;
                dataGridView_cl.Columns[3].Width = 100;

                DataGridViewComboBoxColumn dg_col_mat_elbow = datagrid_to_datatable_combobox(dt_mat_library);



                dataGridView_cl.Columns.Add(dg_col_mat_elbow);


            }
            else
            {
                dataGridView_cl.DataSource = null;
                label_cl.Text = "Centerline not loaded";
                label_cl.ForeColor = Color.Red;



            }


        }

        public void set_label_contents(string file1)
        {
            string continut = "Centerline loaded";
            string continut1 = "Centerline loaded";
            if (file1 != "") continut = file1;
            label_cl.Text = continut;
            label_cl.ForeColor = Color.LightGreen;

        }















        private void dataGridView_cl_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            System.Data.DataTable dt_cl = ds_main.dt_centerline;
            if (dt_cl.Columns.Contains(col_elbow) == true)
            {
                if (e.RowIndex >= 0 && dataGridView_cl.Columns[e.ColumnIndex].Name == col_elbow)
                {

                }
            }
        }

        private void dataGridView_cl_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (e.RowIndex >= 0 && dataGridView_cl.Columns[e.ColumnIndex].Name == col_elbow)
            {
                dataGridView_cl.EndEdit();
            }



        }

        private void button_load_top_from_excel_Click(object sender, EventArgs e)
        {
            string file1 = "";
            try
            {
                using (OpenFileDialog fbd = new OpenFileDialog())
                {
                    fbd.Multiselect = false;
                    fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        file1 = fbd.FileName;

                        Load_top(file1);





                        if (ds_main.dt_top != null && ds_main.dt_top.Rows.Count > 1)
                        {
                            set_label_top_green();
                        }

                        else
                        {
                            ds_main.dt_top = null;
                            set_label_top_red();
                        }


                    }
                    else
                    {
                        ds_main.dt_top = null;
                        set_label_top_red();

                    }
                }
            }
            catch (Exception ex)
            {
                ds_main.dt_top = null;
                set_label_top_red();
                MessageBox.Show(ex.Message);
            }




        }


        public void Load_top(string file1)
        {


            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                bool is_opened = false;
                bool is_top = false;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName == file1)
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                            {

                                Workbook1 = Workbook2;
                                if (Wx.Name == "TOP")
                                {
                                    W1 = Wx;
                                    is_top = true;
                                }
                                is_opened = true;
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
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = false;
                    Workbook1 = Excel1.Workbooks.Open(file1);

                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                    {
                        if (Wx.Name == "TOP")
                        {
                            W1 = Wx;
                            is_top = true;
                        }
                    }
                }

                if (is_top == false)
                {
                    return;
                }

                try
                {


                    ds_main.dt_top = Build_Data_table_top_from_excel(W1, 9);

                    if (is_opened == false)
                    {

                        Workbook1.Close();

                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
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
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable Build_Data_table_top_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {
            System.Data.DataTable dt1 = Creaza_profile_datatable_structure();


            string Col1 = "B";
            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "130000"];
            object[,] values2 = new object[130000, 1];
            values2 = range2.Value2;
            ds_main.tpage_main.set_textbox_client_name(W1.Range["B1"].Value2);
            ds_main.tpage_main.set_textbox_project(W1.Range["B2"].Value2);
            ds_main.tpage_main.set_textbox_segment(W1.Range["B3"].Value2);
            ds_main.tpage_main.set_textbox_pipe_diam(W1.Range["D4"].Value2);

            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    dt1.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }


            if (is_data == false)
            {
                MessageBox.Show("no data found in the PROFILE file");
                return dt1;
            }

            int NrR = dt1.Rows.Count;
            int NrC = dt1.Columns.Count;

            if (is_data == true)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range["B" + Convert.ToString(Start_row) + ":D" + Convert.ToString(NrR + Start_row - 1)];
                object[,] values = new object[NrR - 1, NrC - 1];
                values = range1.Value2;
                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    for (int j = 0; j < dt1.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        dt1.Rows[i][j] = Valoare;
                    }
                }
            }
            //  Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);

            return dt1;
        }

        public static System.Data.DataTable Creaza_profile_datatable_structure()
        {


            System.Type type_z = typeof(double);




            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_3DSta);
            Lista1.Add("EMPTY");
            Lista1.Add(Col_z);



            Lista2.Add(type_z);
            Lista2.Add(type_z);
            Lista2.Add(type_z);



            System.Data.DataTable dt_cl = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt_cl.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt_cl;
        }
        public void set_label_top_green()
        {
            string continut = "TOP loaded";
            label_top.Text = continut;
            label_top.ForeColor = Color.LightGreen;
        }

        public void set_label_top_red()
        {
            string continut = "TOP not loaded";
            label_top.Text = continut;
            label_top.ForeColor = Color.Red;
        }

        private void comboBox_xl_crossing_list_DropDown(object sender, EventArgs e)
        {
            ComboBox combo1 = sender as ComboBox;
            Functions.Load_opened_workbooks_to_combobox(combo1);
            combo1.DropDownWidth = Functions.get_dropdown_width(combo1);
        }



        private void button_load_segment_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    List<string> lista_segments = new List<string>();
                    try
                    {
                        Microsoft.Office.Interop.Excel.Application Excel1 = null;

                        try
                        {
                            Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        }
                        catch (System.Exception ex)
                        {
                            Excel1 = new Microsoft.Office.Interop.Excel.Application();

                        }

                        Excel1.Visible = true;
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fbd.FileName);



                        try
                        {
                            int no_worksheets = Workbook1.Worksheets.Count;
                            foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                            {
                                try
                                {
                                    #region main_worksheet

                                    if (W1.Name == "main_cfg")
                                    {


                                        string b47 = Convert.ToString(W1.Range["B47"].Value);
                                        if (b47 != null && b47.Replace(" ", "") != "")
                                        {
                                            lista_segments = new List<string>();
                                            lista_segments.Add(b47);

                                            string b48 = Convert.ToString(W1.Range["B48"].Value);
                                            if (b48 != null && b48.Replace(" ", "") != "")
                                            {
                                                lista_segments.Add(b48);
                                                string b49 = Convert.ToString(W1.Range["B49"].Value);
                                                if (b49 != null && b49.Replace(" ", "") != "")
                                                {
                                                    lista_segments.Add(b49);
                                                    string b50 = Convert.ToString(W1.Range["B50"].Value);
                                                    if (b50 != null && b50.Replace(" ", "") != "")
                                                    {
                                                        lista_segments.Add(b50);
                                                        string b51 = Convert.ToString(W1.Range["B51"].Value);
                                                        if (b51 != null && b51.Replace(" ", "") != "")
                                                        {
                                                            lista_segments.Add(b51);
                                                            string b52 = Convert.ToString(W1.Range["B52"].Value);
                                                            if (b52 != null && b52.Replace(" ", "") != "")
                                                            {
                                                                lista_segments.Add(b52);
                                                                string b53 = Convert.ToString(W1.Range["B53"].Value);
                                                                if (b53 != null && b53.Replace(" ", "") != "")
                                                                {
                                                                    lista_segments.Add(b53);
                                                                    string b54 = Convert.ToString(W1.Range["B54"].Value);
                                                                    if (b54 != null && b54.Replace(" ", "") != "")
                                                                    {
                                                                        lista_segments.Add(b54);
                                                                        string b55 = Convert.ToString(W1.Range["B55"].Value);
                                                                        if (b55 != null && b55.Replace(" ", "") != "")
                                                                        {
                                                                            lista_segments.Add(b55);
                                                                            string b56 = Convert.ToString(W1.Range["B56"].Value);
                                                                            if (b56 != null && b56.Replace(" ", "") != "")
                                                                            {
                                                                                lista_segments.Add(b56);
                                                                                string b57 = Convert.ToString(W1.Range["B57"].Value);
                                                                                if (b57 != null && b57.Replace(" ", "") != "")
                                                                                {
                                                                                    lista_segments.Add(b57);
                                                                                    string b58 = Convert.ToString(W1.Range["B58"].Value);
                                                                                    if (b58 != null && b58.Replace(" ", "") != "")
                                                                                    {
                                                                                        lista_segments.Add(b58);
                                                                                        string b59 = Convert.ToString(W1.Range["B59"].Value);
                                                                                        if (b59 != null && b59.Replace(" ", "") != "")
                                                                                        {
                                                                                            lista_segments.Add(b59);
                                                                                            string b60 = Convert.ToString(W1.Range["B60"].Value);
                                                                                            if (b60 != null && b60.Replace(" ", "") != "")
                                                                                            {
                                                                                                lista_segments.Add(b60);
                                                                                                string b61 = Convert.ToString(W1.Range["B61"].Value);
                                                                                                if (b61 != null && b61.Replace(" ", "") != "")
                                                                                                {
                                                                                                    lista_segments.Add(b61);
                                                                                                    string b62 = Convert.ToString(W1.Range["B62"].Value);
                                                                                                    if (b62 != null && b62.Replace(" ", "") != "")
                                                                                                    {
                                                                                                        lista_segments.Add(b62);
                                                                                                        string b63 = Convert.ToString(W1.Range["B63"].Value);
                                                                                                        if (b63 != null && b63.Replace(" ", "") != "")
                                                                                                        {
                                                                                                            lista_segments.Add(b63);
                                                                                                            string b64 = Convert.ToString(W1.Range["B64"].Value);
                                                                                                            if (b64 != null)
                                                                                                            {
                                                                                                                lista_segments.Add(b64);
                                                                                                                string b65 = Convert.ToString(W1.Range["B65"].Value);
                                                                                                                if (b65 != null && b65.Replace(" ", "") != "")
                                                                                                                {
                                                                                                                    lista_segments.Add(b65);
                                                                                                                    string b66 = Convert.ToString(W1.Range["B66"].Value);
                                                                                                                    if (b66 != null && b66.Replace(" ", "") != "")
                                                                                                                    {
                                                                                                                        lista_segments.Add(b66);
                                                                                                                        string b67 = Convert.ToString(W1.Range["B67"].Value);
                                                                                                                        if (b67 != null && b67.Replace(" ", "") != "")
                                                                                                                        {
                                                                                                                            lista_segments.Add(b67);
                                                                                                                            string b68 = Convert.ToString(W1.Range["B68"].Value);
                                                                                                                            if (b68 != null && b68.Replace(" ", "") != "")
                                                                                                                            {
                                                                                                                                lista_segments.Add(b68);
                                                                                                                                string b69 = Convert.ToString(W1.Range["B69"].Value);
                                                                                                                                if (b69 != null && b69.Replace(" ", "") != "")
                                                                                                                                {
                                                                                                                                    lista_segments.Add(b69);
                                                                                                                                    string b70 = Convert.ToString(W1.Range["B70"].Value);
                                                                                                                                    if (b70 != null && b70.Replace(" ", "") != "")
                                                                                                                                    {
                                                                                                                                        lista_segments.Add(b70);
                                                                                                                                        string b71 = Convert.ToString(W1.Range["B71"].Value);
                                                                                                                                        if (b71 != null && b71.Replace(" ", "") != "")
                                                                                                                                        {
                                                                                                                                            lista_segments.Add(b71);
                                                                                                                                            string b72 = Convert.ToString(W1.Range["B72"].Value);
                                                                                                                                            if (b72 != null && b72.Replace(" ", "") != "")
                                                                                                                                            {
                                                                                                                                                lista_segments.Add(b72);
                                                                                                                                                string b73 = Convert.ToString(W1.Range["B73"].Value);
                                                                                                                                                if (b73 != null && b73.Replace(" ", "") != "")
                                                                                                                                                {
                                                                                                                                                    lista_segments.Add(b73);
                                                                                                                                                    string b74 = Convert.ToString(W1.Range["B74"].Value);
                                                                                                                                                    if (b74 != null && b74.Replace(" ", "") != "")
                                                                                                                                                    {
                                                                                                                                                        lista_segments.Add(b74);
                                                                                                                                                        string b75 = Convert.ToString(W1.Range["B75"].Value);
                                                                                                                                                        if (b75 != null && b75.Replace(" ", "") != "")
                                                                                                                                                        {
                                                                                                                                                            lista_segments.Add(b75);
                                                                                                                                                            string b76 = Convert.ToString(W1.Range["B76"].Value);
                                                                                                                                                            if (b76 != null && b76.Replace(" ", "") != "")
                                                                                                                                                            {
                                                                                                                                                                lista_segments.Add(b76);
                                                                                                                                                                string b77 = Convert.ToString(W1.Range["B77"].Value);
                                                                                                                                                                if (b77 != null && b77.Replace(" ", "") != "")
                                                                                                                                                                {
                                                                                                                                                                    lista_segments.Add(b77);
                                                                                                                                                                    string b78 = Convert.ToString(W1.Range["B78"].Value);
                                                                                                                                                                    if (b78 != null && b78.Replace(" ", "") != "")
                                                                                                                                                                    {
                                                                                                                                                                        lista_segments.Add(b78);
                                                                                                                                                                        string b79 = Convert.ToString(W1.Range["B79"].Value);
                                                                                                                                                                        if (b79 != null && b79.Replace(" ", "") != "")
                                                                                                                                                                        {
                                                                                                                                                                            lista_segments.Add(b79);
                                                                                                                                                                            string b80 = Convert.ToString(W1.Range["B80"].Value);
                                                                                                                                                                            if (b80 != null && b80.Replace(" ", "") != "")
                                                                                                                                                                            {
                                                                                                                                                                                lista_segments.Add(b80);
                                                                                                                                                                                string b81 = Convert.ToString(W1.Range["B81"].Value);
                                                                                                                                                                                if (b81 != null && b81.Replace(" ", "") != "")
                                                                                                                                                                                {
                                                                                                                                                                                    lista_segments.Add(b81);
                                                                                                                                                                                    string b82 = Convert.ToString(W1.Range["B82"].Value);
                                                                                                                                                                                    if (b82 != null && b82.Replace(" ", "") != "")
                                                                                                                                                                                    {
                                                                                                                                                                                        lista_segments.Add(b82);
                                                                                                                                                                                        string b83 = Convert.ToString(W1.Range["B83"].Value);
                                                                                                                                                                                        if (b83 != null && b83.Replace(" ", "") != "")
                                                                                                                                                                                        {
                                                                                                                                                                                            lista_segments.Add(b83);
                                                                                                                                                                                            string b84 = Convert.ToString(W1.Range["B84"].Value);
                                                                                                                                                                                            if (b84 != null && b84.Replace(" ", "") != "")
                                                                                                                                                                                            {
                                                                                                                                                                                                lista_segments.Add(b84);
                                                                                                                                                                                                string b85 = Convert.ToString(W1.Range["B85"].Value);
                                                                                                                                                                                                if (b85 != null && b85.Replace(" ", "") != "")
                                                                                                                                                                                                {
                                                                                                                                                                                                    lista_segments.Add(b85);
                                                                                                                                                                                                    string b86 = Convert.ToString(W1.Range["B86"].Value);
                                                                                                                                                                                                    if (b86 != null && b86.Replace(" ", "") != "")
                                                                                                                                                                                                    {
                                                                                                                                                                                                        lista_segments.Add(b86);
                                                                                                                                                                                                        string b87 = Convert.ToString(W1.Range["B87"].Value);
                                                                                                                                                                                                        if (b87 != null && b87.Replace(" ", "") != "")
                                                                                                                                                                                                        {
                                                                                                                                                                                                            lista_segments.Add(b87);
                                                                                                                                                                                                            string b88 = Convert.ToString(W1.Range["B88"].Value);
                                                                                                                                                                                                            if (b88 != null && b88.Replace(" ", "") != "")
                                                                                                                                                                                                            {
                                                                                                                                                                                                                lista_segments.Add(b88);
                                                                                                                                                                                                                string b89 = Convert.ToString(W1.Range["B89"].Value);
                                                                                                                                                                                                                if (b89 != null && b89.Replace(" ", "") != "")
                                                                                                                                                                                                                {
                                                                                                                                                                                                                    lista_segments.Add(b89);
                                                                                                                                                                                                                    string b90 = Convert.ToString(W1.Range["B90"].Value);
                                                                                                                                                                                                                    if (b90 != null && b90.Replace(" ", "") != "")
                                                                                                                                                                                                                    {
                                                                                                                                                                                                                        lista_segments.Add(b90);
                                                                                                                                                                                                                        string b91 = Convert.ToString(W1.Range["B91"].Value);
                                                                                                                                                                                                                        if (b91 != null && b91.Replace(" ", "") != "")
                                                                                                                                                                                                                        {
                                                                                                                                                                                                                            lista_segments.Add(b91);
                                                                                                                                                                                                                            string b92 = Convert.ToString(W1.Range["B92"].Value);
                                                                                                                                                                                                                            if (b92 != null && b92.Replace(" ", "") != "")
                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                lista_segments.Add(b92);
                                                                                                                                                                                                                                string b93 = Convert.ToString(W1.Range["B93"].Value);
                                                                                                                                                                                                                                if (b93 != null && b93.Replace(" ", "") != "")
                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                    lista_segments.Add(b93);
                                                                                                                                                                                                                                    string b94 = Convert.ToString(W1.Range["B94"].Value);
                                                                                                                                                                                                                                    if (b94 != null && b94.Replace(" ", "") != "")
                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                        lista_segments.Add(b94);
                                                                                                                                                                                                                                        string b95 = Convert.ToString(W1.Range["B95"].Value);
                                                                                                                                                                                                                                        if (b95 != null && b95.Replace(" ", "") != "")
                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                            lista_segments.Add(b95);
                                                                                                                                                                                                                                            string b96 = Convert.ToString(W1.Range["B96"].Value);
                                                                                                                                                                                                                                            if (b96 != null && b96.Replace(" ", "") != "")
                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                lista_segments.Add(b96);
                                                                                                                                                                                                                                                string b97 = Convert.ToString(W1.Range["B97"].Value);
                                                                                                                                                                                                                                                if (b97 != null && b97.Replace(" ", "") != "")
                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                    lista_segments.Add(b97);
                                                                                                                                                                                                                                                    string b98 = Convert.ToString(W1.Range["B98"].Value);
                                                                                                                                                                                                                                                    if (b98 != null && b98.Replace(" ", "") != "")
                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                        lista_segments.Add(b98);
                                                                                                                                                                                                                                                        string b99 = Convert.ToString(W1.Range["B99"].Value);
                                                                                                                                                                                                                                                        if (b99 != null && b99.Replace(" ", "") != "")
                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                            lista_segments.Add(b99);
                                                                                                                                                                                                                                                            string b100 = Convert.ToString(W1.Range["B100"].Value);
                                                                                                                                                                                                                                                            if (b100 != null && b100.Replace(" ", "") != "")
                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                lista_segments.Add(b100);
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
                                        else
                                        {
                                            lista_segments = null;
                                        }

                                        comboBox_segment_name.Items.Clear();
                                        if (lista_segments != null && lista_segments.Count > 0)
                                        {
                                            try
                                            {
                                                for (int i = 0; i < lista_segments.Count; ++i)
                                                {
                                                    comboBox_segment_name.Items.Add(lista_segments[i]);
                                                }
                                                comboBox_segment_name.SelectedIndex = 0;
                                            }
                                            catch (System.Exception ex)
                                            {
                                                MessageBox.Show(ex.Message);
                                            }
                                        }








                                    }
                                    #endregion

                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(ex.Message);
                                }
                            }




                            Workbook1.Close();

                            if (Excel1.Workbooks.Count == 0)
                            {
                                Excel1.Quit();
                            }
                            else
                            {
                                Excel1.Visible = true;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                            if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);

                        }







                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                }

            }

        }




        private void button_place_labels_on_profile_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt1 = Load_open_tags_from_excel();

            if (dt1 != null && dt1.Rows.Count > 0)

            {



                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Editor1.SetImpliedSelection(Empty_array);
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                try
                {
                    set_enable_false();
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {


                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            double ymin = -1000000;
                            double ymax = 1000000;



                            List<ObjectId> lista_poly = new List<ObjectId>();
                            List<double> lista_start = new List<double>();
                            List<double> lista_end = new List<double>();

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            string Agen_profile_band_V2 = "Agen_profile_band_V2";

                            if (Tables1.IsTableDefined(Agen_profile_band_V2) == true)
                            {
                                foreach (ObjectId id1 in BTrecord)
                                {
                                    Polyline poly_ground = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                    if (poly_ground != null)
                                    {

                                        if (Tables1.IsTableDefined(Agen_profile_band_V2) == true)
                                        {
                                            using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V2])
                                            {

                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                                {
                                                    if (Records1.Count > 0)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {
                                                            double start1 = -123.4;
                                                            double end1 = -123.4;
                                                            string segm1 = "123456";
                                                            for (int i = 0; i < Record1.Count; ++i)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                string Nume_field = Field_def1.Name;
                                                                string Valoare_field = Record1[i].StrValue;

                                                                if (Nume_field.ToLower() == "beginsta")
                                                                {
                                                                    if (Functions.IsNumeric(Valoare_field) == true)
                                                                    {
                                                                        start1 = Convert.ToDouble(Valoare_field);
                                                                    }
                                                                }

                                                                if (Nume_field.ToLower() == "endsta")
                                                                {
                                                                    if (Functions.IsNumeric(Valoare_field) == true)
                                                                    {
                                                                        end1 = Convert.ToDouble(Valoare_field);
                                                                    }
                                                                }
                                                                if (Nume_field.ToLower() == "segment")
                                                                {
                                                                    segm1 = Convert.ToString(Valoare_field);
                                                                }
                                                            }

                                                            string segment2 = comboBox_segment_name.Text;


                                                            if (start1 != -123.4 && end1 != 123.4 && segm1.ToLower() == segment2.ToLower())
                                                            {
                                                                lista_poly.Add(id1);
                                                                lista_start.Add(start1);
                                                                lista_end.Add(end1);
                                                            }

                                                        }
                                                    }
                                                }

                                            }
                                        }

                                    }
                                }
                            }




                            Functions.Creaza_layer("_TAGS", 2, true);

                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {
                                if (dt1.Rows[i]["block_name"] != DBNull.Value && dt1.Rows[i]["sta"] != DBNull.Value)
                                {
                                    string block_name = Convert.ToString(dt1.Rows[i]["block_name"]);
                                    double sta = Convert.ToDouble(dt1.Rows[i]["sta"]);

                                    string display_sta_string = Functions.Get_chainage_from_double(sta, "m", 1);

                                    if (lista_start.Count > 0 && lista_start.Count == lista_end.Count && lista_start.Count == lista_poly.Count)
                                    {
                                        for (int k = 0; k < lista_poly.Count; ++k)
                                        {
                                            if (lista_poly[k] != null && lista_poly[k] != ObjectId.Null)
                                            {
                                                Polyline Poly2d = Trans1.GetObject(lista_poly[k], OpenMode.ForRead) as Polyline;
                                                if (Poly2d != null)
                                                {
                                                    double start1 = lista_start[k];
                                                    double end1 = lista_end[k];
                                                    if (sta >= start1 && sta <= end1)
                                                    {
                                                        for (int n = 0; n < Poly2d.NumberOfVertices - 1; ++n)
                                                        {
                                                            double y = Poly2d.GetPointAtParameter(n).Y;
                                                            if (n == 0)
                                                            {
                                                                ymin = y;
                                                                ymax = y;
                                                            }
                                                            else
                                                            {
                                                                if (ymin > y)
                                                                {
                                                                    ymin = y;
                                                                }
                                                                if (ymax < y)
                                                                {
                                                                    ymax = y;
                                                                }
                                                            }
                                                        }

                                                        double x1 = Poly2d.StartPoint.X + (sta - start1);




                                                        Line line1 = new Line(new Point3d(x1, ymin - 10000, Poly2d.Elevation), new Point3d(x1, ymax + 10000, Poly2d.Elevation));

                                                        Point3dCollection col1 = Functions.Intersect_on_both_operands(Poly2d, line1);


                                                        if (col1.Count == 0)
                                                        {
                                                            col1.Add(new Point3d(x1, Poly2d.GetPoint2dAt(0).Y, Poly2d.Elevation));
                                                        }

                                                        for (int n = 0; n < col1.Count; ++n)
                                                        {
                                                            Point3d inspt = new Point3d();

                                                            inspt = col1[n];


                                                            System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                            System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                            col_atr.Add("sta");
                                                            col_val.Add(display_sta_string);

                                                            for (int j = 2; j < dt1.Columns.Count; ++j)
                                                            {
                                                                string val = "";
                                                                if (dt1.Rows[i][j] != DBNull.Value)
                                                                {
                                                                    val = Convert.ToString(dt1.Rows[i][j]);
                                                                }

                                                                col_atr.Add(dt1.Columns[j].ColumnName);
                                                                col_val.Add(val);
                                                            }

                                                            BlockReference br1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                                                  block_name, inspt, 1 / 1, 0, "_TAGS", col_atr, col_val);


                                                        }

                                                    }
                                                }
                                            }
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
                Editor1.WriteMessage("\nCommand:");





                set_enable_true();
            }
            else
            {
                MessageBox.Show("no  data found");
            }
        }

        public System.Data.DataTable Load_open_tags_from_excel()
        {
            System.Data.DataTable dt2 = new System.Data.DataTable();

            dt2.Columns.Add("sta", typeof(double));
            dt2.Columns.Add("block_name", typeof(string));
            dt2.Columns.Add("xing_id", typeof(string));
            dt2.Columns.Add("owner", typeof(string));
            dt2.Columns.Add("descr", typeof(string));
            dt2.Columns.Add("xing_info", typeof(string));
            dt2.Columns.Add("depth_height", typeof(string));
            dt2.Columns.Add("substance_grade", typeof(string));
            dt2.Columns.Add("size", typeof(string));
            dt2.Columns.Add("depth_of_cover", typeof(string));
            dt2.Columns.Add("xing_method", typeof(string));



            int start1 = -1;
            int end1 = -1;
            if (Functions.IsNumeric(textBox_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_start.Text);
            }
            if (Functions.IsNumeric(textBox_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_end.Text);
            }
            if (start1 == -1)
            {
                MessageBox.Show("start is not numeric");
                return null;
            }
            if (end1 == -1)
            {
                MessageBox.Show("end is not numeric");
                return null;
            }
            try
            {

                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {

                        string file_name = Workbook2.FullName;
                        file_name = System.IO.Path.GetFileName(file_name);
                        if (file_name.ToLower() == comboBox_xl_crossing_list.Text.ToLower())
                        {
                            Workbook1 = Workbook2;

                            W1 = Workbook1.Worksheets[1];



                        }

                    }


                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }
                if (W1 != null)
                {
                    try
                    {



                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add("sta");
                        lista_col.Add("block_name");
                        lista_col.Add("xing_id");
                        lista_col.Add("owner");
                        lista_col.Add("descr");
                        lista_col.Add("xing_info");
                        lista_col.Add("depth_height");
                        lista_col.Add("substance_grade");
                        lista_col.Add("size");
                        lista_col.Add("depth_of_cover");
                        lista_col.Add("xing_method");

                        lista_colxl.Add(textBox_sta.Text);
                        lista_colxl.Add(textBox_block_name.Text);
                        lista_colxl.Add(textBox_xing_id.Text);
                        lista_colxl.Add(textBox_owner.Text);
                        lista_colxl.Add(textBox_descr.Text);
                        lista_colxl.Add(textBox_xing_info.Text);
                        lista_colxl.Add(textBox_depth_height.Text);
                        lista_colxl.Add(textBox_substance_grade.Text);
                        lista_colxl.Add(textBox_size.Text);
                        lista_colxl.Add(textBox_depth_of_cover.Text);
                        lista_colxl.Add(textBox_xing_method.Text);


                        dt2 = Functions.build_dt_from_excel(dt2, W1, start1, end1, lista_col, lista_colxl);



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



            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return dt2;
        }


    }
}

