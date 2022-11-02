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

namespace Alignment_mdi
{
    public partial class Centerline_form : Form
    {
        //Global Variables

        Centerline_form tpage_centerline = null;
        string folder_cl = "";
        string cl_excel_name = "centerline.xlsx";
        string Col_MMid = "MMID";
        string col_x = "X";
        string col_y = "Y";
        string col_z = "Z";
        string col_sta2d = "2DSta";
        string col_sta3d = "3DSta";
        string col_eqsta = "EqSta";
        string Col_BackSta = "BackSta";
        string Col_AheadSta = "AheadSta";
        string Col_DeflAng = "DeflAng";
        string Col_DeflAngDMS = "DeflAngDMS";
        string Col_Bearing = "Bearing";
        string Col_Distance = "Distance";
        string Col_DisplaySta = "DisplaySta";
        string Col_DisplayPI = "DisplayPI";
        string Col_DisplayProf = "DisplayProf";
        string Col_Symbol = "Symbol";
        string col_type = "Type";
        string col_Item_No = "ItemNo";
        string col_descr = "Descr";
        string col_elbow = "ELBOW";
        string col_mat_elbow = "Elbow Item No";
        string layer_elbow = "_md_elbow";
        string col_MSblock = "MS Block";
        string col_sta = "STA";
        string col_block = "BLOCK";

        string col_cat = "Category";

        string col_mmid = "MMID";
        string col_item_no = "ItemNo";
        string col_sta_eq = "EqSta";
        string col_altdesc = "AltDesc";
        string col_symbol = "Symbol";
        string col_atr1 = "ATR1";
        string col_atr2 = "ATR2";
        string col_atr3 = "ATR3";
        string col_atr4 = "ATR4";
        string col_visibility = "Visibility";

        System.Data.DataTable dt_mat_library = null;
        public System.Data.DataTable dt_cl_display = null;
        public System.Data.DataTable dt_cl_display_filtered = null;

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
            lista_butoane.Add(button_load_dwg_cl);
            lista_butoane.Add(button_save_cl);
            lista_butoane.Add(button_filter);
            lista_butoane.Add(button_clear_elbows);
            lista_butoane.Add(button_transfer_to_mat);
            lista_butoane.Add(button_assign_elbows);
            lista_butoane.Add(button_add_stationing);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_xl_centerline);
            lista_butoane.Add(button_load_dwg_cl);
            lista_butoane.Add(button_save_cl);
            lista_butoane.Add(button_filter);
            lista_butoane.Add(button_clear_elbows);
            lista_butoane.Add(button_transfer_to_mat);
            lista_butoane.Add(button_assign_elbows);
            lista_butoane.Add(button_add_stationing);

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
                                if (dt_cl.Rows[i][col_z] != DBNull.Value)
                                {
                                    double z = Convert.ToDouble(dt_cl.Rows[i][col_z]);
                                    if (z != 0)
                                    {
                                        ds_main.is3D = true;
                                        i = dt_cl.Rows.Count;
                                    }
                                }
                            }
                            for (int i = 1; i < dt_cl.Rows.Count - 1; ++i)
                            {
                                if (dt_cl.Rows[i][col_x] != DBNull.Value && dt_cl.Rows[i][col_y] != DBNull.Value)
                                {
                                    if (dt_cl.Rows[i - 1][col_x] != DBNull.Value && dt_cl.Rows[i - 1][col_y] != DBNull.Value)
                                    {
                                        if (dt_cl.Rows[i + 1][col_x] != DBNull.Value && dt_cl.Rows[i + 1][col_y] != DBNull.Value)
                                        {
                                            double x1 = Convert.ToDouble(dt_cl.Rows[i - 1][col_x]);
                                            double y1 = Convert.ToDouble(dt_cl.Rows[i - 1][col_y]);
                                            double x2 = Convert.ToDouble(dt_cl.Rows[i][col_x]);
                                            double y2 = Convert.ToDouble(dt_cl.Rows[i][col_y]);
                                            double x3 = Convert.ToDouble(dt_cl.Rows[i + 1][col_x]);
                                            double y3 = Convert.ToDouble(dt_cl.Rows[i + 1][col_y]);

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

            dt_cl_display = null;
            dt_cl_display_filtered = null;
            populate_datagridview_cl();
            ds_main.centerline_xls = file1;

            input_fisier_cl_in_config();

            set_label_contents(ds_main.centerline_xls);
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


                    ds_main.dt_centerline = Build_Data_table_centerline_from_excel(W1, 10);

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

        public System.Data.DataTable Build_Data_table_centerline_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {
            System.Data.DataTable Data_table_centerline = Creaza_centerline_datatable_structure();
            string Col1 = "C";
            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;


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

        public System.Data.DataTable Creaza_centerline_datatable_structure()
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
            Lista1.Add(col_type);
            Lista1.Add(col_x);
            Lista1.Add(col_y);
            Lista1.Add(col_z);
            Lista1.Add(col_sta2d);
            Lista1.Add(col_sta3d);
            Lista1.Add(col_eqsta);
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





        private void button_load_dwg_cl_Click(object sender, EventArgs e)
        {

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
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult rez_cl;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_cl;
                        prompt_cl = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        prompt_cl.SetRejectMessage("\nSelect a polyline!");
                        prompt_cl.AllowNone = true;
                        prompt_cl.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        prompt_cl.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        rez_cl = ThisDrawing.Editor.GetEntity(prompt_cl);
                        if (rez_cl.Status != PromptStatus.OK)
                        {
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            ds_main.dt_centerline = null;
                            populate_datagridview_cl();
                            set_label_contents("");
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline poly_cl = Trans1.GetObject(rez_cl.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline3d poly_3d_cl = Trans1.GetObject(rez_cl.ObjectId, OpenMode.ForRead) as Polyline3d;

                        if (poly_cl != null)
                        {
                            ds_main.dt_centerline = Creaza_centerline_datatable_structure();
                            ds_main.is3D = false;
                            for (int i = 0; i < poly_cl.NumberOfVertices; ++i)
                            {
                                double x2 = poly_cl.GetPointAtParameter(i).X;
                                double y2 = poly_cl.GetPointAtParameter(i).Y;

                                ds_main.dt_centerline.Rows.Add();
                                ds_main.dt_centerline.Rows[i][col_x] = x2;
                                ds_main.dt_centerline.Rows[i][col_y] = y2;
                                ds_main.dt_centerline.Rows[i][col_z] = 0;
                                ds_main.dt_centerline.Rows[i][col_sta2d] = poly_cl.GetDistanceAtParameter(i);
                                if (i > 0 && i < poly_cl.NumberOfVertices - 1)
                                {
                                    double x1 = poly_cl.GetPointAtParameter(i - 1).X;
                                    double y1 = poly_cl.GetPointAtParameter(i - 1).Y;
                                    double x3 = poly_cl.GetPointAtParameter(i + 1).X;
                                    double y3 = poly_cl.GetPointAtParameter(i + 1).Y;

                                    string defl_dms = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);
                                    double defl_deg = 180 * Functions.Get_deflection_angle_rad(x1, y1, x2, y2, x3, y3) / Math.PI;
                                    ds_main.dt_centerline.Rows[i][Col_DeflAngDMS] = defl_dms;
                                    ds_main.dt_centerline.Rows[i][Col_DeflAng] = defl_deg;

                                }


                            }
                        }

                        if (poly_3d_cl != null)
                        {
                            ds_main.dt_centerline = Creaza_centerline_datatable_structure();
                            ds_main.is3D = true;

                            Polyline poly2d = new Polyline();

                            for (int i = 0; i <= poly_3d_cl.EndParam; ++i)
                            {
                                double x2 = poly_3d_cl.GetPointAtParameter(i).X;
                                double y2 = poly_3d_cl.GetPointAtParameter(i).Y;
                                double z2 = poly_3d_cl.GetPointAtParameter(i).Z;

                                poly2d.AddVertexAt(i, new Point2d(x2, y2), 0, 0, 0);

                                ds_main.dt_centerline.Rows.Add();
                                ds_main.dt_centerline.Rows[i][col_x] = x2;
                                ds_main.dt_centerline.Rows[i][col_y] = y2;
                                ds_main.dt_centerline.Rows[i][col_z] = z2;
                                ds_main.dt_centerline.Rows[i][col_sta3d] = poly_3d_cl.GetDistanceAtParameter(i);
                                ds_main.dt_centerline.Rows[i][col_sta2d] = poly2d.Length;

                                if (i > 0 && i < poly_3d_cl.EndParam)
                                {
                                    double x1 = poly_3d_cl.GetPointAtParameter(i - 1).X;
                                    double y1 = poly_3d_cl.GetPointAtParameter(i - 1).Y;
                                    double x3 = poly_3d_cl.GetPointAtParameter(i + 1).X;
                                    double y3 = poly_3d_cl.GetPointAtParameter(i + 1).Y;

                                    string defl_dms = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);
                                    double defl_deg = 180 * Functions.Get_deflection_angle_rad(x1, y1, x2, y2, x3, y3) / Math.PI;
                                    ds_main.dt_centerline.Rows[i][Col_DeflAngDMS] = defl_dms;
                                    ds_main.dt_centerline.Rows[i][Col_DeflAng] = defl_deg;

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
            this.MdiParent.WindowState = FormWindowState.Normal;

            populate_datagridview_cl();
            dt_cl_display = null;
            set_label_contents("");
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }




        public void set_label_contents(string file1)
        {
            string continut = "Centerline loaded";
            string continut1 = "Centerline loaded";
            if (file1 != "") continut = file1;
            label_cl.Text = continut;
            label_cl.ForeColor = Color.LightGreen;
            ds_main.tpage_mat_design.populate_textbox_cl(continut1);
        }

        public void set_combobox_elbow(string mat1)
        {
            if (comboBox_elbow.Items.Count > 0)
            {
                if (comboBox_elbow.Items.Contains(mat1) == true)
                {
                    comboBox_elbow.SelectedIndex = comboBox_elbow.Items.IndexOf(mat1);
                }
            }
        }


        private void button_save_cl_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt_cl = ds_main.dt_centerline;

            if (dt_cl == null && dt_cl.Rows.Count == 0)
            {
                return;
            }

            if (System.IO.File.Exists(ds_main.config_xls) == true)
            {
                string filename = System.IO.Path.GetFileName(ds_main.config_xls);
                folder_cl = ds_main.config_xls.Replace(filename, "");
            }
            else
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Specify centerline.xlsx folder";

                    fbd.SelectedPath = folder_cl;

                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        folder_cl = fbd.SelectedPath.ToString();
                    }
                }
            }

            if (folder_cl == "") return;

            if (folder_cl.Substring(folder_cl.Length - 1, 1) != "\\")
            {
                folder_cl = folder_cl + "\\";
            }

            string fisier_cl = folder_cl + cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == true)
            {
                if (MessageBox.Show("all data from centerline.xls will be overwriten... \r\nDo you want to continue?", "Material Design", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                {
                    return;
                }
            }

            ds_main.centerline_xls = fisier_cl;
            Transfer_dt_cl_to_file1(dt_cl, fisier_cl);
            label_cl.Text = fisier_cl;
            label_cl.ForeColor = Color.LightGreen;

            string continut1 = "Centerline loaded";
            ds_main. tpage_mat_design.populate_textbox_cl(continut1);
        }

        public void Transfer_dt_cl_to_file1(System.Data.DataTable dt1, string fisier_cl)
        {

            string client1 = ds_main.tpage_main.get_textbox_client_name();
            string project1 = ds_main.tpage_main.get_textbox_project();
            string segment1 = ds_main.tpage_main.get_textbox_segment();
            string version1 = "";
            string diam1 = ds_main.tpage_main.get_textbox_pipe_diam();

            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel._Worksheet W_cfg = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook_cfg = null;

            bool is_opened = false;
            bool save_as = false;
            bool save_and_close = false;

            try
            {
                if (System.IO.File.Exists(fisier_cl) == true)
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {

                            if (Workbook2.FullName == fisier_cl)
                            {
                                Workbook1 = Workbook2;
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                                {

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

                    if (is_opened == false)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = false;
                        Workbook1 = Excel1.Workbooks.Open(fisier_cl);
                        save_and_close = true;
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name == "CenterLine")
                            {
                                W1 = Wx;
                            }
                        }

                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W1.Name = "CenterLine";
                        }
                    }
                    if (W1 == null)
                    {
                        W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W1.Name = "CenterLine";
                    }
                }
                else
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    Workbook1 = Excel1.Workbooks.Add();
                    W1 = Workbook1.Worksheets[1];
                    W1.Name = "CenterLine";
                    save_as = true;
                }
                if (dt1 != null && W1 != null)
                {
                    if (dt1.Columns.Contains(col_elbow) == true)
                    {
                        dt1.Columns.Remove(col_elbow);
                    }

                    if (dt1.Rows.Count > 0)
                    {
                        Create_header_centerline_file(W1, client1, project1, segment1, version1, diam1, dt1);
                        W1.Cells.NumberFormat = "General";
                        int maxRows = dt1.Rows.Count;
                        int maxCols = dt1.Columns.Count;
                        W1.Range["A10:R150000"].ClearContents();
                        W1.Range["A10:R150000"].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A10:R" + (10 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (dt1.Rows[i][j] != DBNull.Value)//&& j > 0 i did not want to save mmid value
                                {
                                    values1[i, j] = dt1.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;
                    }

                    if (save_as == true)
                    {
                        Workbook1.SaveAs(fisier_cl);
                        Workbook1.Close();
                    }
                    if (is_opened == true)
                    {
                        Workbook1.Save();
                    }
                    if (save_and_close == true)
                    {
                        Workbook1.Save();
                        Workbook1.Close();
                    }



                    if (System.IO.File.Exists(ds_main.config_xls) == true)
                    {

                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            if (Workbook2.FullName == ds_main.config_xls)
                            {
                                Workbook_cfg = Workbook2;
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                                {
                                    if (Wx.Name == "MDConfig")
                                    {
                                        W_cfg = Wx;
                                        W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                                    }
                                }
                                if (W_cfg == null)
                                {
                                    W_cfg = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W_cfg.Name = "MDConfig";
                                    W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                                }

                                Workbook_cfg.Save();

                            }
                        }

                        if (Workbook_cfg == null)
                        {
                            Workbook_cfg = Excel1.Workbooks.Open(ds_main.config_xls);
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                            {
                                if (Wx.Name == "MDConfig")
                                {
                                    W_cfg = Wx;
                                    W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                                }
                            }
                            if (W_cfg == null)
                            {
                                W_cfg = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                W_cfg.Name = "MDConfig";
                                W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                            }
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
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (W_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cfg);
                if (Workbook_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_cfg);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }
        }


        public void input_fisier_cl_in_config()
        {

            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel._Worksheet W_cfg = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook_cfg = null;

            try
            {

                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName == ds_main.config_xls)
                        {
                            Workbook_cfg = Workbook2;
                         
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                            {
                                if (Wx.Name == "MDConfig")
                                {
                                    W_cfg = Wx;
                                    W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                                }
                            }
                            if (W_cfg == null)
                            {
                                W_cfg = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                W_cfg.Name = "MDConfig";
                                W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                            }
                            Workbook_cfg.Save();
                        }
                    }

                    if (Workbook_cfg == null)
                    {
                        Workbook_cfg = Excel1.Workbooks.Open(ds_main.config_xls);
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                        {
                            if (Wx.Name == "MDConfig")
                            {
                                W_cfg = Wx;
                                W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                            }
                        }
                        if (W_cfg == null)
                        {
                            W_cfg = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W_cfg.Name = "MDConfig";
                            W_cfg.Range["B5"].Value2 = ds_main.centerline_xls;
                        }
                        Workbook_cfg.Save();
                        Workbook_cfg.Close();
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
                if (W_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cfg);
                if (Workbook_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_cfg);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }
        }


        public static void Create_header_centerline_file(Microsoft.Office.Interop.Excel.Worksheet W1, string Client, string Project, string Segment, string Version, string diam, System.Data.DataTable dt1)
        {
            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:D8"];
            object[,] valuesH = new object[8, 4];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";
            valuesH[3, 1] = Version;
            valuesH[3, 2] = "PIPE DIAMETER";
            valuesH[3, 3] = diam;
            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at: " + DateTime.Now.TimeOfDay;
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;
            valuesH[6, 0] = "Do not manually edit any of the table information below.";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:D6"];

            Functions.Color_border_range_inside(range1, 46);

            W1.Range["A6"].Font.Bold = true;

            range1 = W1.Range["A7:R7"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6); //yellow

            range1 = W1.Range["A8:R8"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 43); //green 3); //red



            range1 = W1.Range["E1:R6"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "Centerline";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Functions.Color_border_range_outside(range1, 0);




            range1 = W1.Range["A9:R9"];
            object[,] values_dt = new object[1, dt1.Columns.Count];
            if (dt1 != null && dt1.Columns.Count > 0)
            {
                for (int i = 0; i < dt1.Columns.Count; ++i)
                {
                    values_dt[0, i] = dt1.Columns[i].ColumnName;
                }
                range1.Value2 = values_dt;
                Functions.Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }
            W1.Range["A:R"].ColumnWidth = 14;
        }

        public void add_elbows_mat_to_combobox(System.Data.DataTable dt_mat_lib)
        {
            int index1 = comboBox_elbow.Items.IndexOf(comboBox_elbow.Text);
            dt_mat_library = dt_mat_lib;
            comboBox_elbow.Items.Clear();
            comboBox_description.Items.Clear();
            for (int i = 0; i < dt_mat_lib.Rows.Count; ++i)
            {
                if (Convert.ToString(dt_mat_lib.Rows[i][col_cat]).ToUpper() == "ELL" && Convert.ToString(dt_mat_lib.Rows[i][col_type]).ToUpper() == "POINT")
                {
                    comboBox_elbow.Items.Add(dt_mat_lib.Rows[i][col_Item_No]);
                    comboBox_description.Items.Add(dt_mat_lib.Rows[i][col_descr]);
                }
            }

            if (comboBox_elbow.Items.Count > index1)
            {
                comboBox_elbow.SelectedIndex = index1;
                comboBox_description.SelectedIndex = index1;
            }

        }






        private void comboBox_elbow_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
            {
                string mat1 = comboBox_elbow.Text;

                if (mat1 == "")
                {
                    if (comboBox_description.Items.Count > 0) comboBox_description.SelectedItem = 0;
                }
                else
                {
                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        if (dt_mat_library.Rows[i][col_descr] != DBNull.Value && dt_mat_library.Rows[i][col_Item_No] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[i][col_Item_No]) == mat1)
                        {
                            string new_descr = Convert.ToString(dt_mat_library.Rows[i][col_descr]);
                            if (comboBox_description.Items.Contains(new_descr) == true)
                            {
                                comboBox_description.SelectedIndex = comboBox_description.Items.IndexOf(new_descr);
                            }

                        }
                    }
                }
            }
        }

        private void comboBox_description_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dt_mat_library != null && dt_mat_library.Rows.Count > 0)
            {
                string descr1 = comboBox_description.Text;

                if (descr1 == "")
                {
                    if (comboBox_elbow.Items.Count > 0) comboBox_elbow.SelectedItem = 0;
                }
                else
                {
                    for (int i = 0; i < dt_mat_library.Rows.Count; ++i)
                    {
                        if (dt_mat_library.Rows[i][col_Item_No] != DBNull.Value && dt_mat_library.Rows[i][col_descr] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[i][col_descr]) == descr1)
                        {
                            string new_mat = Convert.ToString(dt_mat_library.Rows[i][col_Item_No]);
                            if (comboBox_elbow.Items.Contains(new_mat) == true)
                            {
                                comboBox_elbow.SelectedIndex = comboBox_elbow.Items.IndexOf(new_mat);
                            }

                        }
                    }
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


        public void populate_datagridview_cl()
        {
            bool is3D = ds_main.is3D;

            System.Data.DataTable dt_cl = ds_main.dt_centerline;

            if (dt_cl_display_filtered == null || dt_cl_display_filtered.Rows.Count == 0)
            {
                dt_cl_display = new System.Data.DataTable();
                dt_cl_display.Columns.Add(col_x, typeof(double));
                dt_cl_display.Columns.Add(col_y, typeof(double));
                dt_cl_display.Columns.Add(col_z, typeof(double));
                dt_cl_display.Columns.Add(col_sta2d, typeof(double));
                dt_cl_display.Columns.Add(col_sta3d, typeof(double));
                dt_cl_display.Columns.Add(Col_DeflAng, typeof(double));
                dt_cl_display.Columns.Add(Col_DeflAngDMS, typeof(string));
                dt_cl_display.Columns.Add(col_elbow, typeof(bool));
                dt_cl_display.Columns.Add(col_mat_elbow, typeof(string));
            }





            if (dt_cl != null && dt_cl.Rows.Count > 0)
            {

                if (dt_cl_display_filtered == null || dt_cl_display_filtered.Rows.Count == 0)
                {
                    for (int i = 0; i < dt_cl.Rows.Count; ++i)
                    {
                        dt_cl_display.Rows.Add();
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][col_x] = dt_cl.Rows[i][col_x];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][col_y] = dt_cl.Rows[i][col_y];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][col_z] = dt_cl.Rows[i][col_z];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][col_z] = dt_cl.Rows[i][col_z];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][Col_DeflAngDMS] = dt_cl.Rows[i][Col_DeflAngDMS];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][Col_DeflAng] = dt_cl.Rows[i][Col_DeflAng];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][col_sta2d] = dt_cl.Rows[i][col_sta2d];
                        dt_cl_display.Rows[dt_cl_display.Rows.Count - 1][col_sta3d] = dt_cl.Rows[i][col_sta3d];

                    }
                }
                else
                {
                    for (int i = 0; i < dt_cl_display.Rows.Count; ++i)
                    {
                        double sta1 = -1.234;
                        if (dt_cl_display.Rows[i][col_sta2d] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt_cl_display.Rows[i][col_sta2d]);
                        }
                        if (dt_cl_display.Rows[i][col_sta3d] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt_cl_display.Rows[i][col_sta3d]);
                        }


                        for (int j = 0; j < dt_cl_display_filtered.Rows.Count; ++j)
                        {
                            double sta2 = -12.345;
                            if (dt_cl_display_filtered.Rows[j][col_sta2d] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(dt_cl_display_filtered.Rows[j][col_sta2d]);
                            }
                            if (dt_cl_display_filtered.Rows[j][col_sta3d] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(dt_cl_display_filtered.Rows[j][col_sta3d]);
                            }
                            if (Math.Round(sta1, 2) == Math.Round(sta2, 2))
                            {
                                if (dt_cl_display_filtered.Rows[j][col_elbow] != DBNull.Value)
                                {
                                    bool is_elbow = false;
                                    is_elbow = Convert.ToBoolean(dt_cl_display_filtered.Rows[j][col_elbow]);
                                    if (is_elbow == true)
                                    {
                                        dt_cl_display.Rows[i][col_elbow] = dt_cl_display_filtered.Rows[j][col_elbow];
                                        dt_cl_display.Rows[i][col_mat_elbow] = dt_cl_display_filtered.Rows[j][col_mat_elbow];
                                    }
                                }
                            }
                        }
                    }
                }
                string ct1 = "ELL_POINT";

                if (ds_main.tpage_mat_design.ct_list != null)
                {
                    int idx1 = ds_main.tpage_mat_design.ct_list.IndexOf(ct1);
                    if (ds_main.tpage_mat_design.dt_ct[idx1] != null && ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count; ++i)
                        {
                            double sta2d = -1.23;
                            double sta3d = -1.23;
                            if (ds_main.tpage_mat_design.dt_ct[idx1].Rows[i][col_sta2d] != DBNull.Value)
                            {
                                sta2d = Convert.ToDouble(ds_main.tpage_mat_design.dt_ct[idx1].Rows[i][col_sta2d]);
                            }
                            if (ds_main.tpage_mat_design.dt_ct[idx1].Rows[i][col_sta3d] != DBNull.Value)
                            {
                                sta3d = Convert.ToDouble(ds_main.tpage_mat_design.dt_ct[idx1].Rows[i][col_sta3d]);
                            }

                            for (int j = 0; j < dt_cl_display.Rows.Count; ++j)
                            {
                                double sta1 = -4.56;
                                double sta2 = -4.56;
                                if (dt_cl_display.Rows[j][col_sta2d] != DBNull.Value)
                                {
                                    sta1 = Convert.ToDouble(dt_cl_display.Rows[j][col_sta2d]);
                                }
                                if (dt_cl_display.Rows[j][col_sta3d] != DBNull.Value)
                                {
                                    sta2 = Convert.ToDouble(dt_cl_display.Rows[j][col_sta3d]);
                                }

                                if (Math.Round(sta1, 2) == Math.Round(sta2d, 2) || Math.Round(sta2, 2) == Math.Round(sta3d, 2))
                                {
                                    dt_cl_display.Rows[j][col_elbow] = true;
                                    dt_cl_display.Rows[j][col_mat_elbow] = ds_main.tpage_mat_design.dt_ct[idx1].Rows[i][col_item_no];
                                    j = dt_cl_display.Rows.Count;
                                }
                            }
                        }
                    }
                }


                DataGridViewTextBoxColumn dg_col_x = Functions.datagrid_textbox_column(col_x);
                DataGridViewTextBoxColumn dg_col_y = Functions.datagrid_textbox_column(col_y);
                DataGridViewTextBoxColumn dg_col_z = Functions.datagrid_textbox_column(col_z);
                DataGridViewTextBoxColumn dg_col_sta = Functions.datagrid_textbox_column(col_sta2d);
                DataGridViewTextBoxColumn dg_col_defl = Functions.datagrid_textbox_column(Col_DeflAngDMS);
                DataGridViewCheckBoxColumn dg_col_elbow = Functions.datagrid_checkbox_column(col_elbow);
                DataGridViewTextBoxColumn dg_col_mat_elbow = Functions.datagrid_textbox_column(col_mat_elbow);

                if (is3D == true)
                {
                    dg_col_sta = Functions.datagrid_textbox_column(col_sta3d);
                }

                dt_cl_display_filtered = new System.Data.DataTable();
                dt_cl_display_filtered = dt_cl_display.Copy();



                string mat1 = comboBox_elbow.Text;
                string ang_string = textBox_min_angle.Text;
                if (mat1 != "")
                {


                    if (dt_cl_display_filtered != null && dt_cl_display_filtered.Rows.Count > 1)
                    {
                        if (Functions.IsNumeric(ang_string) == true && Math.Abs(Convert.ToDouble(ang_string)) >= 0)
                        {

                            double min_angle = Math.Abs(Convert.ToDouble(ang_string));
                            for (int i = dt_cl_display_filtered.Rows.Count - 1; i >= 0; --i)
                            {
                                if (dt_cl_display_filtered.Rows[i][Col_DeflAng] != DBNull.Value)
                                {
                                    if (Convert.ToDouble(dt_cl_display_filtered.Rows[i][Col_DeflAng]) < min_angle)
                                    {
                                        dt_cl_display_filtered.Rows[i].Delete();
                                    }
                                }
                                else
                                {
                                    dt_cl_display_filtered.Rows[i].Delete();
                                }

                            }


                        }
                    }
                }



                dataGridView_cl.Columns.Clear();
                dataGridView_cl.Columns.AddRange(dg_col_x, dg_col_y, dg_col_z, dg_col_sta, dg_col_defl, dg_col_elbow, dg_col_mat_elbow);



                dataGridView_cl.AutoGenerateColumns = false;
                dataGridView_cl.DataSource = dt_cl_display_filtered;

                dataGridView_cl.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

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

                dataGridView_cl.Columns[6].Name = col_mat_elbow;

            }
            else
            {
                dataGridView_cl.DataSource = null;
                label_cl.Text = "Centerline not loaded";
                label_cl.ForeColor = Color.Red;



                ds_main.tpage_mat_design.populate_textbox_cl("");
            }


        }

        private void button_filter_Click(object sender, EventArgs e)
        {
            populate_datagridview_cl();
        }

        private void button_assign_elbows_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                string mat1 = comboBox_elbow.Text;
                if (mat1 != "")
                {
                    if (dt_cl_display_filtered != null && dt_cl_display_filtered.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_cl_display_filtered.Rows.Count; ++i)
                        {
                            bool is_elbow = false;
                            if (dt_cl_display_filtered.Rows[i][col_elbow] != DBNull.Value)
                            {
                                is_elbow = Convert.ToBoolean(dt_cl_display_filtered.Rows[i][col_elbow]);
                            }
                            //if (is_elbow == false)

                            dt_cl_display_filtered.Rows[i][col_elbow] = true;
                            dt_cl_display_filtered.Rows[i][col_mat_elbow] = mat1;

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



        private void dataGridView_cl_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && dataGridView_cl.Columns[e.ColumnIndex].Name == col_elbow)
            {
                //insert_elbows();
            }
        }

        private void dataGridView_cl_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > -1)
            {
                if (dataGridView_cl.Columns[e.ColumnIndex].Name == col_mat_elbow)
                {
                    DataGridViewComboBoxCell cbox = new DataGridViewComboBoxCell();
                    cbox.Style.BackColor = Color.FromArgb(51, 51, 55);
                    cbox.Style.ForeColor = Color.White;
                    cbox.Style.SelectionBackColor = Color.FromArgb(51, 51, 55);
                    cbox.Style.SelectionForeColor = Color.White;
                    cbox.Style.Padding = new Padding(4, 0, 0, 0);
                    dataGridView_cl[e.ColumnIndex, e.RowIndex] = cbox;
                    if (comboBox_elbow.Items.Count > 0)
                    {
                        cbox.DataSource = comboBox_elbow.Items;
                    }

                }
            }


        }


        private void button_transfer_to_mat_Click(object sender, EventArgs e)
        {

            if (dt_cl_display != null && dt_cl_display.Rows.Count > 0)
            {

                if (dt_cl_display_filtered != null && dt_cl_display_filtered.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_cl_display.Rows.Count; ++i)
                    {
                        double sta2d = -1.23;
                        double sta3d = -4.56;
                        if (dt_cl_display.Rows[i][col_sta2d] != DBNull.Value)
                        {
                            sta2d = Convert.ToDouble(dt_cl_display.Rows[i][col_sta2d]);
                        }

                        if (dt_cl_display.Rows[i][col_sta3d] != DBNull.Value)
                        {
                            sta3d = Convert.ToDouble(dt_cl_display.Rows[i][col_sta3d]);
                        }
                        for (int j = 0; j < dt_cl_display_filtered.Rows.Count; ++j)
                        {
                            double sta1 = -7.89;
                            double sta2 = -0.12;
                            if (dt_cl_display_filtered.Rows[j][col_sta2d] != DBNull.Value)
                            {
                                sta1 = Convert.ToDouble(dt_cl_display_filtered.Rows[j][col_sta2d]);
                            }

                            if (dt_cl_display_filtered.Rows[j][col_sta3d] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(dt_cl_display_filtered.Rows[j][col_sta3d]);
                            }


                            if (sta1 == sta2d || sta2 == sta3d)
                            {
                                dt_cl_display.Rows[i][col_elbow] = dt_cl_display_filtered.Rows[j][col_elbow];
                                dt_cl_display.Rows[i][col_mat_elbow] = dt_cl_display_filtered.Rows[j][col_mat_elbow];
                            }

                        }

                    }

                }

                if (ds_main.dt_points == null)
                {
                    ds_main.dt_points = new System.Data.DataTable();
                    ds_main.dt_points.Columns.Add(col_mmid, typeof(string));
                    ds_main.dt_points.Columns.Add(col_item_no, typeof(string));
                    ds_main.dt_points.Columns.Add(col_sta2d, typeof(double));
                    ds_main.dt_points.Columns.Add(col_sta3d, typeof(double));
                    ds_main.dt_points.Columns.Add(col_eqsta, typeof(double));
                    ds_main.dt_points.Columns.Add(col_symbol, typeof(string));
                    ds_main.dt_points.Columns.Add(col_altdesc, typeof(string));
                    ds_main.dt_points.Columns.Add(col_x, typeof(double));
                    ds_main.dt_points.Columns.Add(col_y, typeof(double));
                    ds_main.dt_points.Columns.Add(col_block, typeof(string));
                    ds_main.dt_points.Columns.Add(col_atr1, typeof(string));
                    ds_main.dt_points.Columns.Add(col_atr2, typeof(string));
                    ds_main.dt_points.Columns.Add(col_atr3, typeof(string));
                    ds_main.dt_points.Columns.Add(col_atr4, typeof(string));
                    ds_main.dt_points.Columns.Add(col_visibility, typeof(string));
                }

                string ct1 = "ELL_POINT";
                int idx1 = ds_main.tpage_mat_design.ct_list.IndexOf(ct1);

                if (ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count > 0)
                {

                    for (int i = 0; i < dt_cl_display.Rows.Count; ++i)
                    {
                        double sta1 = -1.23;
                        if (dt_cl_display.Rows[i][col_sta2d] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt_cl_display.Rows[i][col_sta2d]);
                        }

                        if (dt_cl_display.Rows[i][col_sta3d] != DBNull.Value)
                        {
                            sta1 = Convert.ToDouble(dt_cl_display.Rows[i][col_sta3d]);
                        }

                        for (int j = ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count - 1; j >= 0; --j)
                        {
                            double sta2 = -3.45;
                            if (ds_main.tpage_mat_design.dt_ct[idx1].Rows[j][col_sta2d] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(ds_main.tpage_mat_design.dt_ct[idx1].Rows[j][col_sta2d]);
                            }
                            if (ds_main.tpage_mat_design.dt_ct[idx1].Rows[j][col_sta3d] != DBNull.Value)
                            {
                                sta2 = Convert.ToDouble(ds_main.tpage_mat_design.dt_ct[idx1].Rows[j][col_sta3d]);
                            }

                            if (Math.Round(sta1, 2) == Math.Round(sta2, 2))
                            {
                                ds_main.tpage_mat_design.dt_ct[idx1].Rows[j].Delete();
                            }


                        }

                    }
                }

                for (int i = 0; i < dt_cl_display.Rows.Count; ++i)
                {
                    if (dt_cl_display.Rows[i][col_elbow] != DBNull.Value)
                    {
                        bool is_elbow = Convert.ToBoolean(dt_cl_display.Rows[i][col_elbow]);
                        if (is_elbow == true)
                        {

                            ds_main.tpage_mat_design.dt_ct[idx1].Rows.Add();
                            ds_main.tpage_mat_design.dt_ct[idx1].Rows[ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count - 1][col_sta2d] = dt_cl_display.Rows[i][col_sta2d];
                            ds_main.tpage_mat_design.dt_ct[idx1].Rows[ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count - 1][col_sta3d] = dt_cl_display.Rows[i][col_sta3d];
                            ds_main.tpage_mat_design.dt_ct[idx1].Rows[ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count - 1][col_item_no] = dt_cl_display.Rows[i][col_mat_elbow];
                            string descr1 = "ELBOW";

                            if (dt_cl_display.Rows[i][col_mat_elbow] != DBNull.Value)
                            {
                                string mat1 = Convert.ToString(dt_cl_display.Rows[i][col_mat_elbow]);
                                for (int j = 0; j < dt_mat_library.Rows.Count; ++j)
                                {
                                    if (dt_mat_library.Rows[j][col_item_no] != DBNull.Value && Convert.ToString(dt_mat_library.Rows[j][col_item_no]) == mat1)
                                    {
                                        if (dt_mat_library.Rows[j][col_descr] != DBNull.Value)
                                        {
                                            descr1 = Convert.ToString(dt_mat_library.Rows[j][col_descr]);
                                        }
                                    }
                                }
                            }

                            ds_main.tpage_mat_design.dt_ct[idx1].Rows[ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count - 1][col_altdesc] = descr1;
                            ds_main.tpage_mat_design.dt_ct[idx1].Rows[ds_main.tpage_mat_design.dt_ct[idx1].Rows.Count - 1][col_mmid] = "**ELBOW";

                            ds_main.dt_points.Rows.Add();
                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_sta2d] = dt_cl_display.Rows[i][col_sta2d];
                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_sta3d] = dt_cl_display.Rows[i][col_sta3d];
                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_item_no] = dt_cl_display.Rows[i][col_mat_elbow];
                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_altdesc] = descr1;
                            ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][col_mmid] = "**ELBOW";


                        }

                    }


                }


            }

            //   Functions.Transfer_datatable_to_new_excel_spreadsheet(ds_main.dt_elbows);

        }

        private void button_clear_elbows_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {

                if (dt_cl_display_filtered != null && dt_cl_display_filtered.Rows.Count > 1)
                {
                    for (int i = 0; i < dt_cl_display_filtered.Rows.Count; ++i)
                    {
                        dt_cl_display_filtered.Rows[i][col_elbow] = false;
                        dt_cl_display_filtered.Rows[i][col_mat_elbow] = DBNull.Value;

                    }
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_add_stationing_Click(object sender, EventArgs e)
        {
            try
            {
                if (ds_main.dt_centerline == null || ds_main.dt_centerline.Rows.Count < 2)
                {
                    return;
                }
                set_enable_false();
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {



                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(ds_main.dt_centerline);

                        create_stationing(Poly3D, 0, 2, 4, 100, 10, 4, 2);

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void create_stationing(Polyline3d Poly3D, double start1, double gap1, double texth, double spacing_major, double spacing_minor, double tick_major, double tick_minor)
        {

            string layer_stationing = "MD_Stationing";
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    Polyline Poly2D = Functions.Build_2dpoly_from_3d(Poly3D);
                    Functions.Creaza_layer(layer_stationing, 2, true);


                    int lr = 1;
                    double extra_rot = 0;

                    double first_label_major = Math.Floor((start1 + spacing_major) / spacing_major) * spacing_major;

                    if (start1 + Poly3D.Length >= first_label_major)
                    {
                        int no_major = Convert.ToInt32(Math.Ceiling((start1 + Poly3D.Length - first_label_major) / spacing_major));

                        if (no_major > 0)
                        {
                            for (int i = 0; i < no_major; ++i)
                            {
                                Point3d pt0 = Poly3D.GetPointAtDist((first_label_major - start1) + i * spacing_major);


                                double label_major = first_label_major + i * spacing_major;
                                Autodesk.AutoCAD.DatabaseServices.Line Big1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                double param2 = param1 + 1;
                                if (Poly2D.EndParam < param2)
                                {
                                    param1 = Poly2D.EndParam - 1;
                                    param2 = Poly2D.EndParam;
                                }

                                Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                double rot1 = bear1 - lr * Math.PI / 2;

                                Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                Big1.Layer = layer_stationing;
                                Big1.ColorIndex = 256;


                                BTrecord.AppendEntity(Big1);
                                Trans1.AddNewlyCreatedDBObject(Big1, true);



                                Autodesk.AutoCAD.DatabaseServices.Line l_t = new Autodesk.AutoCAD.DatabaseServices.Line(Big1.StartPoint, Big1.EndPoint);
                                l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                MText mt1 = creaza_mtext_sta(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, "f", 0), texth, bear1 + extra_rot);

                                mt1.Layer = layer_stationing;
                                BTrecord.AppendEntity(mt1);
                                Trans1.AddNewlyCreatedDBObject(mt1, true);



                            }
                        }
                    }

                    double first_label_minor = Math.Floor((start1 + spacing_minor) / spacing_minor) * spacing_minor;

                    if (start1 + Poly3D.Length >= first_label_minor)
                    {
                        int no_minor = Convert.ToInt32(Math.Ceiling((start1 + Poly3D.Length - first_label_minor) / spacing_minor));

                        if (no_minor > 0)
                        {
                            for (int i = 0; i < no_minor; ++i)
                            {
                                Point3d pt0 = Poly3D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                double label_major = first_label_minor + i * spacing_minor;
                                Autodesk.AutoCAD.DatabaseServices.Line small1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                double param2 = param1 + 1;
                                if (Poly2D.EndParam < param2)
                                {
                                    param1 = Poly2D.EndParam - 1;
                                    param2 = Poly2D.EndParam;
                                }


                                Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                double rot1 = bear1 - lr * Math.PI / 2;

                                small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                small1.Layer = layer_stationing;
                                small1.ColorIndex = 256;

                                BTrecord.AppendEntity(small1);
                                Trans1.AddNewlyCreatedDBObject(small1, true);
                            }
                        }
                    }



                    Poly3D.Erase();
                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        public MText creaza_mtext_sta(Point3d pt_ins, string continut, double texth, double rot1)
        {
            MText mtext1 = new MText();
            mtext1.Attachment = AttachmentPoint.BottomCenter;
            mtext1.Contents = continut;
            mtext1.TextHeight = texth;
            mtext1.BackgroundFill = true;
            mtext1.UseBackgroundColor = true;
            mtext1.BackgroundScaleFactor = 1.2;
            mtext1.Location = pt_ins;
            mtext1.Rotation = rot1;
            mtext1.ColorIndex = 256;

            return mtext1;
        }
    }
}

