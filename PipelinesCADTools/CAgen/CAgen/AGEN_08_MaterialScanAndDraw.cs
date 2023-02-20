using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Microsoft.Office.Interop.Excel;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;

namespace Alignment_mdi
{
    public partial class AGEN_MaterialBand : Form
    {
        string col_handle = "handle";
        string col_blockname = "Blockname";
        string col_x = "X";
        string col_y = "Y";
        string col_visibility = "Visibility";
        string col_dist1 = "Distance1";
        string col_dist2 = "Distance2";
        string col_dist3 = "Distance3";
        string col_dist4 = "Distance4";
        string col_sta = "Reference Station";

        System.Data.DataTable dt_pt_extra = null;

        public AGEN_MaterialBand()
        {
            InitializeComponent();
        }


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_mat_band);
            lista_butoane.Add(button_load_materials);

            lista_butoane.Add(button_open_materials);
            lista_butoane.Add(button_place_extra_pts);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_mat_band);
            lista_butoane.Add(button_load_materials);

            lista_butoane.Add(button_open_materials);
            lista_butoane.Add(button_place_extra_pts);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }








        private void button_load_materials_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();

            if (Functions.Get_if_workbook_is_open_in_Excel("Material_Linear.xlsx") == true)
            {
                MessageBox.Show("Please close the material linear file");
                return;
            }

            if (Functions.Get_if_workbook_is_open_in_Excel("Material_Points.xlsx") == true)
            {
                MessageBox.Show("Please close the material points file");
                return;
            }

            if (Functions.Get_if_workbook_is_open_in_Excel("Material_Linear_extra.xlsx") == true)
            {
                MessageBox.Show("Please close the material linear extra file");
                return;
            }

            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }


            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }





            // Ag.WindowState = FormWindowState.Minimized;
            _AGEN_mainform.tpage_processing.Show();

            set_enable_false();

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            string fisier_mat_lin = ProjF + _AGEN_mainform.mat_linear_excel_name;
            string fisier_mat_lin_extra = ProjF + _AGEN_mainform.mat_linear_extra_excel_name;
            string fisier_mat_pt = ProjF + _AGEN_mainform.mat_points_excel_name;
            string fisier_mat = ProjF + _AGEN_mainform.materials_excel_name;

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {

                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                    fisier_mat_lin = ProjF + _AGEN_mainform.mat_linear_excel_name;
                    fisier_mat_pt = ProjF + _AGEN_mainform.mat_points_excel_name;
                    fisier_mat_lin_extra = ProjF + _AGEN_mainform.mat_linear_extra_excel_name;
                    fisier_mat = ProjF + _AGEN_mainform.materials_excel_name;
                }

                if (System.IO.File.Exists(fisier_mat) == false && System.IO.File.Exists(fisier_mat_lin) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the " + _AGEN_mainform.mat_linear_excel_name + " file does not exist");
                    return;
                }

            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            try
            {
                dataGridView_materials.DataSource = null;

                if (System.IO.File.Exists(fisier_mat) == true)
                {
                    Load_existing_materials(fisier_mat, ref _AGEN_mainform.dt_mat_lin, ref _AGEN_mainform.dt_mat_lin_extra, ref _AGEN_mainform.dt_mat_pt, ref dt_pt_extra);
                }
                else
                {
                    _AGEN_mainform.dt_mat_lin = Load_existing_mat_linear(fisier_mat_lin);
                    _AGEN_mainform.dt_mat_lin_extra = null;
                    _AGEN_mainform.dt_mat_pt = null;


                    if (_AGEN_mainform.dt_mat_lin != null && _AGEN_mainform.dt_mat_lin.Rows.Count > 0)
                    {
                        _AGEN_mainform.dt_mat_lin_extra = Load_existing_mat_linear_extra(fisier_mat_lin_extra);
                        _AGEN_mainform.dt_mat_pt = Load_existing_mat_point(fisier_mat_pt);
                    }
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            _AGEN_mainform.tpage_processing.Hide();
            Ag.WindowState = FormWindowState.Normal;
            set_enable_true();

        }


        public void Load_existing_materials(string File1, ref System.Data.DataTable dtml, ref System.Data.DataTable dtextra, ref System.Data.DataTable dtpt, ref System.Data.DataTable dtpt_extra)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the material data file does not exist");
                return;
            }
            dtml = new System.Data.DataTable();
            dtextra = new System.Data.DataTable();
            dtpt = new System.Data.DataTable();
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W2 = null;
            Microsoft.Office.Interop.Excel.Worksheet W3 = null;
            Microsoft.Office.Interop.Excel.Worksheet W4 = null;
            bool close_excel = false;
            string nume_tab_lin = "linear";
            string nume_tab_extra = "linear extra";
            string nume_tab_pts = "points";
            string nume_tab_pts_extra = "points extra";
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == File1.ToLower())
                        {
                            Workbook1 = Workbook2;
                            foreach (Microsoft.Office.Interop.Excel.Worksheet W5 in Workbook1.Worksheets)
                            {
                                if (W5.Name.ToLower() == nume_tab_lin)
                                {
                                    W1 = W5;
                                }
                                if (W5.Name.ToLower() == nume_tab_extra)
                                {
                                    W2 = W5;
                                }
                                if (W5.Name.ToLower() == nume_tab_pts)
                                {
                                    W3 = W5;
                                }
                                if (W5.Name.ToLower() == nume_tab_pts_extra)
                                {
                                    W4 = W5;
                                }
                            }
                        }

                    }


                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }




                if (W1 == null)
                {
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                    Workbook1 = Excel1.Workbooks.Open(File1);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W5 in Workbook1.Worksheets)
                    {
                        if (W5.Name.ToLower() == nume_tab_lin)
                        {
                            W1 = W5;
                        }
                        if (W5.Name.ToLower() == nume_tab_extra)
                        {
                            W2 = W5;
                        }
                        if (W5.Name.ToLower() == nume_tab_pts)
                        {
                            W3 = W5;
                        }
                        if (W5.Name.ToLower() == nume_tab_pts_extra)
                        {
                            W4 = W5;
                        }
                    }
                    if (W1 == null)
                    {
                        W1 = Workbook1.Worksheets[1];
                    }
                    close_excel = true;
                }



                try
                {
                    dtml = Functions.Build_Data_table_mat_linear_from_excel(W1, _AGEN_mainform.Start_row_mat_lin + 1);
                    if (W2 != null) dtextra = Functions.Build_Data_table_mat_linear_from_excel(W2, _AGEN_mainform.Start_row_mat_lin + 1);
                    if (W3 != null) dtpt = Functions.Build_Data_table_mat_point_from_excel(W3, _AGEN_mainform.Start_row_mat_point + 1);
                    if (W4 != null) dtpt_extra = Functions.Build_Data_table_mat_point_from_excel(W4, _AGEN_mainform.Start_row_mat_point + 1);
                    if (close_excel == true) Workbook1.Close();
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
                    if (close_excel == true && W1 != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    }
                    if (close_excel == true && W2 != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                    }
                    if (close_excel == true && W3 != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                    }
                    if (close_excel == true && Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1.Workbooks.Count == 0 && Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }


        public System.Data.DataTable Load_existing_mat_linear(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the linear material data file does not exist");
                return null;
            }
            System.Data.DataTable dtml = new System.Data.DataTable();
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                try
                {
                    dtml = Functions.Build_Data_table_mat_linear_from_excel(W1, _AGEN_mainform.Start_row_mat_lin + 1);
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
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dtml;
        }

        public System.Data.DataTable Load_existing_mat_linear_extra(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {

                return null;
            }


            System.Data.DataTable dtml = new System.Data.DataTable();

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dtml = Functions.Build_Data_table_mat_linear_from_excel(W1, _AGEN_mainform.Start_row_mat_lin + 1);
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
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dtml;

        }

        public System.Data.DataTable Load_existing_mat_point(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {

                return null;
            }


            System.Data.DataTable dtmp = new System.Data.DataTable();

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {


                    dtmp = Functions.Build_Data_table_mat_point_from_excel(W1, _AGEN_mainform.Start_row_mat_point + 1);

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
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dtmp;

        }



        private void button_draw_mat_band_Click(object sender, EventArgs e)
        {
            if (_AGEN_mainform.COUNTRY == "USA" && _AGEN_mainform.Project_type == "2D")
            {
                draw_mat_band_USA_2D();
                return;
            }

            System.Data.DataTable dtmc = new System.Data.DataTable();
            dtmc.Columns.Add("band", typeof(string));
            string debug = "00";

            string lnp = "Agen_no_plot_mat";
            double min_spacing = 0.4;
            if (_AGEN_mainform.COUNTRY == "CANADA")
            {
                min_spacing = 25;
            }

            if (Functions.IsNumeric(textBox_spacing.Text) == true)
            {
                min_spacing = Math.Abs(Convert.ToDouble(textBox_spacing.Text));
            }



            int lr = 1;
            if (_AGEN_mainform.Left_to_Right == false) lr = -1;

            Functions.Kill_excel();


            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }

            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }

            if (_AGEN_mainform.dt_mat_lin == null || _AGEN_mainform.dt_mat_lin.Rows.Count == 0)
            {
                MessageBox.Show("No linear material data found");
                return;
            }

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            //Ag.WindowState = FormWindowState.Minimized;

            if (_AGEN_mainform.Vw_mat_height == 0)
            {
                MessageBox.Show("you did not specified viewport material information");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();

                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();


                _AGEN_mainform.tpage_viewport_settings.Show();

                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }

            set_enable_false();

            _AGEN_mainform.tpage_processing.Show();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }



                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                if (System.IO.File.Exists(fisier_si) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the sheet index data file does not exist");
                    return;
                }

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the sheet index data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }


                _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);

                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("no centerline");
                    return;
                }

            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            if (_AGEN_mainform.dt_mat_lin.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the linear material file does not have any data");
                return;
            }

            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the sheet index file does not have any data");
                return;
            }



            string Sta1 = "Sta1";
            string Sta2 = "Sta2";
            string Sta1_label = "Sta1_label";
            string Sta2_label = "Sta2_label";
            string Match1 = "M1";
            string Match2 = "M2";
            string mat_atr = "material_for_align";
            string Pageno = "Page";
            string Rect_len = "RectangleML";
            string stretch_val = "StrechVal";
            string BandL = "BandL";
            string DeltaX_col = "DeltaX";
            string stretch_val_orig = "StrechValoriginal";


            System.Data.DataTable dt_compiled = new System.Data.DataTable();
            dt_compiled.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));       //0
            dt_compiled.Columns.Add(Sta1, typeof(double));                              //1
            dt_compiled.Columns.Add(Sta2, typeof(double));                              //2
            dt_compiled.Columns.Add(mat_atr, typeof(string));                           //3
            dt_compiled.Columns.Add(Pageno, typeof(int));                               //4
            dt_compiled.Columns.Add(Rect_len, typeof(double));                          //5
            dt_compiled.Columns.Add(BandL, typeof(double));                             //6
            dt_compiled.Columns.Add(DeltaX_col, typeof(double));                        //7
            dt_compiled.Columns.Add(Match1, typeof(double));                            //8
            dt_compiled.Columns.Add(Match2, typeof(double));                            //9
            dt_compiled.Columns.Add(stretch_val, typeof(double));                       //10
            dt_compiled.Columns.Add(stretch_val_orig, typeof(double));                  //11
            dt_compiled.Columns.Add(Sta1_label, typeof(double));                        //12
            dt_compiled.Columns.Add(Sta2_label, typeof(double));                        //13

            for (int n = 15; n < _AGEN_mainform.dt_mat_lin.Columns.Count; ++n)
            {
                dt_compiled.Columns.Add(_AGEN_mainform.dt_mat_lin.Columns[n].ColumnName, typeof(string));
            }

            #region Data_table_compiled_extra
            System.Data.DataTable Data_table_compiled_extra = null;
            if (_AGEN_mainform.dt_mat_lin_extra != null && _AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
            {
                Data_table_compiled_extra = new System.Data.DataTable();
                Data_table_compiled_extra.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));     //0
                Data_table_compiled_extra.Columns.Add(Sta1, typeof(double));                            //1
                Data_table_compiled_extra.Columns.Add(Sta2, typeof(double));                            //2
                Data_table_compiled_extra.Columns.Add(mat_atr, typeof(string));                         //3
                Data_table_compiled_extra.Columns.Add(Pageno, typeof(int));                             //4
                Data_table_compiled_extra.Columns.Add(Rect_len, typeof(double));                        //5
                Data_table_compiled_extra.Columns.Add(BandL, typeof(double));                           //6
                Data_table_compiled_extra.Columns.Add(DeltaX_col, typeof(double));                      //7
                Data_table_compiled_extra.Columns.Add(Match1, typeof(double));                          //8
                Data_table_compiled_extra.Columns.Add(Match2, typeof(double));                          //9
                Data_table_compiled_extra.Columns.Add(stretch_val, typeof(double));                     //10
                Data_table_compiled_extra.Columns.Add(stretch_val_orig, typeof(double));                //11
                Data_table_compiled_extra.Columns.Add(Sta1_label, typeof(double));                      //12
                Data_table_compiled_extra.Columns.Add(Sta2_label, typeof(double));                      //13

                for (int n = 15; n < _AGEN_mainform.dt_mat_lin_extra.Columns.Count; ++n)
                {
                    Data_table_compiled_extra.Columns.Add(_AGEN_mainform.dt_mat_lin_extra.Columns[n].ColumnName, typeof(string));
                }
            }
            #endregion





            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;

                        Polyline3d poly3d = null;
                        Polyline poly2d = null;

                        double poly_length = 0;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            poly_length = poly3d.Length;
                        }


                        poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                        if (_AGEN_mainform.Project_type == "2D")
                        {
                            poly_length = poly2d.Length;
                        }

                        _AGEN_mainform.dt_sheet_index = Functions.Redefine_stations_for_sheet_index(_AGEN_mainform.dt_sheet_index, poly3d, poly2d);


                        #region USA
                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.COUNTRY == "USA")
                        {
                            if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                            {
                                if (_AGEN_mainform.dt_station_equation.Columns.Contains("measured") == false)
                                {
                                    _AGEN_mainform.dt_station_equation.Columns.Add("measured", typeof(double));
                                }

                                for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                    {
                                        double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                        double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);

                                        Point3d pt_on_2d = poly2d.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                        double eq_meas = poly2d.GetDistAtPoint(pt_on_2d);

                                        if (_AGEN_mainform.Project_type == "3D")
                                        {
                                            double param1 = poly2d.GetParameterAtPoint(pt_on_2d);
                                            if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                            eq_meas = poly3d.GetDistanceAtParameter(param1);
                                        }
                                        _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;
                                    }
                                }
                            }

                        }
                        #endregion

                        Functions.Creaza_layer(lnp, 30, false);
                        string lname1 = "Agen_band_mat";
                        Functions.Creaza_layer(lname1, 7, true);

                        List<int> lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_indexes_for_generation(_AGEN_mainform.Point0_mat, _AGEN_mainform.Band_Separation, lnp);

                        #region adauga breaks for matchlines dt_mat_lin
                        for (int i = 0; i < _AGEN_mainform.dt_mat_lin.Rows.Count; ++i)
                        {

                            int m_start = 0;
                            bool Boolean_go_to_check_s1_s2 = false;
                            double Station1 = -1.123;
                            double Station2 = -1.123;
                            string Material1 = "";
                            double Station1_labeled = -1.123;
                            double Station2_labeled = -1.123;

                            if (_AGEN_mainform.Project_type == "2D")
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta1] != DBNull.Value && _AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta2] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta1]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta2]);
                                }
                            }
                            else
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta1] != DBNull.Value && _AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta2] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta1]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta2]);
                                }
                            }

                            Station1_labeled = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                            Station2_labeled = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);

                            if (_AGEN_mainform.COUNTRY == "CANADA" &&
                                _AGEN_mainform.dt_mat_lin.Columns.Contains("MeasuredStartCanada") && _AGEN_mainform.dt_mat_lin.Columns.Contains("MeasuredEndCanada") &&
                                _AGEN_mainform.dt_mat_lin.Rows[i]["MeasuredStartCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin.Rows[i]["MeasuredStartCanada"])) == true &&
                                _AGEN_mainform.dt_mat_lin.Rows[i]["MeasuredEndCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin.Rows[i]["MeasuredEndCanada"])) == true)
                            {
                                Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i]["MeasuredStartCanada"]);
                                Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i]["MeasuredEndCanada"]);
                            }
                            else
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i]["X_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin.Rows[i]["X_Beg"])) == true &&
                                    _AGEN_mainform.dt_mat_lin.Rows[i]["Y_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin.Rows[i]["Y_Beg"])) == true &&
                                    _AGEN_mainform.dt_mat_lin.Rows[i]["X_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin.Rows[i]["X_End"])) == true &&
                                    _AGEN_mainform.dt_mat_lin.Rows[i]["Y_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin.Rows[i]["Y_End"])) == true)
                                {
                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i]["X_Beg"]);
                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i]["Y_Beg"]);
                                    double x2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i]["X_End"]);
                                    double y2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i]["Y_End"]);
                                    Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);
                                    Point3d point_on_poly2D2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, false);

                                    Station1 = poly2d.GetDistAtPoint(point_on_poly2D1);
                                    Station2 = poly2d.GetDistAtPoint(point_on_poly2D2);

                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        double param1 = poly2d.GetParameterAtPoint(point_on_poly2D1);
                                        double param2 = poly2d.GetParameterAtPoint(point_on_poly2D2);
                                        if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                        if (param2 > poly3d.EndParam) param2 = poly3d.EndParam;

                                        Station1 = poly3d.GetDistanceAtParameter(param1);
                                        Station2 = poly3d.GetDistanceAtParameter(param2);
                                    }


                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                    {
                                        double d2d1 = poly2d.GetDistAtPoint(point_on_poly2D1);
                                        double d2d2 = poly2d.GetDistAtPoint(point_on_poly2D2);
                                        double b1 = -1.23456;
                                        double b2 = -1.23456;
                                        Station1_labeled = Functions.get_stationCSF_from_point(poly2d, point_on_poly2D1, d2d1, _AGEN_mainform.dt_centerline, ref b1);
                                        Station2_labeled = Functions.get_stationCSF_from_point(poly2d, point_on_poly2D2, d2d2, _AGEN_mainform.dt_centerline, ref b2);
                                    }
                                    else
                                    {
                                        Station1_labeled = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                        Station2_labeled = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);
                                    }
                                }
                            }
                            if (Station1 != -1.123 && Station2 != -1.123)
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_Material] != DBNull.Value)
                                {
                                    Material1 = _AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_Material].ToString();
                                }
                            L123:
                                for (int j = m_start; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                {
                                    if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2] != DBNull.Value)
                                    {
                                        double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1]);
                                        double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2]);


                                        if (M2 <= M1)
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("End Station is smaller than Start Station on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                            return;
                                        }
                                        if (M2 > poly_length)
                                        {
                                            if (Math.Abs(M2 - poly_length) < 0.99)
                                            {
                                                M2 = poly_length;
                                            }
                                            else
                                            {
                                                _AGEN_mainform.tpage_processing.Hide();
                                                set_enable_true();
                                                MessageBox.Show("End Station is bigger than poly length on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                return;
                                            }
                                        }
                                        if (M1 >= poly_length) M1 = poly_length - 0.0001;
                                        if (M2 >= poly_length) M2 = poly_length - 0.0001;


                                        Point3d pm1 = poly2d.GetPointAtDist(M1);
                                        Point3d pm2 = poly2d.GetPointAtDist(M2);

                                        if (_AGEN_mainform.Project_type == "3D")
                                        {
                                            pm1 = poly3d.GetPointAtDist(M1);
                                            pm2 = poly3d.GetPointAtDist(M2);
                                        }
                                        pm1 = new Point3d(pm1.X, pm1.Y, 0);
                                        pm2 = new Point3d(pm2.X, pm2.Y, 0);

                                        Autodesk.AutoCAD.DatabaseServices.Line Linie_M1_M2 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pm1.X, pm1.Y, 0), new Point3d(pm2.X, pm2.Y, 0));

                                        if (Boolean_go_to_check_s1_s2 == true)
                                        {
                                            if (Math.Round(Station1, 0) == Math.Round(Station2, 0))
                                            {
                                                goto LS12end;
                                            }
                                            goto LS1S2;
                                        }

                                        if (Math.Round(M1, 2) <= Math.Round(Station1, 2) && Math.Round(M2, 2) <= Math.Round(Station2, 2) && Math.Round(M1, 2) <= Math.Round(Station2, 2) && Math.Round(M2, 2) > Math.Round(Station1, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);

                                            }

                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            else
                                            {
                                                double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            Station1 = M2;
                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }

                                        if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                Pt2 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                            }
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = Station2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = Station2_labeled;
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }

                                    LS1S2:
                                        if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                Pt2 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                            }

                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = Station2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = Station2_labeled;
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }
                                        else if (Math.Round(Station1, 2) < Math.Round(M2, 2) && Math.Round(Station1, 2) >= Math.Round(M1, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            }

                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            else
                                            {
                                                double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            Station1 = M2;
                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }
                                    }
                                }
                            LS12end:
                                string xx = "";
                            }
                        }
                        #endregion




                        #region adauga breaks for matchlines dt_mat_lin_extra_extra
                        if (_AGEN_mainform.dt_mat_lin_extra != null && _AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
                        {
                            for (int i = 0; i < _AGEN_mainform.dt_mat_lin_extra.Rows.Count; ++i)
                            {
                                int m_start = 0;
                                bool Boolean_go_to_check_s1_s2 = false;
                                double Station1 = -1.123;
                                double Station2 = -1.123;
                                string Material1 = "";
                                double Station1_labeled = -1.123;
                                double Station2_labeled = -1.123;
                                if (_AGEN_mainform.Project_type == "2D")
                                {
                                    if (_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta1] != DBNull.Value && _AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta2] != DBNull.Value)
                                    {
                                        Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta1]);
                                        Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta2]);
                                    }
                                }
                                else
                                {
                                    if (_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_3DSta1] != DBNull.Value && _AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_3DSta2] != DBNull.Value)
                                    {
                                        Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_3DSta1]);
                                        Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_3DSta2]);
                                    }
                                }

                                Station1_labeled = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                Station2_labeled = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);

                                if (_AGEN_mainform.COUNTRY == "CANADA" &&
                                        _AGEN_mainform.dt_mat_lin_extra.Columns.Contains("MeasuredStartCanada") && _AGEN_mainform.dt_mat_lin_extra.Columns.Contains("MeasuredEndCanada") &&
                                        _AGEN_mainform.dt_mat_lin_extra.Rows[i]["MeasuredStartCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["MeasuredStartCanada"])) == true &&
                                        _AGEN_mainform.dt_mat_lin_extra.Rows[i]["MeasuredEndCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["MeasuredEndCanada"])) == true)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["MeasuredStartCanada"]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["MeasuredEndCanada"]);
                                }

                                else
                                {
                                    if (_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_Beg"])) == true &&
                                        _AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_Beg"])) == true &&
                                        _AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_End"])) == true &&
                                        _AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_End"])) == true)
                                    {
                                        double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_Beg"]);
                                        double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_Beg"]);
                                        double x2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_End"]);
                                        double y2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_End"]);
                                        Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);
                                        Point3d point_on_poly2D2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, false);

                                        Station1 = poly2d.GetDistAtPoint(point_on_poly2D1);
                                        Station2 = poly2d.GetDistAtPoint(point_on_poly2D2);

                                        if (_AGEN_mainform.Project_type == "3D")
                                        {
                                            double param1 = poly2d.GetParameterAtPoint(point_on_poly2D1);
                                            double param2 = poly2d.GetParameterAtPoint(point_on_poly2D2);
                                            if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                            if (param2 > poly3d.EndParam) param2 = poly3d.EndParam;
                                            Station1 = poly3d.GetDistanceAtParameter(param1);
                                            Station2 = poly3d.GetDistanceAtParameter(param2);
                                        }


                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                        {
                                            double d2d1 = poly2d.GetDistAtPoint(point_on_poly2D1);
                                            double d2d2 = poly2d.GetDistAtPoint(point_on_poly2D2);
                                            double b1 = -1.23456;
                                            Station1_labeled = Functions.get_stationCSF_from_point(poly2d, point_on_poly2D1, d2d1, _AGEN_mainform.dt_centerline, ref b1);
                                            double b2 = -1.23456;
                                            Station2_labeled = Functions.get_stationCSF_from_point(poly2d, point_on_poly2D2, d2d2, _AGEN_mainform.dt_centerline, ref b2);
                                        }
                                        else
                                        {
                                            Station1_labeled = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                            Station2_labeled = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);
                                        }
                                    }
                                }



                                if (Station1 != -1.123 && Station2 != -1.123)
                                {
                                    if (_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_Material] != DBNull.Value)
                                    {
                                        Material1 = _AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_Material].ToString();
                                    }
                                L123:
                                    for (int j = m_start; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                    {
                                        if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2] != DBNull.Value)
                                        {
                                            double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1]);
                                            double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2]);

                                            if (M2 <= M1)
                                            {
                                                _AGEN_mainform.tpage_processing.Hide();
                                                set_enable_true();
                                                MessageBox.Show("End Station is smaller than Start Station on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                return;
                                            }
                                            if (M2 > poly_length)
                                            {
                                                if (Math.Abs(M2 - poly_length) < 0.99)
                                                {
                                                    M2 = poly_length;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.tpage_processing.Hide();
                                                    set_enable_true();
                                                    MessageBox.Show("End Station is bigger than poly length on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                    return;
                                                }
                                            }
                                            if (M1 >= poly_length) M1 = poly_length - 0.0001;
                                            if (M2 >= poly_length) M2 = poly_length - 0.0001;

                                            Point3d pm1 = poly2d.GetPointAtDist(M1);
                                            Point3d pm2 = poly2d.GetPointAtDist(M2);

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                pm1 = poly3d.GetPointAtDist(M1);
                                                pm2 = poly3d.GetPointAtDist(M2);
                                            }

                                            pm1 = new Point3d(pm1.X, pm1.Y, 0);
                                            pm2 = new Point3d(pm2.X, pm2.Y, 0);
                                            Autodesk.AutoCAD.DatabaseServices.Line Linie_M1_M2 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pm1.X, pm1.Y, 0), new Point3d(pm2.X, pm2.Y, 0));
                                            if (Boolean_go_to_check_s1_s2 == true)
                                            {
                                                if (Math.Round(Station1, 0) == Math.Round(Station2, 0))
                                                {
                                                    goto LS12end;
                                                }
                                                goto LS1S2;
                                            }

                                            if (Math.Round(M1, 2) <= Math.Round(Station1, 2) && Math.Round(M2, 2) <= Math.Round(Station2, 2) && Math.Round(M1, 2) <= Math.Round(Station2, 2) && Math.Round(M2, 2) > Math.Round(Station1, 2))
                                            {

                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                }
                                                double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                else
                                                {
                                                    double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                int idx_lin = 15;

                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);
                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }
                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;

                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {

                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }



                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }


                                                Station1 = M2;
                                                m_start = j + 1;
                                                Boolean_go_to_check_s1_s2 = true;
                                                goto L123;
                                            }

                                            if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                    Pt2 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                                }

                                                double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = Station2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = Station2_labeled;

                                                int idx_lin = 15;
                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }


                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;

                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }


                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }

                                                j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                goto LS12end;
                                            }

                                        LS1S2:
                                            if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                    Pt2 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                                }

                                                double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = Station2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = Station2_labeled;
                                                int idx_lin = 15;

                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }


                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;

                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {

                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }

                                                j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                goto LS12end;
                                            }
                                            else if (Math.Round(Station1, 2) < Math.Round(M2, 2) && Math.Round(Station1, 2) >= Math.Round(M1, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    Pt1 = Linie_M1_M2.GetClosestPointTo(poly3d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                }

                                                double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                else
                                                {
                                                    double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                int idx_lin = 15;

                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }


                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;


                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }


                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }
                                                Station1 = M2;
                                                m_start = j + 1;
                                                Boolean_go_to_check_s1_s2 = true;
                                                goto L123;
                                            }
                                        }
                                    }
                                LS12end:
                                    string xx = "";
                                }
                            }
                        }

                        #endregion

                        int Pagep = -1;

                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_compiled);

                        if (dt_compiled != null && dt_compiled.Rows.Count > 0)
                        {

                            #region draw red and green rectangles
                            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                            {
                                int Page1 = Convert.ToInt32(dt_compiled.Rows[i][Pageno]);
                                double ml_len = Convert.ToDouble(dt_compiled.Rows[i][Rect_len]);
                                string dwg_name = dt_compiled.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                double strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                if (strech1 < min_spacing)
                                {
                                    double new_stretch = min_spacing;
                                    dt_compiled.Rows[i][stretch_val] = new_stretch;
                                    double Diff = new_stretch - strech1;
                                    for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                    {
                                        int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                        double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                        double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);
                                        if (Page1 == Page2)
                                        {
                                            dt_compiled.Rows[j][BandL] = band_len2 + Diff;
                                            if (i < j)
                                            {
                                                dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                            }
                                        }
                                    }
                                }
                                if (Page1 != Pagep)
                                {
                                    if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                    {

                                        dtmc.Rows.Add();
                                        dtmc.Rows[dtmc.Rows.Count - 1][0] = dwg_name;

                                        Polyline vp_vw1 = new Polyline();
                                        vp_vw1.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_mat.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_mat.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_mat.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_mat.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.Closed = true;
                                        vp_vw1.Layer = lnp;
                                        vp_vw1.ColorIndex = 3;
                                        BTrecord.AppendEntity(vp_vw1);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw1, true);

                                        Polyline vp_vw2 = new Polyline();
                                        vp_vw2.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_mat.X - ml_len / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_mat.X + ml_len / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_mat.X + ml_len / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_mat.X - ml_len / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);

                                        vp_vw2.Closed = true;
                                        vp_vw2.Layer = lnp;
                                        vp_vw2.ColorIndex = 1;
                                        BTrecord.AppendEntity(vp_vw2);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw2, true);

                                        MText Band_label = new MText();
                                        Band_label.Contents = dwg_name;
                                        Band_label.TextHeight = _AGEN_mainform.Vw_mat_height / 3;
                                        Band_label.Rotation = 0;
                                        Band_label.Attachment = AttachmentPoint.MiddleLeft;
                                        Band_label.Location = new Point3d(_AGEN_mainform.Point0_mat.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height / 2 - (Page1 - 1) * _AGEN_mainform.Band_Separation, 0);
                                        Band_label.Layer = lnp;
                                        BTrecord.AppendEntity(Band_label);
                                        Trans1.AddNewlyCreatedDBObject(Band_label, true);
                                    }
                                    Pagep = Page1;
                                }
                            }
                            #endregion

                            if (_AGEN_mainform.dt_mat_pt != null)
                            {
                                if (_AGEN_mainform.dt_mat_pt.Rows.Count > 0)
                                {
                                    if (_AGEN_mainform.Project_type=="2D")
                                    {
                                        _AGEN_mainform.dt_mat_pt = Functions.Sort_data_table(_AGEN_mainform.dt_mat_pt, _AGEN_mainform.Col_2DSta);
                                    }
                                    else
                                    {
                                        _AGEN_mainform.dt_mat_pt = Functions.Sort_data_table(_AGEN_mainform.dt_mat_pt, _AGEN_mainform.Col_3DSta);
                                    }
                                }
                            }

                            if (_AGEN_mainform.dt_mat_lin_extra != null)
                            {
                                if (_AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
                                {
                                    if (_AGEN_mainform.Project_type=="2D")
                                    {
                                        _AGEN_mainform.dt_mat_lin_extra = Functions.Sort_data_table(_AGEN_mainform.dt_mat_lin_extra, "2DSTABEG");
                                    }
                                    else
                                    {
                                        _AGEN_mainform.dt_mat_lin_extra = Functions.Sort_data_table(_AGEN_mainform.dt_mat_lin_extra, "3DSTABEG");
                                    }
                                }
                            }

                            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                            {
                                int Page1 = Convert.ToInt32(dt_compiled.Rows[i][Pageno]);
                                double Station1 = Convert.ToDouble(dt_compiled.Rows[i][Sta1]);
                                double Station2 = Convert.ToDouble(dt_compiled.Rows[i][Sta2]);
                                double strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                double Stretch_from_pts = 0;
                                double Stretch_from_extra1 = 0;
                                double Stretch_from_extra2 = 0;
                                bool crosing_found_inside = false;
                                bool extra_found_inside1 = false;
                                bool extra_found_inside2 = false;

                                #region dt_points
                                if (_AGEN_mainform.dt_mat_pt != null)
                                {
                                    for (int k = 0; k < _AGEN_mainform.dt_mat_pt.Rows.Count; ++k)
                                    {
                                        double Station_pt = -1.123;
                                        if (_AGEN_mainform.Project_type == "2D")
                                        {
                                            Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_2DSta]);
                                        }
                                        else
                                        {
                                            Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_3DSta]);
                                        }

                                        if (_AGEN_mainform.COUNTRY == "CANADA" &&
                                            _AGEN_mainform.dt_mat_pt.Columns.Contains("MeasuredCanada") &&
                                            _AGEN_mainform.dt_mat_pt.Rows[k]["MeasuredCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["MeasuredCanada"])) == true)
                                        {
                                            Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["MeasuredCanada"]);
                                        }
                                        else
                                        {
                                            if (_AGEN_mainform.dt_mat_pt.Rows[k]["X"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["X"])) == true)
                                            {
                                                double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["X"]);
                                                double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["Y"]);
                                                Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);
                                                Station_pt = poly2d.GetDistAtPoint(point_on_poly2D1);

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    double param1 = poly2d.GetParameterAtPoint(point_on_poly2D1);
                                                    if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                    Station_pt = poly3d.GetDistanceAtParameter(param1);
                                                }
                                            }
                                        }



                                        if (Station_pt > Station1 && Station_pt <= Station2)
                                        {
                                            Stretch_from_pts = Stretch_from_pts + 1 * min_spacing;
                                            crosing_found_inside = true;
                                        }
                                        else if (Station_pt == 0)
                                        {
                                            if (Station_pt <= Station2)
                                            {
                                                Stretch_from_pts = Stretch_from_pts + 1 * min_spacing;
                                                crosing_found_inside = true;
                                            }
                                        }
                                    }

                                    if (crosing_found_inside == true && strech1 < Stretch_from_pts)
                                    {
                                        dt_compiled.Rows[i][stretch_val] = Stretch_from_pts;
                                        double Diff = Stretch_from_pts - strech1;
                                        for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                        {
                                            int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                            double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                            double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);

                                            if (Page1 == Page2)
                                            {
                                                dt_compiled.Rows[j][BandL] = band_len2 + Diff;

                                                if (i < j)
                                                {
                                                    dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                                }
                                            }
                                        }


                                    }
                                }

                                #endregion

                                #region dt_extra
                                if (_AGEN_mainform.dt_mat_lin_extra != null && _AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
                                {
                                    strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                    for (int k = 0; k < _AGEN_mainform.dt_mat_lin_extra.Rows.Count; ++k)
                                    {
                                        double Station_pt1 = -1.123;
                                        double Station_pt2 = -1.123;
                                        if (_AGEN_mainform.Project_type == "2D")
                                        {
                                            Station_pt1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["2DSTABEG"]);
                                        }
                                        else
                                        {
                                            Station_pt1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["3DSTABEG"]);
                                        }

                                        if (_AGEN_mainform.Project_type == "2D")
                                        {
                                            Station_pt2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["2DSTAEND"]);
                                        }
                                        else
                                        {
                                            Station_pt2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["3DSTAEND"]);
                                        }

                                        if (_AGEN_mainform.COUNTRY == "CANADA" &&
                                            _AGEN_mainform.dt_mat_lin_extra.Columns.Contains("MeasuredStartCanada") && _AGEN_mainform.dt_mat_lin_extra.Columns.Contains("MeasuredEndCanada") &&
                                            _AGEN_mainform.dt_mat_lin_extra.Rows[k]["MeasuredStartCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["MeasuredStartCanada"])) == true &&
                                            _AGEN_mainform.dt_mat_lin_extra.Rows[k]["MeasuredEndCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["MeasuredEndCanada"])) == true)
                                        {
                                            Station_pt1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["MeasuredStartCanada"]);
                                            Station_pt2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["MeasuredEndCanada"]);
                                        }
                                        else
                                        {
                                            if (_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_Beg"])) == true &&
                                                _AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_End"])) == true &&
                                                _AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_Beg"])) == true &&
                                                _AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_End"])) == true)
                                            {
                                                double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_Beg"]);
                                                double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_Beg"]);

                                                double x2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_End"]);
                                                double y2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_End"]);

                                                Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);
                                                Station_pt1 = poly2d.GetDistAtPoint(point_on_poly2D1);

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    double param1 = poly2d.GetParameterAtPoint(point_on_poly2D1);
                                                    if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;

                                                    Station_pt1 = poly3d.GetDistanceAtParameter(param1);
                                                }

                                                Point3d point_on_poly2D2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, false);
                                                Station_pt2 = poly3d.GetDistAtPoint(point_on_poly2D2);

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    double param2 = poly2d.GetParameterAtPoint(point_on_poly2D2);

                                                    if (param2 > poly3d.EndParam) param2 = poly3d.EndParam;
                                                    Station_pt2 = poly3d.GetDistanceAtParameter(param2);
                                                }
                                            }
                                        }


                                        if (Station_pt1 > Station1 && Station_pt1 <= Station2)
                                        {
                                            Stretch_from_extra1 = Stretch_from_extra1 + 1 * min_spacing;
                                            extra_found_inside1 = true;
                                        }
                                        else if (Station_pt1 == 0)
                                        {
                                            if (Station_pt1 <= Station2)
                                            {
                                                Stretch_from_extra1 = Stretch_from_extra1 + 1 * min_spacing;
                                                extra_found_inside1 = true;
                                            }
                                        }

                                        if (Station_pt2 > Station1 && Station_pt2 <= Station2)
                                        {
                                            Stretch_from_extra2 = Stretch_from_extra2 + 1 * min_spacing;
                                            extra_found_inside2 = true;
                                        }
                                        else if (Station_pt2 == 0)
                                        {
                                            if (Station_pt2 <= Station2)
                                            {
                                                Stretch_from_extra2 = Stretch_from_extra2 + 1 * min_spacing;
                                                extra_found_inside2 = true;
                                            }
                                        }
                                    }

                                    if (extra_found_inside1 == true && strech1 < Stretch_from_extra1)
                                    {
                                        dt_compiled.Rows[i][stretch_val] = Stretch_from_extra1;
                                        double Diff = Stretch_from_extra1 - strech1;
                                        for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                        {
                                            int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                            double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                            double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);

                                            if (Page1 == Page2)
                                            {
                                                dt_compiled.Rows[j][BandL] = band_len2 + Diff;

                                                if (i < j)
                                                {
                                                    dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                                }
                                            }
                                        }
                                        strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                    }

                                    if (extra_found_inside2 == true && strech1 < Stretch_from_extra2)
                                    {
                                        dt_compiled.Rows[i][stretch_val] = Stretch_from_extra2;
                                        double Diff = Stretch_from_extra2 - strech1;
                                        for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                        {
                                            int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                            double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                            double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);

                                            if (Page1 == Page2)
                                            {
                                                dt_compiled.Rows[j][BandL] = band_len2 + Diff;

                                                if (i < j)
                                                {
                                                    dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                            }




                            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                            {
                                double Station1 = Convert.ToDouble(dt_compiled.Rows[i][Sta1]);
                                double Station2 = Convert.ToDouble(dt_compiled.Rows[i][Sta2]);
                                double M1 = Convert.ToDouble(dt_compiled.Rows[i][Match1]);
                                double M2 = Convert.ToDouble(dt_compiled.Rows[i][Match2]);
                                int Page1 = Convert.ToInt32(dt_compiled.Rows[i][Pageno]);
                                string dwg_name = dt_compiled.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();


                                double ml_len = Convert.ToDouble(dt_compiled.Rows[i][Rect_len]);
                                double band_len = Convert.ToDouble(dt_compiled.Rows[i][BandL]);
                                double Diff = (band_len - ml_len) / 2;
                                double deltax = Convert.ToDouble(dt_compiled.Rows[i][DeltaX_col]);

                                double Station1_label = Convert.ToDouble(dt_compiled.Rows[i][Sta1_label]);
                                double Station2_label = Convert.ToDouble(dt_compiled.Rows[i][Sta2_label]);


                                Point3d InsPt = new Point3d(_AGEN_mainform.Point0_mat.X - lr * ml_len / 2 + lr * deltax, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation, 0);

                                string S1 = Functions.Get_chainage_from_double(Station1_label, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                string S2 = Functions.Get_chainage_from_double(Station2_label, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                string Suff = "'";
                                if (_AGEN_mainform.units_of_measurement == "m") Suff = "";

                                string L1 = Convert.ToString(Math.Round(Station2, 0) - Math.Round(Station1, 0)) + Suff;

                                System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                string Block_name = Convert.ToString(dt_compiled.Rows[i][14]);

                                if (BlockTable_data1.Has(Block_name) == false)
                                {
                                    MessageBox.Show("the block name you specified does not belong to the current drawing\r\n" + Block_name);
                                    _AGEN_mainform.tpage_processing.Hide();
                                    set_enable_true();
                                    Ag.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                string suffix1 = "";
                                if (_AGEN_mainform.units_of_measurement == "f") suffix1 = "'";
                                Colectie_nume_atribute.Add("STA1");
                                Colectie_valori.Add(S1);
                                Colectie_nume_atribute.Add("STA2");
                                Colectie_valori.Add(S2);
                                Colectie_nume_atribute.Add("STA11");
                                Colectie_valori.Add(S1);
                                Colectie_nume_atribute.Add("STA21");
                                Colectie_valori.Add(S2);


                                Colectie_nume_atribute.Add("LEN");
                                if (_AGEN_mainform.COUNTRY == "CANADA")
                                {
                                    Colectie_valori.Add(Functions.Get_String_Rounded(Station2_label - Station1_label, _AGEN_mainform.round1) + suffix1);
                                }
                                else
                                {
                                    Colectie_valori.Add(Functions.Get_String_Rounded(Station2 - Station1, _AGEN_mainform.round1) + suffix1);

                                }

                                Colectie_nume_atribute.Add("LENGTH");
                                if (_AGEN_mainform.COUNTRY == "CANADA")
                                {
                                    Colectie_valori.Add(Functions.Get_String_Rounded(Station2_label - Station1_label, _AGEN_mainform.round1) + suffix1);
                                }
                                else
                                {
                                    Colectie_valori.Add(Functions.Get_String_Rounded(Station2 - Station1, _AGEN_mainform.round1) + suffix1);

                                }



                                for (int k = 15; k < dt_compiled.Columns.Count - 1; ++k)
                                {
                                    Colectie_nume_atribute.Add(dt_compiled.Columns[k].ColumnName);
                                    string VAL = "";
                                    if (dt_compiled.Rows[i][k] != DBNull.Value)
                                    {
                                        VAL = Convert.ToString(dt_compiled.Rows[i][k]);
                                    }
                                    Colectie_valori.Add(VAL);
                                }

                                double strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                string visib1 = "";
                                if (dt_compiled.Rows[i][dt_compiled.Columns.Count - 1] != DBNull.Value)
                                {
                                    visib1 = Convert.ToString(dt_compiled.Rows[i][dt_compiled.Columns.Count - 1]);
                                }

                                if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                {
                                    BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", Block_name, InsPt, 1, 0, lname1, Colectie_nume_atribute, Colectie_valori);

                                    #region mat count
                                    int indexcount = -1;
                                    for (int r = 0; r < dtmc.Rows.Count; ++r)
                                    {
                                        string dwg1 = Convert.ToString(dtmc.Rows[r][0]);
                                        if (dwg1 == dwg_name)
                                        {
                                            indexcount = r;
                                            r = dtmc.Rows.Count;
                                        }

                                    }

                                    if (indexcount >= 0)
                                    {
                                        if (Colectie_nume_atribute.Contains("LEN") == true && Colectie_nume_atribute.Contains("MAT") == true)
                                        {
                                            string MAT = Convert.ToString(Colectie_valori[Colectie_nume_atribute.IndexOf("MAT")]);
                                            string LEN = Convert.ToString(Colectie_valori[Colectie_nume_atribute.IndexOf("LEN")]);
                                            string LEN1 = LEN.Replace("'", "").Replace("m", "");

                                            if (MAT != "")
                                            {
                                                if (dtmc.Columns.Contains(MAT) == false)
                                                {
                                                    dtmc.Columns.Add(MAT, typeof(double));
                                                }

                                                double nr = 0;
                                                if (dtmc.Rows[indexcount][MAT] != DBNull.Value)
                                                {
                                                    nr = Convert.ToDouble(dtmc.Rows[indexcount][MAT]);
                                                }


                                                if (Functions.IsNumeric(LEN1) == true)
                                                {
                                                    dtmc.Rows[indexcount][MAT] = nr + Convert.ToDouble(LEN1);
                                                }
                                            }
                                        }
                                    }
                                    #endregion


                                    Functions.Stretch_block(Block1, "Distance1", strech1);
                                    if (visib1 != "") Functions.set_block_visibility(Block1, visib1);



                                    if (_AGEN_mainform.dt_mat_pt != null && _AGEN_mainform.dt_mat_pt.Rows.Count > 0)
                                    {
                                        double x1 = InsPt.X;
                                        double x2 = x1 + lr * strech1;
                                        double y = InsPt.Y;

                                        System.Data.DataTable dt1 = new System.Data.DataTable();
                                        dt1.Columns.Add("blockname", typeof(string));
                                        dt1.Columns.Add("x", typeof(double));
                                        dt1.Columns.Add("y", typeof(double));
                                        dt1.Columns.Add("mat", typeof(string));
                                        dt1.Columns.Add("sta", typeof(string));
                                        dt1.Columns.Add("dnd_zero", typeof(bool));
                                        dt1.Columns.Add("sta2_zero", typeof(bool));
                                        for (int n = 10; n < _AGEN_mainform.dt_mat_pt.Columns.Count; ++n)
                                        {
                                            dt1.Columns.Add(_AGEN_mainform.dt_mat_pt.Columns[n].ColumnName, typeof(string));
                                        }

                                        #region dt_points

                                        for (int k = 0; k < _AGEN_mainform.dt_mat_pt.Rows.Count; ++k)
                                        {
                                            double Station_pt = -1.123;

                                            if (_AGEN_mainform.Project_type == "2D")
                                            {
                                                Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_2DSta]);
                                            }
                                            else
                                            {
                                                Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_3DSta]);
                                            }

                                            double Station_labeled = Functions.Station_equation_ofV2(Station_pt, _AGEN_mainform.dt_station_equation);

                                            if (_AGEN_mainform.COUNTRY == "CANADA" &&
                                                _AGEN_mainform.dt_mat_pt.Columns.Contains("MeasuredCanada") &&
                                                _AGEN_mainform.dt_mat_pt.Rows[k]["MeasuredCanada"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["MeasuredCanada"])) == true)
                                            {
                                                Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["MeasuredCanada"]);
                                            }
                                            else
                                            {
                                                if (_AGEN_mainform.dt_mat_pt.Rows[k]["X"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["X"])) == true &&
                                                       _AGEN_mainform.dt_mat_pt.Rows[k]["Y"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["Y"])) == true)
                                                {
                                                    double X_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["X"]);
                                                    double Y_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["Y"]);

                                                    Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(X_pt, Y_pt, poly2d.Elevation), Vector3d.ZAxis, false);
                                                    Station_pt = poly2d.GetDistAtPoint(point_on_poly2D1);

                                                    if (_AGEN_mainform.Project_type == "3D")
                                                    {
                                                        double param1 = poly2d.GetParameterAtPoint(point_on_poly2D1);
                                                        if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;

                                                        Station_pt = poly3d.GetDistanceAtParameter(param1);
                                                    }




                                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                                    {
                                                        double d2d1 = poly2d.GetDistAtPoint(point_on_poly2D1);
                                                        double b1 = -1.23456;
                                                        Station_labeled = Functions.get_stationCSF_from_point(poly2d, point_on_poly2D1, d2d1, _AGEN_mainform.dt_centerline, ref b1);
                                                    }
                                                    else
                                                    {
                                                        Station_labeled = Functions.Station_equation_ofV2(Station_pt, _AGEN_mainform.dt_station_equation);

                                                    }

                                                }
                                            }



                                            if (Math.Round(Station_pt, 2) == 0 && Math.Round(Station_pt, 2) == Math.Round(Station1, 2) && Math.Round(Station_pt, 2) <= Math.Round(Station2, 2))
                                            {


                                                string material1 = "";
                                                if (_AGEN_mainform.dt_mat_pt.Rows[k][1] != DBNull.Value)
                                                {
                                                    material1 = Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k][1]);
                                                    if (material1 != null)
                                                    {
                                                        dt1.Rows.Add();
                                                        dt1.Rows[dt1.Rows.Count - 1][0] = _AGEN_mainform.dt_mat_pt.Rows[k][9];

                                                        double xp = x1 + lr * (Station_pt - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                        dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                        dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                        dt1.Rows[dt1.Rows.Count - 1][3] = material1;

                                                        string dispsta = Functions.Get_chainage_from_double(Station_labeled, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                        dt1.Rows[dt1.Rows.Count - 1][4] = dispsta;

                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station1, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                        }


                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station2, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = false;
                                                        }

                                                        int idx_pt = 7;
                                                        for (int n = 10; n < _AGEN_mainform.dt_mat_pt.Columns.Count; ++n)
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][idx_pt] = _AGEN_mainform.dt_mat_pt.Rows[k][n];
                                                            ++idx_pt;
                                                        }
                                                    }
                                                }
                                            }

                                            else if (Math.Round(Station_pt, 2) > Math.Round(Station1, 2) && Math.Round(Station_pt, 2) <= Math.Round(Station2, 2))
                                            {
                                                string material1 = "";
                                                if (_AGEN_mainform.dt_mat_pt.Rows[k][1] != DBNull.Value)
                                                {
                                                    material1 = Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k][1]);
                                                    if (material1 != null)
                                                    {
                                                        dt1.Rows.Add();
                                                        dt1.Rows[dt1.Rows.Count - 1][0] = _AGEN_mainform.dt_mat_pt.Rows[k][9];

                                                        double xp = x1 + lr * (Station_pt - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                        dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                        dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                        dt1.Rows[dt1.Rows.Count - 1][3] = material1;

                                                        string dispsta = Functions.Get_chainage_from_double(Station_labeled, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                        dt1.Rows[dt1.Rows.Count - 1][4] = dispsta;

                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station1, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                        }

                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station2, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = false;
                                                        }

                                                        int idx_pt = 7;
                                                        for (int n = 10; n < _AGEN_mainform.dt_mat_pt.Columns.Count; ++n)
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][idx_pt] = _AGEN_mainform.dt_mat_pt.Rows[k][n];
                                                            ++idx_pt;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (dt1.Rows.Count > 0)
                                        {
                                            double xp1 = x1;
                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xi - xp1) < 1 * min_spacing)
                                                {
                                                    xp1 = xp1 + lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp1;


                                                }
                                                else
                                                {
                                                    xp1 = xi;
                                                }

                                                if ((bool)dt1.Rows[m][6] == true)
                                                {
                                                    dt1.Rows[m][1] = x2;
                                                }
                                            }

                                            double xp2 = x2;
                                            for (int m = dt1.Rows.Count - 1; m >= 0; --m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xp2 - xi) < 1 * min_spacing)
                                                {
                                                    xp2 = xp2 - lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp2;
                                                }
                                                else
                                                {
                                                    xp2 = xi;
                                                }
                                                if ((bool)dt1.Rows[m][6] == true)
                                                {
                                                    dt1.Rows[m][1] = x2;
                                                }
                                            }


                                            List<string> lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(Block_name);

                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                string bl = Convert.ToString(dt1.Rows[m][0]);
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double yi = Convert.ToDouble(dt1.Rows[m][2]);
                                                string mati = Convert.ToString(dt1.Rows[m][3]);
                                                string stai = Convert.ToString(dt1.Rows[m][4]);

                                                lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(bl);

                                                System.Collections.Specialized.StringCollection Colectie_nume_atribute_l = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection Colectie_valori_l = new System.Collections.Specialized.StringCollection();
                                                Colectie_nume_atribute_l.Add("STA");
                                                Colectie_valori_l.Add(stai);
                                                Colectie_nume_atribute_l.Add("STA1");
                                                Colectie_valori_l.Add(stai);
                                                Colectie_nume_atribute_l.Add("MAT");
                                                Colectie_valori_l.Add(mati);

                                                for (int n = 7; n < dt1.Columns.Count - 1; ++n)
                                                {
                                                    Colectie_nume_atribute_l.Add(dt1.Columns[n].ColumnName);
                                                    string val = "";
                                                    if (dt1.Rows[m][n] != DBNull.Value)
                                                    {
                                                        val = Convert.ToString(dt1.Rows[m][n]);
                                                    }
                                                    Colectie_valori_l.Add(val);
                                                }


                                                BlockReference Block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                    bl, new Point3d(xi, yi, 0),
                                                                        1, 0, lname1, Colectie_nume_atribute_l, Colectie_valori_l);

                                                #region mat count



                                                if (indexcount >= 0 && Colectie_nume_atribute_l.Contains("MAT") == true)
                                                {
                                                    string MAT = Convert.ToString(Colectie_valori_l[Colectie_nume_atribute_l.IndexOf("MAT")]);


                                                    if (dtmc.Columns.Contains(MAT) == false)
                                                    {
                                                        dtmc.Columns.Add(MAT, typeof(double));
                                                    }

                                                    double nr = 0;
                                                    if (dtmc.Rows[indexcount][MAT] != DBNull.Value)
                                                    {
                                                        nr = Convert.ToDouble(dtmc.Rows[indexcount][MAT]);
                                                    }
                                                    dtmc.Rows[indexcount][MAT] = ++nr;


                                                }
                                                #endregion


                                                string visib2 = "";
                                                if (dt1.Rows[m][dt1.Columns.Count - 1] != DBNull.Value)
                                                {
                                                    visib2 = Convert.ToString(dt1.Rows[m][dt1.Columns.Count - 1]);
                                                }

                                                if (visib2 != "") Functions.set_block_visibility(Block2, visib2);

                                            }

                                        }
                                        #endregion

                                    }


                                    if (Data_table_compiled_extra != null && Data_table_compiled_extra.Rows.Count > 0)
                                    {


                                        double x1 = InsPt.X;
                                        double x2 = x1 + lr * strech1;
                                        double y = InsPt.Y;

                                        System.Data.DataTable dt1 = new System.Data.DataTable();
                                        dt1.Columns.Add("blockname", typeof(string));
                                        dt1.Columns.Add("x1", typeof(double));
                                        dt1.Columns.Add("y1", typeof(double));
                                        dt1.Columns.Add("STA1", typeof(string));
                                        dt1.Columns.Add("STA2", typeof(string));
                                        dt1.Columns.Add("dnd_zero", typeof(bool));

                                        for (int n = 15; n < Data_table_compiled_extra.Columns.Count; ++n)
                                        {
                                            string nume_col = Data_table_compiled_extra.Columns[n].ColumnName;

                                            dt1.Columns.Add(nume_col, typeof(string));
                                        }

                                        #region dt_mat_lin_extra

                                        for (int k = 0; k < Data_table_compiled_extra.Rows.Count; ++k)
                                        {

                                            debug = "k = " + k.ToString() + "\r\ni = " + i.ToString() + "\r\nblock is not present?";


                                            double Station_pt1 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta1]);
                                            double Station_pt2 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta2]);

                                            double Station_labeled1 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta1_label]);
                                            double Station_labeled2 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta2_label]);

                                            if (Math.Round(Station_pt1, 2) == 0 && Math.Round(Station_pt1, 2) == Math.Round(Station1, 2) && Math.Round(Station_pt1, 2) <= Math.Round(Station2, 2))
                                            {
                                                if (Data_table_compiled_extra.Rows[k][1] != DBNull.Value)
                                                {
                                                    string nume_block = Convert.ToString(Data_table_compiled_extra.Rows[k][14]);
                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][0] = Data_table_compiled_extra.Rows[k][14];
                                                    double xp = x1 + lr * (Station_pt1 - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                    string dispsta1 = Functions.Get_chainage_from_double(Station_labeled1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][3] = dispsta1;
                                                    string dispsta2 = Functions.Get_chainage_from_double(Station_labeled2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][4] = dispsta2;
                                                    if (Math.Round(Station_pt1, 2) == Math.Round(Station1, 2))
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                    }
                                                    else
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                    }

                                                    int idx_pt = 6;
                                                    for (int n = 15; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][idx_pt] = Data_table_compiled_extra.Rows[k][n];
                                                        ++idx_pt;
                                                    }
                                                }
                                            }
                                            else if (Math.Round(Station_pt1, 2) >= Math.Round(Station1, 2) && Math.Round(Station_pt1, 2) <= Math.Round(Station2, 2) && Math.Round(Station_pt2, 2) <= Math.Round(M2, 2))
                                            {
                                                if (Data_table_compiled_extra.Rows[k][1] != DBNull.Value)
                                                {
                                                    string nume_block = Convert.ToString(Data_table_compiled_extra.Rows[k][14]);
                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][0] = Data_table_compiled_extra.Rows[k][14];
                                                    double xp = x1 + lr * (Station_pt1 - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                    string dispsta1 = Functions.Get_chainage_from_double(Station_labeled1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][3] = dispsta1;
                                                    string dispsta2 = Functions.Get_chainage_from_double(Station_labeled2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][4] = dispsta2;
                                                    if (Math.Round(Station_pt1, 2) == Math.Round(Station1, 2))
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                    }
                                                    else
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                    }
                                                    int idx_pt = 6;
                                                    for (int n = 15; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][idx_pt] = Data_table_compiled_extra.Rows[k][n];
                                                        ++idx_pt;
                                                    }
                                                }
                                            }
                                        }

                                        if (dt1.Rows.Count > 0)
                                        {
                                            double xp1 = x1;
                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xi - xp1) < 1 * min_spacing)
                                                {
                                                    xp1 = xp1 + lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp1;
                                                }
                                                else
                                                {
                                                    xp1 = xi;
                                                }
                                            }

                                            double xp2 = x2;

                                            for (int m = dt1.Rows.Count - 1; m >= 0; --m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);

                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xp2 - xi) < 1 * min_spacing)
                                                {
                                                    xp2 = xp2 - lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp2;
                                                }
                                                else
                                                {
                                                    xp2 = xi;
                                                }
                                            }



                                            List<string> lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(Block_name);


                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                string bl = Convert.ToString(dt1.Rows[m][0]);
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double yi = Convert.ToDouble(dt1.Rows[m][2]);
                                                string ss1 = Convert.ToString(dt1.Rows[m][3]);
                                                string ss2 = Convert.ToString(dt1.Rows[m][4]);

                                                lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(bl);

                                                System.Collections.Specialized.StringCollection Colectie_nume_atribute_l = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection Colectie_valori_l = new System.Collections.Specialized.StringCollection();
                                                Colectie_nume_atribute_l.Add("STA1");
                                                Colectie_valori_l.Add(ss1);
                                                Colectie_nume_atribute_l.Add("STA11");
                                                Colectie_valori_l.Add(ss1);
                                                Colectie_nume_atribute_l.Add("STA2");
                                                Colectie_valori_l.Add(ss2);
                                                Colectie_nume_atribute_l.Add("STA21");
                                                Colectie_valori_l.Add(ss2);

                                                for (int n = 6; n < dt1.Columns.Count - 1; ++n)
                                                {
                                                    Colectie_nume_atribute_l.Add(dt1.Columns[n].ColumnName);
                                                    string val = "";
                                                    if (dt1.Rows[m][n] != DBNull.Value)
                                                    {
                                                        val = Convert.ToString(dt1.Rows[m][n]);
                                                    }
                                                    Colectie_valori_l.Add(val);
                                                }
                                                BlockReference Block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                    bl, new Point3d(xi, yi, 0), 1, 0, lname1, Colectie_nume_atribute_l, Colectie_valori_l);


                                                #region mat count
                                                if (indexcount >= 0)
                                                {
                                                    if (Colectie_nume_atribute_l.Contains("LEN") == true && Colectie_nume_atribute_l.Contains("MAT") == true)
                                                    {
                                                        string MAT = Convert.ToString(Colectie_valori_l[Colectie_nume_atribute_l.IndexOf("MAT")]);
                                                        string LEN = Convert.ToString(Colectie_valori_l[Colectie_nume_atribute_l.IndexOf("LEN")]);
                                                        string LEN1 = LEN.Replace("'", "").Replace("m", "");
                                                        if (dtmc.Columns.Contains(MAT) == false)
                                                        {
                                                            dtmc.Columns.Add(MAT, typeof(double));
                                                        }
                                                        double nr = 0;
                                                        if (dtmc.Rows[indexcount][MAT] != DBNull.Value)
                                                        {
                                                            nr = Convert.ToDouble(dtmc.Rows[indexcount][MAT]);
                                                        }
                                                        if (Functions.IsNumeric(LEN1) == true)
                                                        {
                                                            dtmc.Rows[indexcount][MAT] = nr + Convert.ToDouble(LEN1);
                                                        }
                                                    }
                                                }
                                                #endregion


                                                string visib2 = "";
                                                if (dt1.Rows[m][dt1.Columns.Count - 1] != DBNull.Value)
                                                {
                                                    visib2 = Convert.ToString(dt1.Rows[m][dt1.Columns.Count - 1]);
                                                }

                                                if (visib2 != "") Functions.set_block_visibility(Block2, visib2);
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }


                        //Alignment_generator.Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_compiled);

                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false)
                        {
                            poly3d.Erase();
                        }


                        Trans1.Commit();

                        dataGridView_materials.DataSource = dtmc;
                        dataGridView_materials.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                        dataGridView_materials.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_materials.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        dataGridView_materials.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_materials.DefaultCellStyle.ForeColor = Color.White;
                        dataGridView_materials.EnableHeadersVisualStyles = false;
                        _AGEN_mainform.dt_mat_lin = null;
                        _AGEN_mainform.dt_mat_lin_extra = null;
                        _AGEN_mainform.dt_mat_pt = null;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + debug);
            }

            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();

            Ag.WindowState = FormWindowState.Normal;
        }


        private void button_open_mat_linear_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (System.IO.Directory.Exists(ProjF) == true)
                {

                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }

                    string fis_mat_lin = ProjF + _AGEN_mainform.mat_linear_excel_name;
                    string fis_mat = ProjF + _AGEN_mainform.materials_excel_name;
                    if (System.IO.File.Exists(fis_mat) == false && System.IO.File.Exists(fis_mat_lin) == false)
                    {
                        set_enable_false();

                        MessageBox.Show("the material linear data file does not exist");
                        return;
                    }

                    Microsoft.Office.Interop.Excel.Application Excel1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    if (Excel1 == null)
                    {
                        MessageBox.Show("PROBLEM WITH EXCEL!");
                        return;
                    }
                    Excel1.Visible = true;
                    if (System.IO.File.Exists(fis_mat) == true)
                    {
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fis_mat);
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fis_mat_lin);
                    }

                }
                else
                {
                    _AGEN_mainform.tpage_processing.Hide();

                    MessageBox.Show("the project folder does not exist");
                }



            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();


        }

        private void button_open_mat_linear_extra_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (System.IO.Directory.Exists(ProjF) == true)
                {

                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }

                    string fis_mat_lin_extra = ProjF + _AGEN_mainform.mat_linear_extra_excel_name;
                    string fis_mat = ProjF + _AGEN_mainform.materials_excel_name;
                    if (System.IO.File.Exists(fis_mat) == false && System.IO.File.Exists(fis_mat_lin_extra) == false)
                    {

                        set_enable_false();

                        MessageBox.Show("the material linear extra data file does not exist");
                        return;
                    }

                    Microsoft.Office.Interop.Excel.Application Excel1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    if (Excel1 == null)
                    {
                        MessageBox.Show("PROBLEM WITH EXCEL!");
                        return;
                    }
                    Excel1.Visible = true;
                    if (System.IO.File.Exists(fis_mat) == true)
                    {
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fis_mat);
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fis_mat_lin_extra);
                    }
                }
                else
                {
                    _AGEN_mainform.tpage_processing.Hide();

                    MessageBox.Show("the project folder does not exist");
                }



            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();


        }

        private void button_open_mat_points_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (System.IO.Directory.Exists(ProjF) == true)
                {

                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }

                    string fis_mat_pt = ProjF + _AGEN_mainform.mat_points_excel_name;
                    string fis_mat = ProjF + _AGEN_mainform.materials_excel_name;
                    if (System.IO.File.Exists(fis_mat) == false && System.IO.File.Exists(fis_mat_pt) == false)
                    {
                        set_enable_false();

                        MessageBox.Show("the material points data file does not exist");
                        return;
                    }

                    Microsoft.Office.Interop.Excel.Application Excel1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }

                    if (Excel1 == null)
                    {
                        MessageBox.Show("PROBLEM WITH EXCEL!");
                        return;
                    }
                    Excel1.Visible = true;
                    if (System.IO.File.Exists(fis_mat) == true)
                    {
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fis_mat);
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fis_mat_pt);
                    }
                }
                else
                {
                    _AGEN_mainform.tpage_processing.Hide();

                    MessageBox.Show("the project folder does not exist");
                }



            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();


        }

        private void button_show_mat_counts_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();
            _AGEN_mainform.tpage_layer_alias.Hide();
            _AGEN_mainform.tpage_crossing_scan.Hide();
            _AGEN_mainform.tpage_crossing_draw.Hide();
            _AGEN_mainform.tpage_profilescan.Hide();
            _AGEN_mainform.tpage_profdraw.Hide();

            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_mat_count.Show();
        }

        public void Fill_combobox_segments()
        {
            comboBox_segment_name.Items.Clear();
            if (_AGEN_mainform.lista_segments != null && _AGEN_mainform.lista_segments.Count > 0)
            {
                try
                {
                    for (int i = 0; i < _AGEN_mainform.lista_segments.Count; ++i)
                    {
                        comboBox_segment_name.Items.Add(_AGEN_mainform.lista_segments[i]);
                    }
                    comboBox_segment_name.SelectedIndex = 0;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void set_combobox_segment_name()
        {
            comboBox_segment_name.SelectedIndex = comboBox_segment_name.Items.IndexOf(_AGEN_mainform.current_segment);
        }

        private void ComboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            _AGEN_mainform.current_segment = comboBox_segment_name.Text;
            _AGEN_mainform.tpage_setup.set_combobox_segment_name();


        }

        private void Button_generate_mat_spreadsheets_with_headers_only_Click(object sender, EventArgs e)
        {
            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.materials_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.materials_excel_name + " file");
                return;
            }



            try
            {
                set_enable_false();
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_mat = ProjFolder + _AGEN_mainform.materials_excel_name;
                    System.Data.DataTable dt_lin = null;
                    System.Data.DataTable dt_extra = null;
                    System.Data.DataTable dt_points = null;

                    if (System.IO.File.Exists(fisier_mat) == false)
                    {
                        dt_lin = Functions.Creaza_dt_mat_lin_structure_for_new_file();
                        dt_extra = Functions.Creaza_dt_mat_lin_structure_for_new_file();
                        dt_points = Functions.Creaza_dt_mat_point_structure_for_new_file();
                    }



                    creaza_headers_for_mat_files(fisier_mat, dt_lin, dt_points, dt_extra);

                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void creaza_headers_for_mat_files(string fis_mat, System.Data.DataTable dt_lin, System.Data.DataTable dt_pts, System.Data.DataTable dt_extra)
        {

            if (dt_lin != null || dt_extra != null || dt_pts != null)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;


                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                try
                {
                    if (dt_lin != null)
                    {
                        Workbook1 = Excel1.Workbooks.Add();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                        Microsoft.Office.Interop.Excel.Worksheet W2 = Workbook1.Worksheets.Add();
                        Microsoft.Office.Interop.Excel.Worksheet W3 = Workbook1.Worksheets.Add();
                        Functions.Create_header_material_linear_file(W3, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), _AGEN_mainform.current_segment, _AGEN_mainform.version, dt_lin);
                        Functions.Create_header_material_points_file(W2, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), _AGEN_mainform.current_segment, _AGEN_mainform.version, dt_pts);
                        Functions.Create_header_material_linear_extra_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), _AGEN_mainform.current_segment, _AGEN_mainform.version, dt_extra);
                        W3.Name = "LINEAR";
                        W2.Name = "POINTS";
                        W1.Name = "LINEAR EXTRA";


                        Workbook1.SaveAs(fis_mat);
                        Workbook1.Close();
                    }


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




        }


        private void draw_mat_band_USA_2D()
        {

            System.Data.DataTable dtmc = new System.Data.DataTable();
            dtmc.Columns.Add("band", typeof(string));
            string debug = "00";

            string lnp = "Agen_no_plot_mat";
            double min_spacing = 0.4;
            if (_AGEN_mainform.COUNTRY == "CANADA")
            {
                min_spacing = 25;
            }

            if (Functions.IsNumeric(textBox_spacing.Text) == true)
            {
                min_spacing = Math.Abs(Convert.ToDouble(textBox_spacing.Text));
            }



            int lr = 1;
            if (_AGEN_mainform.Left_to_Right == false) lr = -1;

            Functions.Kill_excel();


            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }

            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }

            if (_AGEN_mainform.dt_mat_lin == null || _AGEN_mainform.dt_mat_lin.Rows.Count == 0)
            {
                MessageBox.Show("No linear material data found");
                return;
            }

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            //Ag.WindowState = FormWindowState.Minimized;

            if (_AGEN_mainform.Vw_mat_height == 0)
            {
                MessageBox.Show("you did not specified viewport material information");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();

                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();


                _AGEN_mainform.tpage_viewport_settings.Show();

                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }

            set_enable_false();

            _AGEN_mainform.tpage_processing.Show();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }



                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                if (System.IO.File.Exists(fisier_si) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the sheet index data file does not exist");
                    return;
                }

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the sheet index data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }


                _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);

                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("no centerline");
                    return;
                }

            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            if (_AGEN_mainform.dt_mat_lin.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the linear material file does not have any data");
                return;
            }

            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the sheet index file does not have any data");
                return;
            }



            string Sta1 = "Sta1";
            string Sta2 = "Sta2";
            string Sta1_label = "Sta1_label";
            string Sta2_label = "Sta2_label";
            string Match1 = "M1";
            string Match2 = "M2";
            string mat_atr = "material_for_align";
            string Pageno = "Page";
            string Rect_len = "RectangleML";
            string stretch_val = "StrechVal";
            string BandL = "BandL";
            string DeltaX_col = "DeltaX";
            string stretch_val_orig = "StrechValoriginal";


            System.Data.DataTable dt_compiled = new System.Data.DataTable();
            dt_compiled.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));       //0
            dt_compiled.Columns.Add(Sta1, typeof(double));                              //1
            dt_compiled.Columns.Add(Sta2, typeof(double));                              //2
            dt_compiled.Columns.Add(mat_atr, typeof(string));                           //3
            dt_compiled.Columns.Add(Pageno, typeof(int));                               //4
            dt_compiled.Columns.Add(Rect_len, typeof(double));                          //5
            dt_compiled.Columns.Add(BandL, typeof(double));                             //6
            dt_compiled.Columns.Add(DeltaX_col, typeof(double));                        //7
            dt_compiled.Columns.Add(Match1, typeof(double));                            //8
            dt_compiled.Columns.Add(Match2, typeof(double));                            //9
            dt_compiled.Columns.Add(stretch_val, typeof(double));                       //10
            dt_compiled.Columns.Add(stretch_val_orig, typeof(double));                  //11
            dt_compiled.Columns.Add(Sta1_label, typeof(double));                        //12
            dt_compiled.Columns.Add(Sta2_label, typeof(double));                        //13

            for (int n = 15; n < _AGEN_mainform.dt_mat_lin.Columns.Count; ++n)
            {
                dt_compiled.Columns.Add(_AGEN_mainform.dt_mat_lin.Columns[n].ColumnName, typeof(string));
            }

            #region Data_table_compiled_extra
            System.Data.DataTable Data_table_compiled_extra = null;
            if (_AGEN_mainform.dt_mat_lin_extra != null && _AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
            {
                Data_table_compiled_extra = new System.Data.DataTable();
                Data_table_compiled_extra.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));     //0
                Data_table_compiled_extra.Columns.Add(Sta1, typeof(double));                            //1
                Data_table_compiled_extra.Columns.Add(Sta2, typeof(double));                            //2
                Data_table_compiled_extra.Columns.Add(mat_atr, typeof(string));                         //3
                Data_table_compiled_extra.Columns.Add(Pageno, typeof(int));                             //4
                Data_table_compiled_extra.Columns.Add(Rect_len, typeof(double));                        //5
                Data_table_compiled_extra.Columns.Add(BandL, typeof(double));                           //6
                Data_table_compiled_extra.Columns.Add(DeltaX_col, typeof(double));                      //7
                Data_table_compiled_extra.Columns.Add(Match1, typeof(double));                          //8
                Data_table_compiled_extra.Columns.Add(Match2, typeof(double));                          //9
                Data_table_compiled_extra.Columns.Add(stretch_val, typeof(double));                     //10
                Data_table_compiled_extra.Columns.Add(stretch_val_orig, typeof(double));                //11
                Data_table_compiled_extra.Columns.Add(Sta1_label, typeof(double));                      //12
                Data_table_compiled_extra.Columns.Add(Sta2_label, typeof(double));                      //13

                for (int n = 15; n < _AGEN_mainform.dt_mat_lin_extra.Columns.Count; ++n)
                {
                    Data_table_compiled_extra.Columns.Add(_AGEN_mainform.dt_mat_lin_extra.Columns[n].ColumnName, typeof(string));
                }
            }
            #endregion





            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;



                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                        #region USA
                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.COUNTRY == "USA")
                        {
                            if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                            {
                                if (_AGEN_mainform.dt_station_equation.Columns.Contains("measured") == false)
                                {
                                    _AGEN_mainform.dt_station_equation.Columns.Add("measured", typeof(double));
                                }

                                for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                    {
                                        double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                        double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);


                                        Point3d pt_on_2d = poly2d.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);

                                        _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = poly2d.GetDistAtPoint(pt_on_2d);

                                    }
                                }
                            }

                        }
                        #endregion

                        Functions.Creaza_layer(lnp, 30, false);
                        string lname1 = "Agen_band_mat";
                        Functions.Creaza_layer(lname1, 7, true);

                        List<int> lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_indexes_for_generation(_AGEN_mainform.Point0_mat, _AGEN_mainform.Band_Separation, lnp);

                        #region adauga breaks for matchlines dt_mat_lin
                        for (int i = 0; i < _AGEN_mainform.dt_mat_lin.Rows.Count; ++i)
                        {

                            int m_start = 0;
                            bool Boolean_go_to_check_s1_s2 = false;
                            double Station1 = -1.123;
                            double Station2 = -1.123;
                            string Material1 = "";
                            double Station1_labeled = -1.123;
                            double Station2_labeled = -1.123;

                            if (_AGEN_mainform.Project_type == "2D")
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta1] != DBNull.Value &&
                                    _AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta2] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta1]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_2DSta2]);
                                }
                            }
                            else
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta1] != DBNull.Value && _AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta2] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta1]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_3DSta2]);
                                }
                            }




                            Station1_labeled = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                            Station2_labeled = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);

                            if (Station1 != -1.123 && Station2 != -1.123)
                            {
                                if (_AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_Material] != DBNull.Value)
                                {
                                    Material1 = _AGEN_mainform.dt_mat_lin.Rows[i][_AGEN_mainform.Col_Material].ToString();
                                }
                            L123:
                                for (int j = m_start; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                {
                                    if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2] != DBNull.Value)
                                    {
                                        double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1]);
                                        double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2]);


                                        if (M2 <= M1)
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("End Station is smaller than Start Station on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                            return;
                                        }
                                        if (M2 > poly2d.Length)
                                        {
                                            if (Math.Abs(M2 - poly2d.Length) < 0.99)
                                            {
                                                M2 = poly2d.Length;
                                            }
                                            else
                                            {
                                                _AGEN_mainform.tpage_processing.Hide();
                                                set_enable_true();
                                                MessageBox.Show("End Station is bigger than poly length on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                return;
                                            }
                                        }
                                        if (M1 >= poly2d.Length) M1 = poly2d.Length - 0.0001;
                                        if (M2 >= poly2d.Length) M2 = poly2d.Length - 0.0001;
                                        Point3d pm1 = poly2d.GetPointAtDist(M1);
                                        Point3d pm2 = poly2d.GetPointAtDist(M2);
                                        pm1 = new Point3d(pm1.X, pm1.Y, 0);
                                        pm2 = new Point3d(pm2.X, pm2.Y, 0);
                                        Autodesk.AutoCAD.DatabaseServices.Line Linie_M1_M2 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pm1.X, pm1.Y, 0), new Point3d(pm2.X, pm2.Y, 0));
                                        if (Boolean_go_to_check_s1_s2 == true)
                                        {
                                            if (Math.Round(Station1, 0) == Math.Round(Station2, 0))
                                            {
                                                goto LS12end;
                                            }
                                            goto LS1S2;
                                        }

                                        if (Math.Round(M1, 2) <= Math.Round(Station1, 2) && Math.Round(M2, 2) <= Math.Round(Station2, 2) && Math.Round(M1, 2) <= Math.Round(Station2, 2) && Math.Round(M2, 2) > Math.Round(Station1, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            else
                                            {
                                                double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            Station1 = M2;
                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }

                                        if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = Station2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = Station2_labeled;
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }

                                    LS1S2:
                                        if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = Station2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = Station2_labeled;
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }
                                        else if (Math.Round(Station1, 2) < Math.Round(M2, 2) && Math.Round(Station1, 2) >= Math.Round(M1, 2))
                                        {
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows.Add();
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1] = Station1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][mat_atr] = Material1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match1] = M1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Match2] = M2;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            else
                                            {
                                                double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][Sta2_label] = label2;
                                                Station1_labeled = label2;
                                            }
                                            int idx_lin = 15;
                                            for (int n = 14; n < dt_compiled.Columns.Count; ++n)
                                            {
                                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin.Rows[i][idx_lin];
                                                ++idx_lin;
                                            }
                                            Station1 = M2;
                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }
                                    }
                                }
                            LS12end:
                                string xx = "";
                            }
                        }
                        #endregion




                        #region adauga breaks for matchlines dt_mat_lin_extra_extra
                        if (_AGEN_mainform.dt_mat_lin_extra != null && _AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
                        {
                            for (int i = 0; i < _AGEN_mainform.dt_mat_lin_extra.Rows.Count; ++i)
                            {
                                int m_start = 0;
                                bool Boolean_go_to_check_s1_s2 = false;
                                double Station1 = -1.123;
                                double Station2 = -1.123;
                                string Material1 = "";
                                double Station1_labeled = -1.123;
                                double Station2_labeled = -1.123;

                                if (_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta1] != DBNull.Value && _AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta2] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta1]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_2DSta2]);
                                }






                                if (_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_Beg"])) == true &&
                                    _AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_Beg"])) == true &&
                                    _AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_End"])) == true &&
                                    _AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_End"])) == true)
                                {
                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_Beg"]);
                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_Beg"]);
                                    double x2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["X_End"]);
                                    double y2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[i]["Y_End"]);
                                    Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);
                                    Point3d point_on_poly2D2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, false);

                                    Station1 = poly2d.GetDistAtPoint(point_on_poly2D1);
                                    Station2 = poly2d.GetDistAtPoint(point_on_poly2D2);


                                }


                                Station1_labeled = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                Station2_labeled = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);

                                if (Station1 != -1.123 && Station2 != -1.123)
                                {
                                    if (_AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_Material] != DBNull.Value)
                                    {
                                        Material1 = _AGEN_mainform.dt_mat_lin_extra.Rows[i][_AGEN_mainform.Col_Material].ToString();
                                    }
                                L123:
                                    for (int j = m_start; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                    {
                                        if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2] != DBNull.Value)
                                        {
                                            double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1]);
                                            double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2]);

                                            if (M2 <= M1)
                                            {
                                                _AGEN_mainform.tpage_processing.Hide();
                                                set_enable_true();
                                                MessageBox.Show("End Station is smaller than Start Station on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                return;
                                            }
                                            if (M2 > poly2d.Length)
                                            {
                                                if (Math.Abs(M2 - poly2d.Length) < 0.99)
                                                {
                                                    M2 = poly2d.Length;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.tpage_processing.Hide();
                                                    set_enable_true();
                                                    MessageBox.Show("End Station is bigger than poly length on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                    return;
                                                }
                                            }
                                            if (M1 >= poly2d.Length) M1 = poly2d.Length - 0.0001;
                                            if (M2 >= poly2d.Length) M2 = poly2d.Length - 0.0001;
                                            Point3d pm1 = poly2d.GetPointAtDist(M1);
                                            Point3d pm2 = poly2d.GetPointAtDist(M2);
                                            pm1 = new Point3d(pm1.X, pm1.Y, 0);
                                            pm2 = new Point3d(pm2.X, pm2.Y, 0);
                                            Autodesk.AutoCAD.DatabaseServices.Line Linie_M1_M2 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pm1.X, pm1.Y, 0), new Point3d(pm2.X, pm2.Y, 0));
                                            if (Boolean_go_to_check_s1_s2 == true)
                                            {
                                                if (Math.Round(Station1, 0) == Math.Round(Station2, 0))
                                                {
                                                    goto LS12end;
                                                }
                                                goto LS1S2;
                                            }

                                            if (Math.Round(M1, 2) <= Math.Round(Station1, 2) && Math.Round(M2, 2) <= Math.Round(Station2, 2) && Math.Round(M1, 2) <= Math.Round(Station2, 2) && Math.Round(M2, 2) > Math.Round(Station1, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                else
                                                {
                                                    double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                int idx_lin = 15;

                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);
                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }
                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;

                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {

                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }



                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }


                                                Station1 = M2;
                                                m_start = j + 1;
                                                Boolean_go_to_check_s1_s2 = true;
                                                goto L123;
                                            }

                                            if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                                double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = Station2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = Station2_labeled;

                                                int idx_lin = 15;
                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }


                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;

                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }


                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }

                                                j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                goto LS12end;
                                            }

                                        LS1S2:
                                            if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station2), Vector3d.ZAxis, false);
                                                double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = Station2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = Station2_labeled;
                                                int idx_lin = 15;

                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }


                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;

                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {

                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (Station2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }

                                                j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                goto LS12end;
                                            }
                                            else if (Math.Round(Station1, 2) < Math.Round(M2, 2) && Math.Round(Station1, 2) >= Math.Round(M1, 2))
                                            {
                                                Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(poly2d.GetPointAtDist(Station1), Vector3d.ZAxis, false);
                                                double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                                double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows.Add();
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1] = Station1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][mat_atr] = Material1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Pageno] = j + 1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match1] = M1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Match2] = M2;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][stretch_val_orig] = stretch01;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][DeltaX_col] = deltax1;
                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta1_label] = Station1_labeled;
                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    double label2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                else
                                                {
                                                    double label2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                    Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][Sta2_label] = label2;
                                                    Station1_labeled = label2;
                                                }
                                                int idx_lin = 15;

                                                bool is_buoyancy = false;
                                                double spacing = 0;
                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (idx_lin == 15)
                                                    {
                                                        if (bn.ToUpper() == "SA")
                                                        {
                                                            is_buoyancy = true;
                                                        }
                                                    }

                                                    if (is_buoyancy == true)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "SPACING")
                                                        {
                                                            if (Functions.IsNumeric(bn.Replace(" C/C", "")) == true)
                                                            {
                                                                spacing = Convert.ToDouble(bn.Replace(" C/C", ""));
                                                            }
                                                        }
                                                    }


                                                    ++idx_lin;
                                                }

                                                idx_lin = 15;


                                                for (int n = 14; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                {
                                                    string bn = Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin]);

                                                    if (is_buoyancy == true && spacing > 0)
                                                    {
                                                        if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "NO")
                                                        {
                                                            string numar = Functions.extrage_integer_pozitiv_number_din_text_de_la_inceputul_textului(bn);
                                                            if (Functions.IsNumeric(numar) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }


                                                                bn = bn.Replace(numar, new_no.ToString());
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = bn;

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else if (_AGEN_mainform.dt_mat_lin_extra.Columns[idx_lin].ColumnName.ToUpper() == "QTY")
                                                        {

                                                            if (Functions.IsNumeric(bn) == true)
                                                            {
                                                                int new_no = 0;
                                                                double math_nr = (M2 - Station1) / spacing;
                                                                double up1 = Math.Ceiling(math_nr);
                                                                double down1 = Math.Floor(math_nr);
                                                                double df1 = up1 - math_nr;
                                                                double df2 = math_nr - down1;

                                                                if (df1 <= df2)
                                                                {
                                                                    new_no = 1 + Convert.ToInt32(up1);
                                                                }
                                                                else
                                                                {
                                                                    new_no = Convert.ToInt32(up1);
                                                                }

                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = new_no.ToString();

                                                            }
                                                            else
                                                            {
                                                                Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Data_table_compiled_extra.Rows[Data_table_compiled_extra.Rows.Count - 1][n] = _AGEN_mainform.dt_mat_lin_extra.Rows[i][idx_lin];
                                                    }

                                                    ++idx_lin;
                                                }
                                                Station1 = M2;
                                                m_start = j + 1;
                                                Boolean_go_to_check_s1_s2 = true;
                                                goto L123;
                                            }
                                        }
                                    }
                                LS12end:
                                    string xx = "";
                                }
                            }
                        }

                        #endregion

                        int Pagep = -1;

                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_compiled);

                        if (dt_compiled != null && dt_compiled.Rows.Count > 0)
                        {

                            #region draw red and green rectangles
                            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                            {
                                int Page1 = Convert.ToInt32(dt_compiled.Rows[i][Pageno]);
                                double ml_len = Convert.ToDouble(dt_compiled.Rows[i][Rect_len]);
                                string dwg_name = dt_compiled.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                double strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                if (strech1 < min_spacing)
                                {
                                    double new_stretch = min_spacing;
                                    dt_compiled.Rows[i][stretch_val] = new_stretch;
                                    double Diff = new_stretch - strech1;
                                    for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                    {
                                        int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                        double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                        double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);
                                        if (Page1 == Page2)
                                        {
                                            dt_compiled.Rows[j][BandL] = band_len2 + Diff;
                                            if (i < j)
                                            {
                                                dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                            }
                                        }
                                    }
                                }
                                if (Page1 != Pagep)
                                {
                                    if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                    {

                                        dtmc.Rows.Add();
                                        dtmc.Rows[dtmc.Rows.Count - 1][0] = dwg_name;

                                        Polyline vp_vw1 = new Polyline();
                                        vp_vw1.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_mat.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_mat.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_mat.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_mat.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.Closed = true;
                                        vp_vw1.Layer = lnp;
                                        vp_vw1.ColorIndex = 3;
                                        BTrecord.AppendEntity(vp_vw1);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw1, true);

                                        Polyline vp_vw2 = new Polyline();
                                        vp_vw2.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_mat.X - ml_len / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_mat.X + ml_len / 2, _AGEN_mainform.Point0_mat.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_mat.X + ml_len / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_mat.X - ml_len / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);

                                        vp_vw2.Closed = true;
                                        vp_vw2.Layer = lnp;
                                        vp_vw2.ColorIndex = 1;
                                        BTrecord.AppendEntity(vp_vw2);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw2, true);

                                        MText Band_label = new MText();
                                        Band_label.Contents = dwg_name;
                                        Band_label.TextHeight = _AGEN_mainform.Vw_mat_height / 3;
                                        Band_label.Rotation = 0;
                                        Band_label.Attachment = AttachmentPoint.MiddleLeft;
                                        Band_label.Location = new Point3d(_AGEN_mainform.Point0_mat.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height / 2 - (Page1 - 1) * _AGEN_mainform.Band_Separation, 0);
                                        Band_label.Layer = lnp;
                                        BTrecord.AppendEntity(Band_label);
                                        Trans1.AddNewlyCreatedDBObject(Band_label, true);
                                    }
                                    Pagep = Page1;
                                }
                            }
                            #endregion

                            if (_AGEN_mainform.dt_mat_pt != null)
                            {
                                if (_AGEN_mainform.dt_mat_pt.Rows.Count > 0)
                                {

                                    _AGEN_mainform.dt_mat_pt = Functions.Sort_data_table(_AGEN_mainform.dt_mat_pt, _AGEN_mainform.Col_2DSta);


                                }
                            }

                            if (_AGEN_mainform.dt_mat_lin_extra != null)
                            {
                                if (_AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
                                {

                                    _AGEN_mainform.dt_mat_lin_extra = Functions.Sort_data_table(_AGEN_mainform.dt_mat_lin_extra, "2DSTABEG");


                                }
                            }

                            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                            {
                                int Page1 = Convert.ToInt32(dt_compiled.Rows[i][Pageno]);
                                double Station1 = Convert.ToDouble(dt_compiled.Rows[i][Sta1]);
                                double Station2 = Convert.ToDouble(dt_compiled.Rows[i][Sta2]);
                                double strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                double Stretch_from_pts = 0;
                                double Stretch_from_extra1 = 0;
                                double Stretch_from_extra2 = 0;
                                bool crosing_found_inside = false;
                                bool extra_found_inside1 = false;
                                bool extra_found_inside2 = false;

                                #region dt_points
                                if (_AGEN_mainform.dt_mat_pt != null)
                                {
                                    for (int k = 0; k < _AGEN_mainform.dt_mat_pt.Rows.Count; ++k)
                                    {

                                        double Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_2DSta]);




                                        if (_AGEN_mainform.dt_mat_pt.Rows[k]["X"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["X"])) == true)
                                        {
                                            double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["X"]);
                                            double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["Y"]);
                                            Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);

                                            Station_pt = poly2d.GetDistAtPoint(point_on_poly2D1);
                                        }




                                        if (Station_pt > Station1 && Station_pt <= Station2)
                                        {
                                            Stretch_from_pts = Stretch_from_pts + 1 * min_spacing;
                                            crosing_found_inside = true;
                                        }
                                        else if (Station_pt == 0)
                                        {
                                            if (Station_pt <= Station2)
                                            {
                                                Stretch_from_pts = Stretch_from_pts + 1 * min_spacing;
                                                crosing_found_inside = true;
                                            }
                                        }
                                    }

                                    if (crosing_found_inside == true && strech1 < Stretch_from_pts)
                                    {
                                        dt_compiled.Rows[i][stretch_val] = Stretch_from_pts;
                                        double Diff = Stretch_from_pts - strech1;
                                        for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                        {
                                            int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                            double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                            double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);

                                            if (Page1 == Page2)
                                            {
                                                dt_compiled.Rows[j][BandL] = band_len2 + Diff;

                                                if (i < j)
                                                {
                                                    dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                                }
                                            }
                                        }


                                    }
                                }

                                #endregion

                                #region dt_extra
                                if (_AGEN_mainform.dt_mat_lin_extra != null && _AGEN_mainform.dt_mat_lin_extra.Rows.Count > 0)
                                {
                                    strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                    for (int k = 0; k < _AGEN_mainform.dt_mat_lin_extra.Rows.Count; ++k)
                                    {
                                        double Station_pt1 = -1.123;
                                        double Station_pt2 = -1.123;
                                        if (_AGEN_mainform.Project_type=="2D")
                                        {
                                            Station_pt1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["2DSTABEG"]);
                                        }
                                        else
                                        {
                                            Station_pt1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["3DSTABEG"]);
                                        }

                                        if (_AGEN_mainform.Project_type == "2D")
                                        {
                                            Station_pt2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["2DSTAEND"]);
                                        }
                                        else
                                        {
                                            Station_pt2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["3DSTAEND"]);
                                        }



                                        if (_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_Beg"])) == true &&
                                            _AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_End"])) == true &&
                                            _AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_Beg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_Beg"])) == true &&
                                            _AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_End"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_End"])) == true)
                                        {
                                            double x1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_Beg"]);
                                            double y1 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_Beg"]);

                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["X_End"]);
                                            double y2 = Convert.ToDouble(_AGEN_mainform.dt_mat_lin_extra.Rows[k]["Y_End"]);

                                            Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);

                                            Station_pt1 = poly2d.GetDistAtPoint(point_on_poly2D1);


                                            Point3d point_on_poly2D2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, false);
                                            Station_pt2 = poly2d.GetDistAtPoint(point_on_poly2D2);

                                        }



                                        if (Station_pt1 > Station1 && Station_pt1 <= Station2)
                                        {
                                            Stretch_from_extra1 = Stretch_from_extra1 + 1 * min_spacing;
                                            extra_found_inside1 = true;
                                        }
                                        else if (Station_pt1 == 0)
                                        {
                                            if (Station_pt1 <= Station2)
                                            {
                                                Stretch_from_extra1 = Stretch_from_extra1 + 1 * min_spacing;
                                                extra_found_inside1 = true;
                                            }
                                        }

                                        if (Station_pt2 > Station1 && Station_pt2 <= Station2)
                                        {
                                            Stretch_from_extra2 = Stretch_from_extra2 + 1 * min_spacing;
                                            extra_found_inside2 = true;
                                        }
                                        else if (Station_pt2 == 0)
                                        {
                                            if (Station_pt2 <= Station2)
                                            {
                                                Stretch_from_extra2 = Stretch_from_extra2 + 1 * min_spacing;
                                                extra_found_inside2 = true;
                                            }
                                        }
                                    }

                                    if (extra_found_inside1 == true && strech1 < Stretch_from_extra1)
                                    {
                                        dt_compiled.Rows[i][stretch_val] = Stretch_from_extra1;
                                        double Diff = Stretch_from_extra1 - strech1;
                                        for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                        {
                                            int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                            double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                            double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);

                                            if (Page1 == Page2)
                                            {
                                                dt_compiled.Rows[j][BandL] = band_len2 + Diff;

                                                if (i < j)
                                                {
                                                    dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                                }
                                            }
                                        }
                                        strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                    }

                                    if (extra_found_inside2 == true && strech1 < Stretch_from_extra2)
                                    {
                                        dt_compiled.Rows[i][stretch_val] = Stretch_from_extra2;
                                        double Diff = Stretch_from_extra2 - strech1;
                                        for (int j = 0; j < dt_compiled.Rows.Count; ++j)
                                        {
                                            int Page2 = Convert.ToInt32(dt_compiled.Rows[j][Pageno]);
                                            double deltax2 = Convert.ToDouble(dt_compiled.Rows[j][DeltaX_col]);
                                            double band_len2 = Convert.ToDouble(dt_compiled.Rows[j][BandL]);

                                            if (Page1 == Page2)
                                            {
                                                dt_compiled.Rows[j][BandL] = band_len2 + Diff;

                                                if (i < j)
                                                {
                                                    dt_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                            }




                            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                            {
                                double Station1 = Convert.ToDouble(dt_compiled.Rows[i][Sta1]);
                                double Station2 = Convert.ToDouble(dt_compiled.Rows[i][Sta2]);
                                double M1 = Convert.ToDouble(dt_compiled.Rows[i][Match1]);
                                double M2 = Convert.ToDouble(dt_compiled.Rows[i][Match2]);
                                int Page1 = Convert.ToInt32(dt_compiled.Rows[i][Pageno]);
                                string dwg_name = dt_compiled.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();


                                double ml_len = Convert.ToDouble(dt_compiled.Rows[i][Rect_len]);
                                double band_len = Convert.ToDouble(dt_compiled.Rows[i][BandL]);
                                double Diff = (band_len - ml_len) / 2;
                                double deltax = Convert.ToDouble(dt_compiled.Rows[i][DeltaX_col]);

                                double Station1_label = Convert.ToDouble(dt_compiled.Rows[i][Sta1_label]);
                                double Station2_label = Convert.ToDouble(dt_compiled.Rows[i][Sta2_label]);


                                Point3d InsPt = new Point3d(_AGEN_mainform.Point0_mat.X - lr * ml_len / 2 + lr * deltax, _AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height - (Page1 - 1) * _AGEN_mainform.Band_Separation, 0);

                                string S1 = Functions.Get_chainage_from_double(Station1_label, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                string S2 = Functions.Get_chainage_from_double(Station2_label, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                string Suff = "'";
                                if (_AGEN_mainform.units_of_measurement == "m") Suff = "";

                                string L1 = Convert.ToString(Math.Round(Station2, 0) - Math.Round(Station1, 0)) + Suff;

                                System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                string Block_name = Convert.ToString(dt_compiled.Rows[i][14]);

                                if (BlockTable_data1.Has(Block_name) == false)
                                {
                                    MessageBox.Show("the block name you specified does not belong to the current drawing\r\n" + Block_name);
                                    _AGEN_mainform.tpage_processing.Hide();
                                    set_enable_true();
                                    Ag.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                string suffix1 = "";
                                if (_AGEN_mainform.units_of_measurement == "f") suffix1 = "'";
                                Colectie_nume_atribute.Add("STA1");
                                Colectie_valori.Add(S1);
                                Colectie_nume_atribute.Add("STA2");
                                Colectie_valori.Add(S2);
                                Colectie_nume_atribute.Add("STA11");
                                Colectie_valori.Add(S1);
                                Colectie_nume_atribute.Add("STA21");
                                Colectie_valori.Add(S2);


                                Colectie_nume_atribute.Add("LEN");
                                Colectie_valori.Add(Functions.Get_String_Rounded(Station2 - Station1, _AGEN_mainform.round1) + suffix1);

                                Colectie_nume_atribute.Add("LENGTH");
                                Colectie_valori.Add(Functions.Get_String_Rounded(Station2 - Station1, _AGEN_mainform.round1) + suffix1);



                                for (int k = 15; k < dt_compiled.Columns.Count - 1; ++k)
                                {
                                    Colectie_nume_atribute.Add(dt_compiled.Columns[k].ColumnName);
                                    string VAL = "";
                                    if (dt_compiled.Rows[i][k] != DBNull.Value)
                                    {
                                        VAL = Convert.ToString(dt_compiled.Rows[i][k]);
                                    }
                                    Colectie_valori.Add(VAL);
                                }

                                double strech1 = Convert.ToDouble(dt_compiled.Rows[i][stretch_val]);
                                string visib1 = "";
                                if (dt_compiled.Rows[i][dt_compiled.Columns.Count - 1] != DBNull.Value)
                                {
                                    visib1 = Convert.ToString(dt_compiled.Rows[i][dt_compiled.Columns.Count - 1]);
                                }

                                if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                {
                                    BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", Block_name, InsPt, 1, 0, lname1, Colectie_nume_atribute, Colectie_valori);

                                    #region mat count
                                    int indexcount = -1;
                                    for (int r = 0; r < dtmc.Rows.Count; ++r)
                                    {
                                        string dwg1 = Convert.ToString(dtmc.Rows[r][0]);
                                        if (dwg1 == dwg_name)
                                        {
                                            indexcount = r;
                                            r = dtmc.Rows.Count;
                                        }

                                    }

                                    if (indexcount >= 0)
                                    {
                                        if (Colectie_nume_atribute.Contains("LEN") == true && Colectie_nume_atribute.Contains("MAT") == true)
                                        {
                                            string MAT = Convert.ToString(Colectie_valori[Colectie_nume_atribute.IndexOf("MAT")]);
                                            string LEN = Convert.ToString(Colectie_valori[Colectie_nume_atribute.IndexOf("LEN")]);
                                            string LEN1 = LEN.Replace("'", "").Replace("m", "");

                                            if (MAT != "")
                                            {
                                                if (dtmc.Columns.Contains(MAT) == false)
                                                {
                                                    dtmc.Columns.Add(MAT, typeof(double));
                                                }

                                                double nr = 0;
                                                if (dtmc.Rows[indexcount][MAT] != DBNull.Value)
                                                {
                                                    nr = Convert.ToDouble(dtmc.Rows[indexcount][MAT]);
                                                }


                                                if (Functions.IsNumeric(LEN1) == true)
                                                {
                                                    dtmc.Rows[indexcount][MAT] = nr + Convert.ToDouble(LEN1);
                                                }
                                            }
                                        }
                                    }
                                    #endregion


                                    Functions.Stretch_block(Block1, "Distance1", strech1);
                                    if (visib1 != "") Functions.set_block_visibility(Block1, visib1);



                                    if (_AGEN_mainform.dt_mat_pt != null && _AGEN_mainform.dt_mat_pt.Rows.Count > 0)
                                    {
                                        double x1 = InsPt.X;
                                        double x2 = x1 + lr * strech1;
                                        double y = InsPt.Y;

                                        System.Data.DataTable dt1 = new System.Data.DataTable();
                                        dt1.Columns.Add("blockname", typeof(string));
                                        dt1.Columns.Add("x", typeof(double));
                                        dt1.Columns.Add("y", typeof(double));
                                        dt1.Columns.Add("mat", typeof(string));
                                        dt1.Columns.Add("sta", typeof(string));
                                        dt1.Columns.Add("dnd_zero", typeof(bool));
                                        dt1.Columns.Add("sta2_zero", typeof(bool));
                                        for (int n = 10; n < _AGEN_mainform.dt_mat_pt.Columns.Count; ++n)
                                        {
                                            dt1.Columns.Add(_AGEN_mainform.dt_mat_pt.Columns[n].ColumnName, typeof(string));
                                        }

                                        #region dt_points

                                        for (int k = 0; k < _AGEN_mainform.dt_mat_pt.Rows.Count; ++k)
                                        {
                                            double Station_pt = -1.123;

                                            if (_AGEN_mainform.Project_type == "2D")
                                            {
                                                Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_2DSta]);
                                            }
                                            else
                                            {
                                                Station_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k][_AGEN_mainform.Col_3DSta]);
                                            }

                                            double Station_labeled = Functions.Station_equation_ofV2(Station_pt, _AGEN_mainform.dt_station_equation);


                                            if (_AGEN_mainform.dt_mat_pt.Rows[k]["X"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["X"])) == true &&
                                                   _AGEN_mainform.dt_mat_pt.Rows[k]["Y"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k]["Y"])) == true)
                                            {
                                                double X_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["X"]);
                                                double Y_pt = Convert.ToDouble(_AGEN_mainform.dt_mat_pt.Rows[k]["Y"]);

                                                Point3d point_on_poly2D1 = poly2d.GetClosestPointTo(new Point3d(X_pt, Y_pt, poly2d.Elevation), Vector3d.ZAxis, false);

                                                Station_pt = poly2d.GetDistAtPoint(point_on_poly2D1);
                                                Station_labeled = Functions.Station_equation_ofV2(Station_pt, _AGEN_mainform.dt_station_equation);
                                            }




                                            if (Math.Round(Station_pt, 2) == 0 && Math.Round(Station_pt, 2) == Math.Round(Station1, 2) && Math.Round(Station_pt, 2) <= Math.Round(Station2, 2))
                                            {


                                                string material1 = "";
                                                if (_AGEN_mainform.dt_mat_pt.Rows[k][1] != DBNull.Value)
                                                {
                                                    material1 = Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k][1]);
                                                    if (material1 != null)
                                                    {
                                                        dt1.Rows.Add();
                                                        dt1.Rows[dt1.Rows.Count - 1][0] = _AGEN_mainform.dt_mat_pt.Rows[k][9];

                                                        double xp = x1 + lr * (Station_pt - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                        dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                        dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                        dt1.Rows[dt1.Rows.Count - 1][3] = material1;

                                                        string dispsta = Functions.Get_chainage_from_double(Station_labeled, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                        dt1.Rows[dt1.Rows.Count - 1][4] = dispsta;

                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station1, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                        }


                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station2, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = false;
                                                        }

                                                        int idx_pt = 7;
                                                        for (int n = 10; n < _AGEN_mainform.dt_mat_pt.Columns.Count; ++n)
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][idx_pt] = _AGEN_mainform.dt_mat_pt.Rows[k][n];
                                                            ++idx_pt;
                                                        }
                                                    }
                                                }
                                            }

                                            else if (Math.Round(Station_pt, 2) > Math.Round(Station1, 2) && Math.Round(Station_pt, 2) <= Math.Round(Station2, 2))
                                            {
                                                string material1 = "";
                                                if (_AGEN_mainform.dt_mat_pt.Rows[k][1] != DBNull.Value)
                                                {
                                                    material1 = Convert.ToString(_AGEN_mainform.dt_mat_pt.Rows[k][1]);
                                                    if (material1 != null)
                                                    {
                                                        dt1.Rows.Add();
                                                        dt1.Rows[dt1.Rows.Count - 1][0] = _AGEN_mainform.dt_mat_pt.Rows[k][9];

                                                        double xp = x1 + lr * (Station_pt - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                        dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                        dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                        dt1.Rows[dt1.Rows.Count - 1][3] = material1;

                                                        string dispsta = Functions.Get_chainage_from_double(Station_labeled, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                        dt1.Rows[dt1.Rows.Count - 1][4] = dispsta;

                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station1, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                        }

                                                        if (Math.Round(Station_pt, 2) == Math.Round(Station2, 2))
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = true;
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][6] = false;
                                                        }

                                                        int idx_pt = 7;
                                                        for (int n = 10; n < _AGEN_mainform.dt_mat_pt.Columns.Count; ++n)
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][idx_pt] = _AGEN_mainform.dt_mat_pt.Rows[k][n];
                                                            ++idx_pt;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (dt1.Rows.Count > 0)
                                        {
                                            double xp1 = x1;
                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xi - xp1) < 1 * min_spacing)
                                                {
                                                    xp1 = xp1 + lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp1;


                                                }
                                                else
                                                {
                                                    xp1 = xi;
                                                }

                                                if ((bool)dt1.Rows[m][6] == true)
                                                {
                                                    dt1.Rows[m][1] = x2;
                                                }
                                            }

                                            double xp2 = x2;
                                            for (int m = dt1.Rows.Count - 1; m >= 0; --m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xp2 - xi) < 1 * min_spacing)
                                                {
                                                    xp2 = xp2 - lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp2;
                                                }
                                                else
                                                {
                                                    xp2 = xi;
                                                }
                                                if ((bool)dt1.Rows[m][6] == true)
                                                {
                                                    dt1.Rows[m][1] = x2;
                                                }
                                            }


                                            List<string> lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(Block_name);

                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                string bl = Convert.ToString(dt1.Rows[m][0]);
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double yi = Convert.ToDouble(dt1.Rows[m][2]);
                                                string mati = Convert.ToString(dt1.Rows[m][3]);
                                                string stai = Convert.ToString(dt1.Rows[m][4]);

                                                lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(bl);

                                                System.Collections.Specialized.StringCollection Colectie_nume_atribute_l = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection Colectie_valori_l = new System.Collections.Specialized.StringCollection();
                                                Colectie_nume_atribute_l.Add("STA");
                                                Colectie_valori_l.Add(stai);
                                                Colectie_nume_atribute_l.Add("STA1");
                                                Colectie_valori_l.Add(stai);
                                                Colectie_nume_atribute_l.Add("MAT");
                                                Colectie_valori_l.Add(mati);

                                                for (int n = 7; n < dt1.Columns.Count - 1; ++n)
                                                {
                                                    Colectie_nume_atribute_l.Add(dt1.Columns[n].ColumnName);
                                                    string val = "";
                                                    if (dt1.Rows[m][n] != DBNull.Value)
                                                    {
                                                        val = Convert.ToString(dt1.Rows[m][n]);
                                                    }
                                                    Colectie_valori_l.Add(val);
                                                }


                                                BlockReference Block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                    bl, new Point3d(xi, yi, 0),
                                                                        1, 0, lname1, Colectie_nume_atribute_l, Colectie_valori_l);

                                                #region mat count



                                                if (indexcount >= 0 && Colectie_nume_atribute_l.Contains("MAT") == true)
                                                {
                                                    string MAT = Convert.ToString(Colectie_valori_l[Colectie_nume_atribute_l.IndexOf("MAT")]);


                                                    if (dtmc.Columns.Contains(MAT) == false)
                                                    {
                                                        dtmc.Columns.Add(MAT, typeof(double));
                                                    }

                                                    double nr = 0;
                                                    if (dtmc.Rows[indexcount][MAT] != DBNull.Value)
                                                    {
                                                        nr = Convert.ToDouble(dtmc.Rows[indexcount][MAT]);
                                                    }
                                                    dtmc.Rows[indexcount][MAT] = ++nr;


                                                }
                                                #endregion


                                                string visib2 = "";
                                                if (dt1.Rows[m][dt1.Columns.Count - 1] != DBNull.Value)
                                                {
                                                    visib2 = Convert.ToString(dt1.Rows[m][dt1.Columns.Count - 1]);
                                                }

                                                if (visib2 != "") Functions.set_block_visibility(Block2, visib2);

                                            }

                                        }
                                        #endregion

                                    }


                                    if (Data_table_compiled_extra != null && Data_table_compiled_extra.Rows.Count > 0)
                                    {


                                        double x1 = InsPt.X;
                                        double x2 = x1 + lr * strech1;
                                        double y = InsPt.Y;

                                        System.Data.DataTable dt1 = new System.Data.DataTable();
                                        dt1.Columns.Add("blockname", typeof(string));
                                        dt1.Columns.Add("x1", typeof(double));
                                        dt1.Columns.Add("y1", typeof(double));
                                        dt1.Columns.Add("STA1", typeof(string));
                                        dt1.Columns.Add("STA2", typeof(string));
                                        dt1.Columns.Add("dnd_zero", typeof(bool));

                                        for (int n = 15; n < Data_table_compiled_extra.Columns.Count; ++n)
                                        {
                                            string nume_col = Data_table_compiled_extra.Columns[n].ColumnName;

                                            dt1.Columns.Add(nume_col, typeof(string));
                                        }

                                        #region dt_mat_lin_extra

                                        for (int k = 0; k < Data_table_compiled_extra.Rows.Count; ++k)
                                        {

                                            debug = "k = " + k.ToString() + "\r\ni = " + i.ToString() + "\r\nblock is not present?";


                                            double Station_pt1 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta1]);
                                            double Station_pt2 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta2]);

                                            double Station_labeled1 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta1_label]);
                                            double Station_labeled2 = Convert.ToDouble(Data_table_compiled_extra.Rows[k][Sta2_label]);

                                            if (Math.Round(Station_pt1, 2) == 0 && Math.Round(Station_pt1, 2) == Math.Round(Station1, 2) && Math.Round(Station_pt1, 2) <= Math.Round(Station2, 2))
                                            {
                                                if (Data_table_compiled_extra.Rows[k][1] != DBNull.Value)
                                                {
                                                    string nume_block = Convert.ToString(Data_table_compiled_extra.Rows[k][14]);
                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][0] = Data_table_compiled_extra.Rows[k][14];
                                                    double xp = x1 + lr * (Station_pt1 - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                    string dispsta1 = Functions.Get_chainage_from_double(Station_labeled1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][3] = dispsta1;
                                                    string dispsta2 = Functions.Get_chainage_from_double(Station_labeled2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][4] = dispsta2;
                                                    if (Math.Round(Station_pt1, 2) == Math.Round(Station1, 2))
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                    }
                                                    else
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                    }

                                                    int idx_pt = 6;
                                                    for (int n = 15; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][idx_pt] = Data_table_compiled_extra.Rows[k][n];
                                                        ++idx_pt;
                                                    }
                                                }
                                            }
                                            else if (Math.Round(Station_pt1, 2) >= Math.Round(Station1, 2) && Math.Round(Station_pt1, 2) <= Math.Round(Station2, 2) && Math.Round(Station_pt2, 2) <= Math.Round(M2, 2))
                                            {
                                                if (Data_table_compiled_extra.Rows[k][1] != DBNull.Value)
                                                {
                                                    string nume_block = Convert.ToString(Data_table_compiled_extra.Rows[k][14]);
                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][0] = Data_table_compiled_extra.Rows[k][14];
                                                    double xp = x1 + lr * (Station_pt1 - Station1) * Math.Abs(x2 - x1) / (Station2 - Station1);
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = xp;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                    string dispsta1 = Functions.Get_chainage_from_double(Station_labeled1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][3] = dispsta1;
                                                    string dispsta2 = Functions.Get_chainage_from_double(Station_labeled2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][4] = dispsta2;
                                                    if (Math.Round(Station_pt1, 2) == Math.Round(Station1, 2))
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = true;
                                                    }
                                                    else
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][5] = false;
                                                    }
                                                    int idx_pt = 6;
                                                    for (int n = 15; n < Data_table_compiled_extra.Columns.Count; ++n)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][idx_pt] = Data_table_compiled_extra.Rows[k][n];
                                                        ++idx_pt;
                                                    }
                                                }
                                            }
                                        }

                                        if (dt1.Rows.Count > 0)
                                        {
                                            double xp1 = x1;
                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xi - xp1) < 1 * min_spacing)
                                                {
                                                    xp1 = xp1 + lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp1;
                                                }
                                                else
                                                {
                                                    xp1 = xi;
                                                }
                                            }

                                            double xp2 = x2;

                                            for (int m = dt1.Rows.Count - 1; m >= 0; --m)
                                            {
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);

                                                double temp_dndlr = min_spacing;
                                                if ((bool)dt1.Rows[m][5] == true)
                                                {
                                                    temp_dndlr = 0;
                                                }
                                                if (Math.Abs(xp2 - xi) < 1 * min_spacing)
                                                {
                                                    xp2 = xp2 - lr * temp_dndlr;
                                                    dt1.Rows[m][1] = xp2;
                                                }
                                                else
                                                {
                                                    xp2 = xi;
                                                }
                                            }



                                            List<string> lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(Block_name);


                                            for (int m = 0; m < dt1.Rows.Count; ++m)
                                            {
                                                string bl = Convert.ToString(dt1.Rows[m][0]);
                                                double xi = Convert.ToDouble(dt1.Rows[m][1]);
                                                double yi = Convert.ToDouble(dt1.Rows[m][2]);
                                                string ss1 = Convert.ToString(dt1.Rows[m][3]);
                                                string ss2 = Convert.ToString(dt1.Rows[m][4]);

                                                lista_atribute_din_block = Functions.Incarca_existing_Atributes_to_list(bl);

                                                System.Collections.Specialized.StringCollection Colectie_nume_atribute_l = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection Colectie_valori_l = new System.Collections.Specialized.StringCollection();
                                                Colectie_nume_atribute_l.Add("STA1");
                                                Colectie_valori_l.Add(ss1);
                                                Colectie_nume_atribute_l.Add("STA11");
                                                Colectie_valori_l.Add(ss1);
                                                Colectie_nume_atribute_l.Add("STA2");
                                                Colectie_valori_l.Add(ss2);
                                                Colectie_nume_atribute_l.Add("STA21");
                                                Colectie_valori_l.Add(ss2);

                                                for (int n = 6; n < dt1.Columns.Count - 1; ++n)
                                                {
                                                    Colectie_nume_atribute_l.Add(dt1.Columns[n].ColumnName);
                                                    string val = "";
                                                    if (dt1.Rows[m][n] != DBNull.Value)
                                                    {
                                                        val = Convert.ToString(dt1.Rows[m][n]);
                                                    }
                                                    Colectie_valori_l.Add(val);
                                                }
                                                BlockReference Block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                    bl, new Point3d(xi, yi, 0), 1, 0, lname1, Colectie_nume_atribute_l, Colectie_valori_l);


                                                #region mat count
                                                if (indexcount >= 0)
                                                {
                                                    if (Colectie_nume_atribute_l.Contains("LEN") == true && Colectie_nume_atribute_l.Contains("MAT") == true)
                                                    {
                                                        string MAT = Convert.ToString(Colectie_valori_l[Colectie_nume_atribute_l.IndexOf("MAT")]);
                                                        string LEN = Convert.ToString(Colectie_valori_l[Colectie_nume_atribute_l.IndexOf("LEN")]);
                                                        string LEN1 = LEN.Replace("'", "").Replace("m", "");
                                                        if (dtmc.Columns.Contains(MAT) == false)
                                                        {
                                                            dtmc.Columns.Add(MAT, typeof(double));
                                                        }
                                                        double nr = 0;
                                                        if (dtmc.Rows[indexcount][MAT] != DBNull.Value)
                                                        {
                                                            nr = Convert.ToDouble(dtmc.Rows[indexcount][MAT]);
                                                        }
                                                        if (Functions.IsNumeric(LEN1) == true)
                                                        {
                                                            dtmc.Rows[indexcount][MAT] = nr + Convert.ToDouble(LEN1);
                                                        }
                                                    }
                                                }
                                                #endregion


                                                string visib2 = "";
                                                if (dt1.Rows[m][dt1.Columns.Count - 1] != DBNull.Value)
                                                {
                                                    visib2 = Convert.ToString(dt1.Rows[m][dt1.Columns.Count - 1]);
                                                }

                                                if (visib2 != "") Functions.set_block_visibility(Block2, visib2);
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }


                        Trans1.Commit();

                        dataGridView_materials.DataSource = dtmc;
                        dataGridView_materials.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                        dataGridView_materials.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_materials.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        dataGridView_materials.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_materials.DefaultCellStyle.ForeColor = Color.White;
                        dataGridView_materials.EnableHeadersVisualStyles = false;
                        _AGEN_mainform.dt_mat_lin = null;
                        _AGEN_mainform.dt_mat_lin_extra = null;
                        _AGEN_mainform.dt_mat_pt = null;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + debug);
            }

            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();

            Ag.WindowState = FormWindowState.Normal;
        }

        private void button_scan_heavy_wall_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();

            this.MdiParent.WindowState = FormWindowState.Minimized;
            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline file does not exist");

                    return;
                }

                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline file does not have any data");
                    return;
                }


            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }




            System.Data.DataTable Dt_poly = Functions.Creaza_prof_poly_dt_structure();




            try
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                ObjectId[] Empty_array = null;
                Editor1.SetImpliedSelection(Empty_array);




                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly;
                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                Prompt_poly.MessageForAdding = "\nselect heavy wall lines";
                Prompt_poly.SingleOnly = false;
                Rezultat_poly = ThisDrawing.Editor.GetSelection(Prompt_poly);

                if (Rezultat_poly.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        LayerTable layertable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("STA1", typeof(double));
                        dt1.Columns.Add("STA2", typeof(double));
                        dt1.Columns.Add("X1", typeof(double));
                        dt1.Columns.Add("Y1", typeof(double));
                        dt1.Columns.Add("Z1", typeof(double));
                        dt1.Columns.Add("X2", typeof(double));
                        dt1.Columns.Add("Y2", typeof(double));
                        dt1.Columns.Add("Z2", typeof(double));
                        dt1.Columns.Add("LAYER", typeof(string));
                        dt1.Columns.Add("LAYER DESCRIPTION", typeof(string));

                        dt1.Columns.Add("2D Distance", typeof(double));



                        Polyline poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        Polyline3d poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        for (int i = 0; i < Rezultat_poly.Value.Count; ++i)
                        {
                            Entity Ent1 = Trans1.GetObject(Rezultat_poly.Value[i].ObjectId, OpenMode.ForRead) as Entity;
                            if (Ent1 != null && (Ent1 is Line || Ent1 is Polyline || Ent1 is Polyline3d || Ent1 is Polyline2d || Ent1 is DBPoint))
                            {
                                Curve curve1 = Ent1 as Curve;

                                DBPoint pt_obj = Ent1 as DBPoint;
                                if (curve1 != null)
                                {
                                    Point3d p1 = curve1.StartPoint;
                                    Point3d p2 = curve1.EndPoint;

                                    Point3d pt1 = poly2D.GetClosestPointTo(p1, Vector3d.ZAxis, false);
                                    Point3d pt2 = poly2D.GetClosestPointTo(p2, Vector3d.ZAxis, false);

                                    double d1 = -1;
                                    double d2 = -1;

                                    double param1 = -1;
                                    double param2 = -1;

                                    if (_AGEN_mainform.Project_type == "2D")
                                    {
                                        d1 = poly2D.GetDistAtPoint(pt1);
                                        d2 = poly2D.GetDistAtPoint(pt2);
                                    }
                                    else
                                    {
                                        param1 = poly2D.GetParameterAtPoint(pt1);
                                        param2 = poly2D.GetParameterAtPoint(pt2);
                                        pt1 = poly3D.GetPointAtParameter(param1);
                                        pt2 = poly3D.GetPointAtParameter(param2);
                                        d1 = poly3D.GetDistanceAtParameter(param1);
                                        d2 = poly3D.GetDistanceAtParameter(param2);
                                    }

                                    if (d1 > d2)
                                    {
                                        Point3d t = pt1;
                                        pt1 = pt2;
                                        pt2 = t;
                                        double tt = param1;
                                        param1 = param2;
                                        param2 = tt;
                                        tt = d1;
                                        d1 = d2;
                                        d2 = tt;
                                    }

                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1]["X1"] = pt1.X;
                                    dt1.Rows[dt1.Rows.Count - 1]["Y1"] = pt1.Y;
                                    dt1.Rows[dt1.Rows.Count - 1]["Z1"] = pt1.Z;
                                    dt1.Rows[dt1.Rows.Count - 1]["X2"] = pt2.X;
                                    dt1.Rows[dt1.Rows.Count - 1]["Y2"] = pt2.Y;
                                    dt1.Rows[dt1.Rows.Count - 1]["Z2"] = pt2.Z;
                                    dt1.Rows[dt1.Rows.Count - 1]["LAYER"] = Ent1.Layer;

                                    LayerTableRecord ltr = Trans1.GetObject(layertable1[Ent1.Layer], OpenMode.ForRead) as LayerTableRecord;
                                    if (ltr != null)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1]["LAYER DESCRIPTION"] = ltr.Description;
                                    }


                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                    {
                                        double d1_2d = poly2D.GetDistanceAtParameter(param1);
                                        double d2_2d = poly2D.GetDistanceAtParameter(param2);
                                        double b11 = -1.23456;
                                        double b21 = -1.23456;
                                        double Sta1 = Functions.get_stationCSF_from_point(poly2D, pt1, d1_2d, _AGEN_mainform.dt_centerline, ref b11);
                                        double Sta2 = Functions.get_stationCSF_from_point(poly2D, pt2, d2_2d, _AGEN_mainform.dt_centerline, ref b21);
                                        dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(Sta1, _AGEN_mainform.round1);
                                        dt1.Rows[dt1.Rows.Count - 1]["STA2"] = Math.Round(Sta2, _AGEN_mainform.round1);
                                    }
                                    else
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(d1, _AGEN_mainform.round1);
                                        dt1.Rows[dt1.Rows.Count - 1]["STA2"] = Math.Round(d2, _AGEN_mainform.round1);
                                    }
                                }

                                if (pt_obj != null)
                                {
                                    Point3d p1 = pt_obj.Position;
                                    Point3d pt1 = poly2D.GetClosestPointTo(p1, Vector3d.ZAxis, false);

                                    double d1 = -1;

                                    double param1 = -1;


                                    if (_AGEN_mainform.Project_type == "2D")
                                    {
                                        d1 = poly2D.GetDistAtPoint(pt1);

                                    }
                                    else
                                    {
                                        param1 = poly2D.GetParameterAtPoint(pt1);

                                        pt1 = poly3D.GetPointAtParameter(param1);

                                        d1 = poly3D.GetDistanceAtParameter(param1);

                                    }

                                    double dist_from_cl = Math.Pow(Math.Pow(p1.X - pt1.X, 2) + Math.Pow(p1.Y - pt1.Y, 2), 0.5);

                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1]["X1"] = pt1.X;
                                    dt1.Rows[dt1.Rows.Count - 1]["Y1"] = pt1.Y;
                                    dt1.Rows[dt1.Rows.Count - 1]["Z1"] = pt1.Z;
                                    dt1.Rows[dt1.Rows.Count - 1]["2D Distance"] = dist_from_cl;

                                    dt1.Rows[dt1.Rows.Count - 1]["LAYER"] = Ent1.Layer;
                                    LayerTableRecord ltr = Trans1.GetObject(layertable1[Ent1.Layer], OpenMode.ForRead) as LayerTableRecord;
                                    if (ltr != null)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1]["LAYER DESCRIPTION"] = ltr.Description;
                                    }


                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                    {
                                        double d1_2d = poly2D.GetDistanceAtParameter(param1);
                                        double b11 = -1.23456;
                                        double Sta1 = Functions.get_stationCSF_from_point(poly2D, pt1, d1_2d, _AGEN_mainform.dt_centerline, ref b11);
                                        dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(Sta1, _AGEN_mainform.round1);
                                    }
                                    else
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(d1, _AGEN_mainform.round1);
                                    }
                                }

                                Functions.add_object_data_to_datatable(dt1, Tables1, Ent1.ObjectId);
                            }
                        }

                        poly3D.Erase();
                        Trans1.Commit();
                        dt1 = Functions.Sort_data_table(dt1, "STA1");
                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            set_enable_true();

            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        private void button_place_extra_pts_Click(object sender, EventArgs e)
        {
            if (dt_pt_extra == null || dt_pt_extra.Rows.Count == 0) return;


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        string ln = "_PTS_extra";

                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as LayerTable;





                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("x0", typeof(double));
                        dt1.Columns.Add("y0", typeof(double));
                        dt1.Columns.Add("dist1", typeof(double));
                        dt1.Columns.Add("sta1", typeof(double));
                        dt1.Columns.Add("sta2", typeof(double));





                        foreach (ObjectId id1 in BTrecord)
                        {
                            BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                            if (block1 != null)
                            {
                                LayerTableRecord layer1 = Trans1.GetObject(LayerTable1[block1.Layer], OpenMode.ForRead) as LayerTableRecord;
                                if (layer1.IsFrozen == false && layer1.IsOff == false && layer1.IsDependent == false)
                                {
                                    if (block1.AttributeCollection.Count > 0 && block1.IsDynamicBlock == true)
                                    {

                                        double dist1 = Functions.Get_Param_Value_block(block1, "Distance1");


                                        if (dist1 > 0)
                                        {

                                            double sta1 = -1;
                                            double sta2 = -1;


                                            foreach (ObjectId id2 in block1.AttributeCollection)
                                            {
                                                AttributeReference atr1 = Trans1.GetObject(id2, OpenMode.ForRead) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    string atr_name = atr1.Tag;
                                                    string atr_val = atr1.TextString;


                                                    if (atr_name.ToUpper() == "STA1")
                                                    {
                                                        if (Functions.IsNumeric(atr_val.Replace(" ", "").Replace("+", "")) == true)
                                                        {
                                                            sta1 = Convert.ToDouble(atr_val.Replace(" ", "").Replace("+", ""));


                                                        }
                                                    }

                                                    if (atr_name.ToUpper() == "STA2")
                                                    {
                                                        if (Functions.IsNumeric(atr_val.Replace(" ", "").Replace("+", "")) == true)
                                                        {
                                                            sta2 = Convert.ToDouble(atr_val.Replace(" ", "").Replace("+", ""));

                                                        }
                                                    }


                                                }
                                            }


                                            if (sta1 != sta2 && sta1 != -1 && sta2 != -1)
                                            {

                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["x0"] = block1.Position.X;
                                                dt1.Rows[dt1.Rows.Count - 1]["y0"] = block1.Position.Y;
                                                dt1.Rows[dt1.Rows.Count - 1]["dist1"] = dist1;
                                                dt1.Rows[dt1.Rows.Count - 1]["sta1"] = sta1;
                                                dt1.Rows[dt1.Rows.Count - 1]["sta2"] = sta2;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        List<int> lista_processed = new List<int>();

                        if (dt1.Rows.Count > 0)
                        {
                            Functions.Creaza_layer(ln, 2, true);
                            System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();

                            for (int i = 10; i < dt_pt_extra.Columns.Count - 1; i++)
                            {
                                col_atr.Add(dt_pt_extra.Columns[i].ColumnName);
                            }
                            col_atr.Add("STA");


                            for (int i = 0; i < dt_pt_extra.Rows.Count; ++i)
                            {
                                double sta = -1.123;
                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();


                                if (dt_pt_extra.Rows[i][9] != DBNull.Value)
                                {
                                    string bn = Convert.ToString(dt_pt_extra.Rows[i][9]);
                                    if (_AGEN_mainform.Project_type == "2D")
                                    {
                                        if (dt_pt_extra.Rows[i][_AGEN_mainform.Col_2DSta] != DBNull.Value)
                                        {
                                            sta = Convert.ToDouble(dt_pt_extra.Rows[i][_AGEN_mainform.Col_2DSta]);
                                        }
                                    }
                                    else
                                    {
                                        if (dt_pt_extra.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                                        {
                                            sta = Convert.ToDouble(dt_pt_extra.Rows[i][_AGEN_mainform.Col_3DSta]);
                                        }
                                    }

                                    if (dt_pt_extra.Rows[i][_AGEN_mainform.Col_eqsta] != DBNull.Value)
                                    {
                                        sta = Convert.ToDouble(dt_pt_extra.Rows[i][_AGEN_mainform.Col_eqsta]);
                                    }

                                    if (sta != -1.123)
                                    {
                                        for (int j = 10; j < dt_pt_extra.Columns.Count - 1; j++)
                                        {

                                            if (dt_pt_extra.Rows[i][j] != DBNull.Value)
                                            {
                                                col_val.Add(Convert.ToString(dt_pt_extra.Rows[i][j]));
                                            }
                                            else
                                            {
                                                col_val.Add("");
                                            }
                                        }
                                        for (int k = 0; k < dt1.Rows.Count; k++)
                                        {
                                            double x0 = Convert.ToDouble(dt1.Rows[k]["x0"]);
                                            double y0 = Convert.ToDouble(dt1.Rows[k]["y0"]);
                                            double dist1 = Convert.ToDouble(dt1.Rows[k]["dist1"]);
                                            double sta1 = Convert.ToDouble(dt1.Rows[k]["sta1"]);
                                            double sta2 = Convert.ToDouble(dt1.Rows[k]["sta2"]);

                                            double x = -1.23456;

                                            if (sta <= sta2 && sta >= sta1)
                                            {
                                                double deltax = (sta - sta1) * dist1 / (sta2 - sta1);
                                                if (_AGEN_mainform.Left_to_Right == true)
                                                {
                                                    x = x0 + deltax;
                                                }
                                                else
                                                {
                                                    x = x0 - deltax;
                                                }

                                            }

                                            if (x != -1.23456)
                                            {
                                                col_val.Add(Functions.Get_chainage_from_double(sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1));
                                                BlockReference Block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                    bn, new Point3d(x, y0, 0), 1, 0, ln, col_atr, col_val);
                                                if (lista_processed.Contains(i) == false) lista_processed.Add(i);
                                            }

                                        }

                                    }


                                }


                            }

                        }

                        Trans1.Commit();

                        if (lista_processed.Count>0 && lista_processed.Count < dt_pt_extra.Rows.Count)
                        {
                            System.Data.DataTable dt2 = dt_pt_extra.Clone();
                            for (int i = 0; i < dt_pt_extra.Rows.Count; ++i)
                            {
                                if (lista_processed.Contains(i) == false)
                                {
                                    System.Data.DataRow row2 = dt2.NewRow();
                                    row2.ItemArray = dt_pt_extra.Rows[i].ItemArray;
                                    dt2.Rows.Add(row2);
                                }
                            }

                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt2);
                        }

                    }
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
    }
}



