using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace Alignment_mdi
{
    public partial class AGEN_CrossingScan : Form
    {
        bool first_message_box = false;

        public AGEN_CrossingScan()
        {
            InitializeComponent();
        }



        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_crossing_table_from_excel);
            lista_butoane.Add(button_open_crossing_xlxs);
            lista_butoane.Add(button_save_crossing_table_to_excel);
            lista_butoane.Add(button_scan_for_crossings);
            lista_butoane.Add(button_scan_with_offsets);
            lista_butoane.Add(button_show_crossing_band_settings);
            lista_butoane.Add(button_show_scan_and_draw_crossings);
            lista_butoane.Add(button_transfer_to_excel_crossing_band);
            lista_butoane.Add(button_calc_station_from_point);
            lista_butoane.Add(button_create_empty_crossing);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_crossing_table_from_excel);
            lista_butoane.Add(button_open_crossing_xlxs);
            lista_butoane.Add(button_save_crossing_table_to_excel);
            lista_butoane.Add(button_scan_for_crossings);
            lista_butoane.Add(button_scan_with_offsets);
            lista_butoane.Add(button_show_crossing_band_settings);
            lista_butoane.Add(button_show_scan_and_draw_crossings);
            lista_butoane.Add(button_transfer_to_excel_crossing_band);
            lista_butoane.Add(button_calc_station_from_point);
            lista_butoane.Add(button_create_empty_crossing);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private void button_show_crossing_band_settings_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();

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


            _AGEN_mainform.tpage_layer_alias.Show();

        }

        private void textBox_offset_value_KeyPress(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }

        private void create_crossing_file(string File1)
        {
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

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                if (System.IO.File.Exists(File1) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                    try
                    {
                        string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                        if (segment1 == "not defined") segment1 = "";
                        Functions.Create_header_crossing_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
                        Workbook1.SaveAs(File1);
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
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
        }


        private void button_scan_for_intersections_Click(object sender, EventArgs e)
        {
            first_message_box = false;
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            



            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            if (_AGEN_mainform.dt_layer_alias == null)
            {
                MessageBox.Show("No layer alias table loaded");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_viewport_settings.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();

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


                _AGEN_mainform.tpage_layer_alias.Show();

                return;
            }


            if (_AGEN_mainform.dt_layer_alias.Rows.Count == 0)
            {
                MessageBox.Show("The layer alias table contains no data");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_viewport_settings.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();

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


                _AGEN_mainform.tpage_layer_alias.Show();
                return;
            }



            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == true)
            {
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                set_enable_true();
                MessageBox.Show("the centerline data file does not exist");
                return;
            }

            string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

            if (System.IO.File.Exists(fisier_cs) == false)
            {
                create_crossing_file(fisier_cs);
            }

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0 || _AGEN_mainform.tpage_setup.get_no_segments() > 1)
            {
                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
            }





            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("no centerline data!");
                return;
            }




            List<string> Lista_layere_de_scanat = new List<string>();
            List<string> Lista_display_in_crossing_band = new List<string>();

            for (int i = 0; i < _AGEN_mainform.dt_layer_alias.Rows.Count; ++i)
            {
                if (_AGEN_mainform.dt_layer_alias.Rows[i][0] != DBNull.Value)
                {
                    Lista_layere_de_scanat.Add(_AGEN_mainform.dt_layer_alias.Rows[i][0].ToString().ToLower());
                }
                else
                {
                    Lista_layere_de_scanat.Add("danpopescuistheking");
                }
                if (_AGEN_mainform.dt_layer_alias.Rows[i][20] != DBNull.Value)
                {
                    string Display = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i][20]);
                    if (Display.ToUpper() == "YES") Lista_display_in_crossing_band.Add("YES");
                    if (Display.ToUpper() == "NO") Lista_display_in_crossing_band.Add("NO");

                }
                else
                {
                    Lista_display_in_crossing_band.Add("NO");
                }

            }

            System.Data.DataTable dt_alias = _AGEN_mainform.dt_layer_alias;

            double poly_length = 0;

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
                        BlockTableRecord BTrecord = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                        Polyline3d poly3d = null;
                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        poly_length = poly2d.Length;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            poly_length = poly3d.Length;

                        }

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
                                            eq_meas = poly3d.GetDistanceAtParameter(param1);
                                        }

                                        _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (_AGEN_mainform.COUNTRY == "USA") _AGEN_mainform.dt_station_equation = null;
                        }
                        #endregion


                        _AGEN_mainform.Data_Table_crossings = Functions.Creaza_crossing_datatable_structure();

                        foreach (ObjectId odid in BTrecord)
                        {
                            Curve Curba0 = Trans1.GetObject(odid, OpenMode.ForRead) as Curve;
                            if (Curba0 != null && Curba0.IsErased == false)
                            {
                                if (_AGEN_mainform.Project_type == "2D" || (_AGEN_mainform.Project_type == "3D" && Curba0.ObjectId != poly3d.ObjectId))
                                {
                                    if (Lista_layere_de_scanat.Contains(Curba0.Layer.ToLower()) == true)
                                    {
                                        int idx = Lista_layere_de_scanat.IndexOf(Curba0.Layer.ToLower());


                                        Line l1 = Curba0.Clone() as Line;
                                        Polyline p1 = Curba0.Clone() as Polyline;
                                        Arc a1 = Curba0.Clone() as Arc;
                                        Polyline3d p3 = Curba0.Clone() as Polyline3d;

                                        Point3dCollection Col_int = new Point3dCollection();

                                        double Z_int = 0;

                                        if (l1 != null)
                                        {
                                            l1.StartPoint = new Point3d(l1.StartPoint.X, l1.StartPoint.Y, poly2d.Elevation);
                                            l1.EndPoint = new Point3d(l1.EndPoint.X, l1.EndPoint.Y, poly2d.Elevation);
                                            Col_int = Functions.Intersect_on_both_operands(l1, poly2d);
                                            if (Col_int.Count > 0)
                                            {
                                                Point3d pt_on_line = Curba0.GetClosestPointTo(Col_int[0], Vector3d.ZAxis, false);
                                                Z_int = pt_on_line.Z;
                                            }

                                        }

                                        if (p1 != null)
                                        {
                                            Z_int = p1.Elevation;
                                            p1.Elevation = poly2d.Elevation;
                                            Col_int = Functions.Intersect_on_both_operands(p1, poly2d);

                                        }

                                        if (a1 != null)
                                        {

                                            a1.Center = new Point3d(a1.Center.X, a1.Center.Y, poly2d.Elevation);

                                            Col_int = Functions.Intersect_on_both_operands(a1, poly2d);
                                            if (Col_int.Count > 0)
                                            {
                                                Point3d pt_on_line = Curba0.GetClosestPointTo(Col_int[0], Vector3d.ZAxis, false);
                                                Z_int = pt_on_line.Z;
                                            }
                                        }

                                        if (p3 != null)
                                        {
                                            p1 = Functions.Build_2dpoly_from_3d(p3);
                                            p1.Elevation = poly2d.Elevation;
                                            Col_int = Functions.Intersect_on_both_operands(p1, poly2d);
                                            if (Col_int.Count > 0)
                                            {
                                                Point3d pt_on_line = p1.GetClosestPointTo(Col_int[0], Vector3d.ZAxis, false);
                                                double param1 = p1.GetParameterAtPoint(pt_on_line);

                                                Z_int = p3.GetPointAtParameter(param1).Z;
                                            }


                                        }

                                        if (Col_int.Count > 0)
                                        {
                                            for (int j = 0; j < Col_int.Count; ++j)
                                            {
                                                try
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows.Add();

                                                    double sta_meas = 0;
                                                    if (_AGEN_mainform.Project_type == "2D")
                                                    {
                                                        Point3d pt1 = poly2d.GetClosestPointTo(Col_int[j], Vector3d.ZAxis, false);
                                                        sta_meas = poly2d.GetDistAtPoint(pt1);

                                                        _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][1] = sta_meas;

                                                    }

                                                    else
                                                    {
                                                        Point3d pt1 = poly2d.GetClosestPointTo(Col_int[j], Vector3d.ZAxis, false);
                                                        double param2d = poly2d.GetParameterAtPoint(pt1);
                                                        sta_meas = poly3d.GetDistanceAtParameter(param2d);

                                                        _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][2] = sta_meas;

                                                    }

                                                    if (_AGEN_mainform.dt_station_equation != null)
                                                    {
                                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][3] =
                                                               Math.Round(Functions.Station_equation_ofV2(sta_meas, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                        }
                                                        else
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][3] = DBNull.Value;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][3] = DBNull.Value;
                                                    }

                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][7] = Col_int[j].X;
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][8] = Col_int[j].Y;
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][9] = Z_int;

                                                    string alias1 = get_description_for_obj(dt_alias, idx, Curba0);

                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][6] = alias1;
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][4] = Curba0.GetType().ToString().Replace("Autodesk.AutoCAD.DatabaseServices.", "");
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][5] = Curba0.Layer;
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispXing"] = Lista_display_in_crossing_band[idx];
                                                    List<string> lista_bl_atr = get_block_and_atr_from_layer_alias(Curba0.Layer);

                                                    if (lista_bl_atr.Count == 3)
                                                    {
                                                        if (lista_bl_atr[0] != "")
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["Prof Block Name"] = lista_bl_atr[0];
                                                            if (lista_bl_atr[1] != "") _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["Attrib Sta Prof"] = lista_bl_atr[1];
                                                            if (lista_bl_atr[2] != "") _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["Attrib Desc Prof"] = lista_bl_atr[2];
                                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispProf"] = "YES";
                                                        }
                                                        else
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispProf"] = "NO";
                                                        }

                                                    }
                                                    else
                                                    {
                                                        _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispProf"] = "NO";
                                                    }


                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show(ex.Message);
                                                }
                                            }


                                        }


                                    }

                                }
                            }
                        }




                        if (radioButton_append_to_crossing_table.Checked == true)
                        {
                            System.Data.DataTable table1 = Load_existing_crossing(fisier_cs);
                            if (table1.Rows.Count > 0)
                            {
                                for (int i = 0; i < table1.Rows.Count; ++i)
                                {
                                    if (table1.Rows[i][1] != DBNull.Value || table1.Rows[i][2] != DBNull.Value)
                                    {
                                        _AGEN_mainform.Data_Table_crossings.Rows.Add();
                                        for (int j = 0; j < 19; ++j)
                                        {
                                            if (table1.Rows[i][j] != DBNull.Value)
                                            {
                                                _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][j] = table1.Rows[i][j];
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (_AGEN_mainform.Data_Table_crossings != null)
                        {
                            if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                            {

                                string Display_in_crossing_band = "NO";


                                for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                {


                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"] == DBNull.Value) _AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"] = Display_in_crossing_band;

                                    double St1 = 0;

                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value || _AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value)
                                    {
                                        double div1 = 10;
                                        if (_AGEN_mainform.round1 == 1) div1 = 100;
                                        if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                        if (_AGEN_mainform.round1 == 3) div1 = 10000;

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value)
                                        {
                                            St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][1]), _AGEN_mainform.round1);
                                            if (St1 >= poly2d.Length) St1 = Math.Floor(Math.Round(poly2d.Length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][1] = St1;
                                        }

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value)
                                        {
                                            St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][2]), _AGEN_mainform.round1);
                                            if (St1 >= poly_length) St1 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][2] = St1;
                                        }

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value)
                                        {
                                            St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][1]), _AGEN_mainform.round1);
                                        }
                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value)
                                        {
                                            St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][2]), _AGEN_mainform.round1);
                                        }


                                        if (_AGEN_mainform.dt_station_equation != null)
                                        {
                                            if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                            {
                                                _AGEN_mainform.Data_Table_crossings.Rows[i][3] =
                                                   Math.Round(Functions.Station_equation_ofV2(St1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                            }
                                            else
                                            {
                                                _AGEN_mainform.Data_Table_crossings.Rows[i][3] = DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][3] = DBNull.Value;
                                        }
                                    }
                                }




                                if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                                Trans1.Commit();
                            }
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


            dataGridView_xing.DataSource = _AGEN_mainform.Data_Table_crossings;
            dataGridView_xing.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_xing.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_xing.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_xing.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_xing.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_xing.EnableHeadersVisualStyles = false;


            radioButton_append_to_crossing_table.Checked = true;
            MessageBox.Show("done");
        }

        private string get_description_for_obj(System.Data.DataTable dt_alias, int index_layer, Entity Ent1)
        {
            string alias1 = "";
            bool use_OD = false;
            string generic_descr = "";
            if (dt_alias.Rows[index_layer][1] != DBNull.Value)
            {
                generic_descr = dt_alias.Rows[index_layer][1].ToString();
                if (dt_alias.Rows[index_layer][2] != DBNull.Value)
                {

                    string yn = Convert.ToString(dt_alias.Rows[index_layer][2]);
                    if (yn.ToLower() == "yes")
                    {
                        use_OD = true;
                    }

                    alias1 = generic_descr;

                }
            }

            if (use_OD == true)
            {


                string Pref1 = "";
                string Suff1 = "";
                string Pref2 = "";
                string Suff2 = "";
                string Pref3 = "";
                string Suff3 = "";
                string Pref4 = "";
                string Suff4 = "";
                string nf1 = "";
                string nf2 = "";
                string nf3 = "";
                string nf4 = "";

                if (dt_alias.Rows[index_layer][5] != DBNull.Value)
                {
                    Pref1 = dt_alias.Rows[index_layer][5].ToString();
                }

                if (dt_alias.Rows[index_layer][8] != DBNull.Value)
                {
                    Pref2 = dt_alias.Rows[index_layer][8].ToString();
                }

                if (dt_alias.Rows[index_layer][11] != DBNull.Value)
                {
                    Pref3 = dt_alias.Rows[index_layer][11].ToString();
                }

                if (dt_alias.Rows[index_layer][14] != DBNull.Value)
                {
                    Pref4 = dt_alias.Rows[index_layer][14].ToString();
                }


                if (dt_alias.Rows[index_layer][7] != DBNull.Value)
                {
                    Suff1 = dt_alias.Rows[index_layer][7].ToString();
                }

                if (dt_alias.Rows[index_layer][10] != DBNull.Value)
                {
                    Suff2 = dt_alias.Rows[index_layer][10].ToString();
                }

                if (dt_alias.Rows[index_layer][13] != DBNull.Value)
                {
                    Suff3 = dt_alias.Rows[index_layer][13].ToString();
                }

                if (dt_alias.Rows[index_layer][16] != DBNull.Value)
                {
                    Suff4 = dt_alias.Rows[index_layer][16].ToString();
                }



                if (dt_alias.Rows[index_layer][6] != DBNull.Value)
                {
                    nf1 = dt_alias.Rows[index_layer][6].ToString();
                }

                if (dt_alias.Rows[index_layer][9] != DBNull.Value)
                {
                    nf2 = dt_alias.Rows[index_layer][9].ToString();
                }

                if (dt_alias.Rows[index_layer][12] != DBNull.Value)
                {
                    nf3 = dt_alias.Rows[index_layer][12].ToString();
                }

                if (dt_alias.Rows[index_layer][15] != DBNull.Value)
                {
                    nf4 = dt_alias.Rows[index_layer][15].ToString();
                }

                string descr1 = "";
                string descr2 = "";
                string descr3 = "";
                string descr4 = "";

                string spatiu = " ";

                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                Autodesk.Gis.Map.ObjectData.Table Tabla1;

                if (first_message_box == false)
                {
                    if (Tables1.IsTableDefined(Ent1.Layer) == false)
                    {
                        MessageBox.Show("Object Data table does not match layer name or\r\nobject data is missing for the object\r\nLayer: " + (char)34 + Ent1.Layer + (char)34);
                        first_message_box = true;
                    }
                }



                if (Tables1.IsTableDefined(Ent1.Layer) == true)
                {
                    Tabla1 = Tables1[Ent1.Layer];


                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                    {
                        if (Records1 != null)
                        {
                            if (Records1.Count > 0)
                            {
                                Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;

                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                {
                                    for (int n = 0; n < Record1.Count; ++n)
                                    {
                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[n];
                                        string Nume_field = Field_def1.Name;
                                        string Valoare1 = Record1[n].StrValue;
                                        if (Nume_field.ToLower() == nf1.ToLower() && Valoare1 != "")
                                        {
                                            descr1 = Pref1 + spatiu + Valoare1 + spatiu + Suff1;
                                        }

                                        if (Nume_field.ToLower() == nf2.ToLower() && Valoare1 != "")
                                        {
                                            descr2 = Pref2 + spatiu + Valoare1 + spatiu + Suff2;
                                        }


                                        if (Nume_field.ToLower() == nf3.ToLower() && Valoare1 != "")
                                        {
                                            descr3 = Pref3 + spatiu + Valoare1 + spatiu + Suff3;
                                        }


                                        if (Nume_field.ToLower() == nf4.ToLower() && Valoare1 != "")
                                        {
                                            descr4 = Pref4 + spatiu + Valoare1 + spatiu + Suff4;
                                        }

                                    }
                                }
                            }
                        }
                    }

                }

                if (descr1 == "" && (Pref1 != "" || Suff1 != ""))
                {
                    descr1 = Pref1 + spatiu + Suff1;
                    if (descr1.Length > 1)
                    {
                        if (descr1.Substring(0, 1) == " ")
                        {
                            descr1 = descr1.Substring(1, descr1.Length - 1);
                        }
                    }
                    if (descr1.Length > 1)
                    {
                        if (descr1.Substring(descr1.Length - 1, 1) == " ")
                        {
                            descr1 = descr1.Substring(0, descr1.Length - 1);
                        }
                    }

                }



                if (descr2 == "" && (Pref2 != "" || Suff2 != ""))
                {
                    descr2 = Pref2 + spatiu + Suff2;
                    if (descr2.Length > 1)
                    {
                        if (descr2.Substring(0, 1) == " ")
                        {
                            descr2 = descr2.Substring(1, descr2.Length - 1);
                        }
                    }
                    if (descr2.Length > 1)
                    {
                        if (descr2.Substring(descr2.Length - 1, 1) == " ")
                        {
                            descr2 = descr2.Substring(0, descr2.Length - 1);
                        }
                    }

                }

                if (descr3 == "" && (Pref3 != "" || Suff3 != ""))
                {
                    descr3 = Pref3 + spatiu + Suff3;
                    if (descr3.Length > 1)
                    {
                        if (descr3.Substring(0, 1) == " ")
                        {
                            descr3 = descr3.Substring(1, descr3.Length - 1);
                        }
                    }
                    if (descr3.Length > 1)
                    {
                        if (descr3.Substring(descr3.Length - 1, 1) == " ")
                        {
                            descr3 = descr3.Substring(0, descr3.Length - 1);
                        }
                    }

                }

                if (descr4 == "" && (Pref4 != "" || Suff4 != ""))
                {
                    descr4 = Pref4 + spatiu + Suff4;
                    if (descr4.Length > 1)
                    {
                        if (descr4.Substring(0, 1) == " ")
                        {
                            descr4 = descr4.Substring(1, descr4.Length - 1);
                        }
                    }
                    if (descr4.Length > 1)
                    {
                        if (descr4.Substring(descr4.Length - 1, 1) == " ")
                        {
                            descr4 = descr4.Substring(0, descr4.Length - 1);
                        }
                    }

                }


                if (descr1 != "" || descr2 != "" || descr3 != "" || descr4 != "")
                {
                    alias1 = descr1 + spatiu + descr2 + spatiu + descr3 + spatiu + descr4;

                }
            }

            do
            {
                alias1 = alias1.Replace("  ", " ");
            }
            while (alias1.Contains("  ") == true);

            return alias1;
        }

        private List<string> get_block_and_atr_from_layer_alias(string layer_name)
        {
            List<string> lista1 = new List<string>();
            if (_AGEN_mainform.dt_layer_alias != null)
            {
                if (_AGEN_mainform.dt_layer_alias.Rows.Count > 0)
                {
                    if (_AGEN_mainform.dt_layer_alias.Columns.Contains("Layer name") == true &&
                       _AGEN_mainform.dt_layer_alias.Columns.Contains("Prof Block Name") == true &&
                       _AGEN_mainform.dt_layer_alias.Columns.Contains("Attrib Sta Prof") == true &&
                       _AGEN_mainform.dt_layer_alias.Columns.Contains("Attrib Desc Prof") == true)
                    {
                        for (int i = 0; i < _AGEN_mainform.dt_layer_alias.Rows.Count; ++i)
                        {
                            string block = "";
                            string sta_atr = "";
                            string desc_atr = "";
                            string ln = "";
                            if (_AGEN_mainform.dt_layer_alias.Rows[i]["Layer name"] != DBNull.Value)
                            {
                                ln = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i]["Layer name"]);
                                if (ln == layer_name)
                                {
                                    if (_AGEN_mainform.dt_layer_alias.Rows[i]["Prof Block Name"] != DBNull.Value)
                                    {
                                        block = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i]["Prof Block Name"]);
                                        if (_AGEN_mainform.dt_layer_alias.Rows[i]["Attrib Sta Prof"] != DBNull.Value)
                                        {
                                            sta_atr = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i]["Attrib Sta Prof"]);
                                        }
                                        if (_AGEN_mainform.dt_layer_alias.Rows[i]["Attrib Desc Prof"] != DBNull.Value)
                                        {
                                            desc_atr = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i]["Attrib Desc Prof"]);
                                        }
                                    }
                                    lista1.Add(block);
                                    lista1.Add(sta_atr);
                                    lista1.Add(desc_atr);
                                    i = _AGEN_mainform.dt_layer_alias.Rows.Count;
                                }
                            }
                        }
                    }
                }
            }
            return lista1;
        }


        public void Build_3d_from_2d_poly(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord, System.Data.DataTable dt1, String layer_name)
        {
            if (_AGEN_mainform.Project_type == "2D")
            {
                Polyline poly2d = new Polyline();
                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    double x = 0;
                    double y = 0;
                    if (dt1.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value)
                    {
                        x = (double)dt1.Rows[i][_AGEN_mainform.Col_x];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no X value for centerline in row " + (i).ToString());
                        return;
                    }
                    if (dt1.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                    {
                        y = (double)dt1.Rows[i][_AGEN_mainform.Col_y];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                        return;
                    }
                    poly2d.AddVertexAt(i, new Point2d(x, y), 0, 0, 0);
                }

                poly2d.Layer = layer_name;
                BTrecord.AppendEntity(poly2d);
                Trans1.AddNewlyCreatedDBObject(poly2d, true);

                Polyline3d Poly3D = new Polyline3d();
                Poly3D.SetDatabaseDefaults();
                Poly3D.Layer = layer_name;
                BTrecord.AppendEntity(Poly3D);
                Trans1.AddNewlyCreatedDBObject(Poly3D, true);

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    double x = 0;
                    double y = 0;
                    double z = 0;
                    if (dt1.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value)
                    {
                        x = (double)dt1.Rows[i][_AGEN_mainform.Col_x];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no X value for centerline in row " + (i).ToString());
                        return;
                    }
                    if (dt1.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                    {
                        y = (double)dt1.Rows[i][_AGEN_mainform.Col_y];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                        return;
                    }
                    PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                    Poly3D.AppendVertex(Vertex_new);
                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                }
            }
            else if (_AGEN_mainform.Project_type == "3D")
            {
                Polyline3d Poly3D = new Polyline3d();
                Poly3D.SetDatabaseDefaults();
                Poly3D.Layer = layer_name;
                BTrecord.AppendEntity(Poly3D);
                Trans1.AddNewlyCreatedDBObject(Poly3D, true);

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    double x = 0;
                    double y = 0;
                    double z = 0;

                    if (dt1.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value)
                    {
                        x = (double)dt1.Rows[i][_AGEN_mainform.Col_x];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no X value for centerline in row " + (i).ToString());
                        return;
                    }
                    if (dt1.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                    {
                        y = (double)dt1.Rows[i][_AGEN_mainform.Col_y];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                        return;
                    }
                    if (dt1.Rows[i][_AGEN_mainform.Col_z] != DBNull.Value)
                    {
                        z = (double)dt1.Rows[i][_AGEN_mainform.Col_z];
                    }
                    else
                    {
                        set_enable_true();
                        MessageBox.Show("no Z value for centerline in row " + (i).ToString());
                        return;
                    }

                    PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                    Poly3D.AppendVertex(Vertex_new);
                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                }


                Polyline poly2d = Functions.Build_2dpoly_from_3d(Poly3D);
                poly2d.Layer = layer_name;
                BTrecord.AppendEntity(poly2d);
                Trans1.AddNewlyCreatedDBObject(poly2d, true);
            }
        }

        public void Populate_crossing_file(string File1)
        {
            try
            {
                if (System.IO.File.Exists(File1) == false)
                {
                    create_crossing_file(File1);
                }
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                bool excel_is_opened = false;
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == File1.ToLower())
                        {
                            Workbook1 = Workbook2;
                            W1 = Workbook1.Worksheets[1];
                            excel_is_opened = true;
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
                    W1 = Workbook1.Worksheets[1];
                }

                try
                {
                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";


                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.Data_Table_crossings, _AGEN_mainform.Start_row_crossing, "General");
                    Functions.Create_header_crossing_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                    Workbook1.Save();

                    if (excel_is_opened == false)
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



        public System.Data.DataTable Load_existing_crossing(string File1, string sheetname = "XYZAa")
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the crosings data file does not exist");
                return null;
            }
            System.Data.DataTable dt2 = new System.Data.DataTable();
            bool excel_is_opened = false;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            try
            {
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == File1.ToLower())
                        {
                            Workbook1 = Workbook2;
                            bool exista = false;
                            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in Excel1.Worksheets)
                            {
                                if (sheet.Name == sheetname)
                                {
                                    exista = true;
                                }
                            }
                            if (exista == false)
                            {
                                sheetname = "XYZAa";
                            }
                            if (sheetname == "XYZAa")
                            {
                                W1 = Workbook1.Worksheets[1];
                            }
                            else
                            {
                                W1 = Workbook1.Worksheets[sheetname];
                            }
                            excel_is_opened = true;
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
                    bool exista = false;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in Excel1.Worksheets)
                    {
                        if (sheet.Name == sheetname)
                        {
                            exista = true;
                        }
                    }
                    if (exista == false)
                    {
                        sheetname = "XYZAa";
                    }
                    if (sheetname == "XYZAa")
                    {
                        W1 = Workbook1.Worksheets[1];
                    }
                    else
                    {
                        W1 = Workbook1.Worksheets[sheetname];
                    }
                    if (sheetname == "XYZAa")
                    {
                        W1 = Workbook1.Worksheets[1];
                    }
                    else
                    {
                        W1 = Workbook1.Worksheets[sheetname];
                    }
                }

                try
                {
                    dt2 = Functions.Build_Data_table_crossings_from_excel(W1, _AGEN_mainform.Start_row_crossing + 1);
                    if (excel_is_opened == false)
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
            return dt2;
        }


        private void button_scan_with_offsets_Click(object sender, EventArgs e)
        {
            first_message_box = false;
            if (radioButton_use_one_offset.Checked == false)
            {
                MessageBox.Show("this feature has not been enabled\r\nsoon it will be available");
                return;
            }

            if (checkBox_scan_for_points.Checked == false && checkBox_scan_for_blocks.Checked == false)
            {
                MessageBox.Show("Please select scan points and/or blocks");
                return;
            }

            if (Functions.IsNumeric(textBox_offset_value.Text) == false)
            {
                MessageBox.Show("Please specify scanning offset distance");
                return;
            }

            double Offset1 = Math.Abs(Convert.ToDouble(textBox_offset_value.Text));

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            

            if (Functions.Get_if_workbook_is_open_in_Excel("crossing.xlsx") == true)
            {
                MessageBox.Show("Please close the crossing file");
                return;
            }


            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            if (_AGEN_mainform.dt_layer_alias == null)
            {
                MessageBox.Show("No layer alias table loaded");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_viewport_settings.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();

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


                _AGEN_mainform.tpage_layer_alias.Show();

                return;
            }


            if (_AGEN_mainform.dt_layer_alias.Rows.Count == 0)
            {
                MessageBox.Show("The layer alias table contains no data");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_viewport_settings.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();

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


                _AGEN_mainform.tpage_layer_alias.Show();
                return;
            }

            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == true)
            {
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                set_enable_true();
                MessageBox.Show("the centerline data file does not exist");
                _AGEN_mainform.dt_station_equation = null;
                return;
            }

            string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

            if (System.IO.File.Exists(fisier_cs) == false)
            {
                create_crossing_file(fisier_cs);
            }

            string ProjFolder_for_layer_alias = _AGEN_mainform.tpage_setup.Get_project_database_folder_without_segment();

            string fisier_alias = ProjFolder_for_layer_alias + _AGEN_mainform.layer_alias_excel_name;
            if (System.IO.File.Exists(fisier_alias) == false)
            {
                set_enable_true();
                MessageBox.Show("the file alias not found");
                return;
            }


            if (_AGEN_mainform.dt_centerline == null)
            {
                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
            }


            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("no centerline data!");
                return;
            }

            List<string> Lista1 = new List<string>();
            List<string> Lista2 = new List<string>();
            for (int i = 0; i < _AGEN_mainform.dt_layer_alias.Rows.Count; ++i)
            {
                if (_AGEN_mainform.dt_layer_alias.Rows[i][0] != DBNull.Value)
                {
                    Lista1.Add(_AGEN_mainform.dt_layer_alias.Rows[i][0].ToString().ToLower());
                }
                else
                {
                    Lista1.Add("danpopescuistheking");
                }
                if (_AGEN_mainform.dt_layer_alias.Rows[i][20] != DBNull.Value)
                {
                    string Display = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i][20]);
                    if (Display.ToUpper() == "YES") Lista2.Add("YES");
                    if (Display.ToUpper() == "NO") Lista2.Add("NO");

                }
                else
                {
                    Lista2.Add("NO");
                }

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
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                        _AGEN_mainform.Data_Table_crossings = Functions.Creaza_crossing_datatable_structure();

                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        double poly_length = poly2d.Length;
                        Polyline3d poly3d = null;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            poly_length = poly3d.Length;
                        }

                        System.Data.DataTable dt_offset_layer = new System.Data.DataTable();
                        dt_offset_layer.Columns.Add("layer", typeof(string));
                        dt_offset_layer.Columns.Add("offset", typeof(double));

                        foreach (ObjectId odid in BTrecord)
                        {
                            DBPoint DBPoint1 = Trans1.GetObject(odid, OpenMode.ForRead) as DBPoint;
                            BlockReference Block1 = Trans1.GetObject(odid, OpenMode.ForRead) as BlockReference;
                            Point3d Position1 = new Point3d(0, 0, -123.123123);
                            Entity Ent1 = null;


                            if (checkBox_scan_for_points.Checked == true)
                            {
                                if (DBPoint1 != null)
                                {
                                    Position1 = DBPoint1.Position;
                                    Ent1 = DBPoint1 as Entity;
                                }
                            }


                            if (checkBox_scan_for_blocks.Checked == true)
                            {
                                if (Block1 != null)
                                {
                                    Position1 = Block1.Position;
                                    Ent1 = Block1 as Entity;
                                }
                            }


                            if (Ent1 != null)
                            {


                                if (Lista1.Contains(Ent1.Layer.ToLower()) == true)
                                {
                                    int i = Lista1.IndexOf(Ent1.Layer.ToLower());

                                    Point3d pt1 = poly2d.GetClosestPointTo(new Point3d(Position1.X, Position1.Y, 0), Vector3d.ZAxis, false);

                                    double Dist1 = pt1.DistanceTo(new Point3d(Position1.X, Position1.Y, 0));



                                    if (Dist1 <= Offset1)
                                    {

                                        try
                                        {

                                            dt_offset_layer.Rows.Add();
                                            dt_offset_layer.Rows[dt_offset_layer.Rows.Count - 1][0] = Ent1.Layer;

                                            dt_offset_layer.Rows[dt_offset_layer.Rows.Count - 1][1] = Offset1;


                                            _AGEN_mainform.Data_Table_crossings.Rows.Add();

                                            if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                            {
                                                double sta2d = Math.Round(poly2d.GetDistAtPoint(pt1), 3);
                                                if (sta2d >= poly2d.Length)
                                                {
                                                    sta2d = sta2d - 0.001;
                                                }

                                                _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][1] = sta2d;
                                                if (_AGEN_mainform.dt_station_equation != null)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                    {
                                                        _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][3] =
                                                            Math.Round(Functions.Station_equation_of(sta2d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                }
                                            }

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                double sta2d = poly2d.GetDistAtPoint(pt1);
                                                double param1 = poly2d.GetParameterAtDistance(sta2d);
                                                double sta3d = Math.Round(poly3d.GetDistanceAtParameter(param1), 3);

                                                if (sta3d >= poly3d.Length)
                                                {
                                                    sta3d = sta3d - 0.001;
                                                }

                                                _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][2] = sta3d;
                                                if (_AGEN_mainform.dt_station_equation != null)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                    {
                                                        _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][3] =
                                                            Math.Round(Functions.Station_equation_of(sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                }
                                            }

                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][7] = Position1.X;
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][8] = Position1.Y;
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][9] = Position1.Z;

                                            string alias1 = get_description_for_obj(_AGEN_mainform.dt_layer_alias, i, Ent1);

                                            string spatiu = " ";
                                            string suffix = "'";

                                            if (_AGEN_mainform.units_of_measurement == "m") suffix = "";

                                            string left_right = Functions.Angle_left_right(poly2d, new Point3d(Position1.X, Position1.Y, 0));

                                            alias1 = alias1 + spatiu + Functions.Get_String_Rounded(Dist1, 0) + suffix + spatiu + left_right;
                                            alias1 = alias1.Replace("  ", " ");

                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][10] = Dist1;
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][11] = left_right;
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][6] = alias1;
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][4] = Ent1.GetType().ToString().Replace("Autodesk.AutoCAD.DatabaseServices.", "");
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][5] = Ent1.Layer;
                                            _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispXing"] = Lista2[i];
                                            List<string> lista_bl_atr = get_block_and_atr_from_layer_alias(Ent1.Layer);
                                            if (lista_bl_atr.Count == 3)
                                            {
                                                if (lista_bl_atr[0] != "")
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["Prof Block Name"] = lista_bl_atr[0];
                                                    if (lista_bl_atr[1] != "") _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["Attrib Sta Prof"] = lista_bl_atr[1];
                                                    if (lista_bl_atr[2] != "") _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["Attrib Desc Prof"] = lista_bl_atr[2];
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispProf"] = "YES";
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispProf"] = "NO";
                                                }
                                            }
                                            else
                                            {
                                                _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1]["DispProf"] = "NO";
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            MessageBox.Show(ex.Message);
                                        }
                                    }
                                }

                            }
                        }

                        if (radioButton_append_to_crossing_table.Checked == true)
                        {
                            System.Data.DataTable table1 = Load_existing_crossing(fisier_cs);

                            if (table1.Rows.Count > 0)
                            {
                                for (int i = 0; i < table1.Rows.Count; ++i)
                                {
                                    if (table1.Rows[i][1] != DBNull.Value || table1.Rows[i][2] != DBNull.Value)
                                    {
                                        _AGEN_mainform.Data_Table_crossings.Rows.Add();

                                        for (int j = 0; j < 17; ++j)
                                        {
                                            if (table1.Rows[i][j] != DBNull.Value)
                                            {
                                                string val1 = Convert.ToString(table1.Rows[i][j]);
                                                _AGEN_mainform.Data_Table_crossings.Rows[_AGEN_mainform.Data_Table_crossings.Rows.Count - 1][j] = Convert.ToString(val1);
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        if (_AGEN_mainform.Data_Table_crossings != null)
                        {
                            if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                            {
                                string Display_in_crossing_band = "NO";

                                for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                {
                                    double St1 = 0;
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"] == DBNull.Value) _AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"] = Display_in_crossing_band;

                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value || _AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value)
                                    {
                                        double div1 = 10;
                                        if (_AGEN_mainform.round1 == 1) div1 = 100;
                                        if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                        if (_AGEN_mainform.round1 == 3) div1 = 10000;

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value)
                                        {
                                            St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][1]), _AGEN_mainform.round1);
                                            if (St1 >= poly2d.Length) St1 = Math.Floor(Math.Round(poly2d.Length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][1] = St1;
                                        }

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value)
                                        {
                                            St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][2]), _AGEN_mainform.round1);
                                            if (St1 >= poly_length) St1 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][2] = St1;
                                        }

                                        if (_AGEN_mainform.dt_station_equation != null)
                                        {
                                            if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                            {
                                                _AGEN_mainform.Data_Table_crossings.Rows[i][3] =
                                                    Math.Round(Functions.Station_equation_of(St1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                            }
                                            else
                                            {
                                                _AGEN_mainform.Data_Table_crossings.Rows[i][3] = DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][3] = DBNull.Value;
                                        }
                                    }
                                }
                            }
                        }

                        Functions.create_backup(fisier_cs);
                        Populate_crossing_file(fisier_cs);

                        if (dt_offset_layer.Rows.Count > 0)
                        {
                            for (int i = 0; i < _AGEN_mainform.dt_layer_alias.Rows.Count; ++i)
                            {
                                string Layer_name = _AGEN_mainform.dt_layer_alias.Rows[i][0].ToString();
                                for (int j = 0; j < dt_offset_layer.Rows.Count; ++j)
                                {
                                    if (Layer_name.ToLower() == dt_offset_layer.Rows[j][0].ToString().ToLower())
                                    {
                                        _AGEN_mainform.dt_layer_alias.Rows[i][4] = dt_offset_layer.Rows[j][1];
                                        j = dt_offset_layer.Rows.Count;
                                    }
                                }

                            }
                        }
                        _AGEN_mainform.tpage_layer_alias.update_alias_file_from_crossing(_AGEN_mainform.dt_layer_alias, fisier_alias);



                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");

            dataGridView_xing.DataSource = _AGEN_mainform.Data_Table_crossings;
            dataGridView_xing.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_xing.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_xing.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_xing.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_xing.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_xing.EnableHeadersVisualStyles = false;

            set_enable_true();

            radioButton_append_to_crossing_table.Checked = true;
            MessageBox.Show("done");
        }

        private void radioButton_use_one_offset_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_use_one_offset.Checked == true)
            {
                hide_individual_offsets();
            }
            else
            {
                panel_individual_offsets.Visible = true;
                textBox_offset_value.Visible = false;
            }
        }

        public void hide_individual_offsets()
        {
            panel_individual_offsets.Visible = false;
            textBox_offset_value.Visible = true;
        }

        public void set_textBox_offset_value(double offset1)
        {
            textBox_offset_value.Text = offset1.ToString();
        }






        private void button_show_scan_and_draw_crossings_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();
            _AGEN_mainform.tpage_layer_alias.Hide();
            _AGEN_mainform.tpage_crossing_scan.Hide();

            _AGEN_mainform.tpage_profilescan.Hide();
            _AGEN_mainform.tpage_profdraw.Hide();
            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_crossing_draw.Show();
        }

        private void label_crossing_settings_Click(object sender, EventArgs e)
        {

            if (panel_individual_offsets.Visible == false)
            {

                radioButton_use_multiple_offsets.Visible = true;
                panel_individual_offsets.Visible = true;
                button_transfer_to_excel_crossing_band.Visible = true;
            }
            else
            {

                radioButton_use_multiple_offsets.Visible = false;
                panel_individual_offsets.Visible = false;
                button_transfer_to_excel_crossing_band.Visible = false;
            }

        }

        private void button_open_excel_crossing_Click(object sender, EventArgs e)
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

                    string fisier_crossing = ProjF + _AGEN_mainform.crossing_excel_name;

                    if (System.IO.File.Exists(fisier_crossing) == false)
                    {
                        set_enable_true();
                        MessageBox.Show("the crossing data file does not exist");
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
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fisier_crossing);
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

        private void button_transfer_to_excel_crossing_band_Click(object sender, EventArgs e)
        {
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Editor1.SetImpliedSelection(Empty_array);
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        System.Data.DataTable dt1 = Functions.Creaza_crossing_datatable_structure();

                        Ag.WindowState = FormWindowState.Minimized;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect existing crossing band text:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            Ag.WindowState = FormWindowState.Normal;
                            Editor1.SetImpliedSelection(Empty_array);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            DBText text1 = Obj1.ObjectId.GetObject(OpenMode.ForRead) as DBText;
                            MText mtext1 = Obj1.ObjectId.GetObject(OpenMode.ForRead) as MText;
                            string continut = "";
                            if (text1 != null)
                            {
                                continut = text1.TextString;
                            }

                            if (mtext1 != null)
                            {
                                continut = mtext1.Text;
                            }

                            if (continut != "")
                            {
                                if (continut.Substring(0, 1) == " ")
                                {
                                    continut = continut.Substring(1, continut.Length - 1);
                                }

                                string station = Functions.extrage_STATION_din_text_de_la_inceputul_textului(continut);
                                if (Functions.IsNumeric(station.Replace("+", "")) == true)
                                {
                                    double sta1 = Convert.ToDouble(station.Replace("+", ""));
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1]["EqSta"] = sta1;
                                    string descr = continut.Replace(station, "");
                                    do
                                    {
                                        if (descr.Substring(0, 1) == " ")
                                        {
                                            descr = descr.Substring(1, descr.Length - 1);
                                        }
                                    } while (descr.Substring(0, 1) == " ");

                                    do
                                    {
                                        if (descr.Substring(descr.Length - 1, 1) == " ")
                                        {
                                            descr = descr.Substring(0, descr.Length - 1);
                                        }
                                    } while (descr.Substring(descr.Length - 1, 1) == " ");

                                    dt1.Rows[dt1.Rows.Count - 1]["Desc"] = descr;

                                }

                            }
                        }

                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Ag.WindowState = FormWindowState.Normal;

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();



            MessageBox.Show("done");
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

        private void button_save_crossing_table_to_excel_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {

                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                    {
                        ProjFolder = ProjFolder + "\\";
                    }
                }
                else
                {
                    set_enable_true();
                    MessageBox.Show("the project folder does not exist");
                    return;
                }

                string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline data file does not exist");
                    return;
                }

                string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

                Functions.create_backup(fisier_cs);
                Populate_crossing_file(fisier_cs);
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_load_crossing_table_from_excel_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                    {
                        ProjFolder = ProjFolder + "\\";
                    }
                }
                else
                {
                    set_enable_true();
                    MessageBox.Show("the project folder does not exist");
                    return;
                }

                string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline data file does not exist");
                    return;
                }

                string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

                _AGEN_mainform.Data_Table_crossings = Load_existing_crossing(fisier_cs);

                dataGridView_xing.DataSource = _AGEN_mainform.Data_Table_crossings;
                dataGridView_xing.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_xing.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_xing.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_xing.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_xing.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_xing.EnableHeadersVisualStyles = false;
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }



            set_enable_true();


        }

        private void button_calc_station_from_point_Click(object sender, EventArgs e)
        {

            string Col_2d = "2DSta";
            string Col_3d = "3DSta";
            string Col_eq = "EqSta";
            string Col_MMid = "MMID";

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {

                        string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

                        if (System.IO.File.Exists(fisier_cs) == false)
                        {
                            MessageBox.Show("no crossing.xls found", "AGEN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        _AGEN_mainform.Data_Table_crossings = Load_existing_crossing(fisier_cs);
                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                        if (_AGEN_mainform.Data_Table_crossings == null || _AGEN_mainform.Data_Table_crossings.Rows.Count == 0)
                        {
                            MessageBox.Show("no data in crossing.xls found", "AGEN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        try
                        {
                            set_enable_false();
                            using (DocumentLock lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                                    Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                    Polyline3d poly3d = null;
                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                    }

                                    set_enable_false();





                                    for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                    {
                                        double x = -1.234;
                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x])) == true)
                                        {
                                            x = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x]);
                                        }


                                        double y = -1.234;
                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y])) == true)
                                        {
                                            y = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y]);
                                        }





                                        if (x != -1.234 && y != -1.234)
                                        {
                                            Point3d pt2d = poly2d.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                            double calc_sta = poly2d.GetDistAtPoint(pt2d);


                                            double offset1 = Math.Pow(Math.Pow(x - pt2d.X, 2) + Math.Pow(y - pt2d.Y, 2), 0.5);

                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {

                                                if (_AGEN_mainform.Project_type == "2D")
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i][Col_2d] = Math.Round(calc_sta, _AGEN_mainform.round1);
                                                }
                                                else
                                                {
                                                    double param1 = poly2d.GetParameterAtPoint(pt2d);
                                                    calc_sta = poly3d.GetDistanceAtParameter(param1);
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i][Col_3d] = Math.Round(calc_sta, _AGEN_mainform.round1);
                                                }

                                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                {
                                                    double steq = Functions.Station_equation_ofV2(Math.Round(calc_sta, _AGEN_mainform.round1), _AGEN_mainform.dt_station_equation);
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i][Col_eq] = Math.Round(steq, _AGEN_mainform.round1);
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i][Col_eq] = DBNull.Value;
                                                }

                                                if (offset1 > 1)
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i][Col_MMid] = "offset = " + Math.Round(offset1, 1).ToString();
                                                }
                                            }
                                        }
                                    }
                                    if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                                    Trans1.Commit();

                                    Functions.create_backup(fisier_cs);
                                    Populate_crossing_file(fisier_cs);
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_create_empty_crossing_Click(object sender, EventArgs e)
        {


            if (Functions.Get_if_workbook_is_open_in_Excel("crossing.xlsx") == true)
            {
                MessageBox.Show("Please close the crossing file");
                return;
            }
            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }
            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == true)
            {
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }


            string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

            if (System.IO.File.Exists(fisier_cs) == false)
            {
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

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                    Workbook1 = Excel1.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                    W1.Name = "Crossing";
                    try
                    {
                        string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                        if (segment1 == "not defined") segment1 = "";
                        Functions.Create_header_crossing_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
                        Workbook1.SaveAs(fisier_cs);
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
            }



        }

        private void dataGridView_xing_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
